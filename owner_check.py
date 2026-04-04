"""
売上明細 vs 請求CSV 照合チェックツール（全校舎対応）

config.yaml に校舎を追加するだけで、新校舎の照合チェックに対応。
売上申請（Excel）に記載された金額が、請求データ（CSV）と一致しているか検証する。
"""

import csv
import re
import sys
import os
import glob
from dataclasses import dataclass, field
from pathlib import Path
from datetime import datetime

try:
    import yaml
except ImportError:
    print("PyYAML が必要です: pip install pyyaml")
    sys.exit(1)

try:
    import openpyxl
except ImportError:
    print("openpyxl が必要です: pip install openpyxl")
    sys.exit(1)


# ====================================================================
# データクラス
# ====================================================================

@dataclass(frozen=True)
class SchoolConfig:
    name: str
    excel_dir: str
    excel_pattern: str
    sheet_index: int = 2


@dataclass(frozen=True)
class MonthPair:
    year: int
    month: int
    csv_path: str
    excel_path: str

    @property
    def label(self) -> str:
        return f"{self.year}年{self.month}月"


@dataclass
class CheckResult:
    school: str
    month_label: str
    result_type: str      # "TOTAL_DIFF", "TOTAL_MATCH", "NOT_IN_CSV"
    sid: str
    name: str
    row: int
    diffs: list = field(default_factory=list)
    excel_total: float = 0.0
    csv_total: float = 0.0


# ====================================================================
# ブランド名 → Excel列 のマッピング
# Excel列構造:
#   D=設備費, E=授業料(学習塾), F=月会費(学習塾), G=テスト対策/講習会費, H=模試代, I=講習会費
#   J=授業料(プログラミング), K=月会費(プログラミング), L=講習会費(プログラミング)
#   M=授業料(アン), N=月会費(アン)
#   O=授業料(そろばん), P=月会費(そろばん), Q=講習会費(そろばん)
#   R=授業料(筆っこ), S=月会費(筆っこ), T=講習会費(筆っこ)
#   U=授業料(将棋), V=月会費(将棋), W=講習会費(将棋)
#   X=入会金, Y=その他, Z=売上合計
# ====================================================================

JUKU_BRANDS = {
    "アン進学ジム【小学部】",
    "アン進学ジム【中学部】",
    "アン進学ジム【高校部】",
    "SEEDS 小学部",
    "SEEDS 中学部",
    "SEEDS 高校部",
    "SEEDS Lepton",
    "星煌学院本科",
    "千種校　中学受験コース",
    "小学生選抜クラス",
    "小学生学力向上コース",
    "中学生選抜コース",
    "志望校合格コース",
    "英検対策",
    "マンツーマン学院",
    "メプレス",
    "須田塾 高校部",
}

BRAND_COLUMN_MAP = {
    "アンイングリッシュクラブ": ("M", "N"),
    "アンさんこくキッズ": ("M", "N"),
    "アンそろばんクラブ": ("O", "P"),
    "アンそろばんクラブ【選択講習会】": ("O", "P"),
    "アンプログラミングクラブ": ("J", "K"),
    "アン美文字クラブ": ("R", "S"),
    "アン将棋クラブ": ("U", "V"),
}

COL_DISPLAY = {
    "D": "設備費",
    "E": "授業料(学習塾)",
    "F": "月会費(学習塾)",
    "G": "テスト対策/講習会費",
    "H": "模試代",
    "I": "講習会費",
    "J": "授業料(プログラミング)",
    "K": "月会費(プログラミング)",
    "L": "講習会費(プログラミング)",
    "M": "授業料(アン)",
    "N": "月会費(アン)",
    "O": "授業料(そろばん)",
    "P": "月会費(そろばん)",
    "Q": "講習会費(そろばん)",
    "R": "授業料(筆っこ)",
    "S": "月会費(筆っこ)",
    "T": "講習会費(筆っこ)",
    "U": "授業料(将棋)",
    "V": "月会費(将棋)",
    "W": "講習会費(将棋)",
    "X": "入会金",
    "Y": "その他",
    "Z": "売上合計",
}


# ====================================================================
# 設定読み込み
# ====================================================================

def load_config(config_path: str = None) -> dict:
    if config_path is None:
        config_path = str(Path(__file__).parent / "config.yaml")
    with open(config_path, encoding="utf-8") as f:
        return yaml.safe_load(f)


# ====================================================================
# ファイル自動検出
# ====================================================================

def discover_csv_files(csv_base_dir: str) -> dict[tuple[int, int], str]:
    """
    請求CSVを自動検出。ディレクトリ名から請求対象月を推定。
    CSVディレクトリの日付 +1ヶ月 = 請求対象月
    返り値: {(year, month): csv_path}
    """
    result = {}
    base = Path(csv_base_dir)

    for d in sorted(base.iterdir()):
        if not d.is_dir():
            continue
        # ディレクトリ名から日付を抽出 (YYYYMMDD_...)
        m = re.match(r"(\d{4})(\d{2})\d{2}", d.name)
        if not m:
            continue

        # CSVファイルを探す
        csvs = list(d.glob("AC_5_*UTF8*.csv"))
        if not csvs:
            continue

        dir_year = int(m.group(1))
        dir_month = int(m.group(2))

        # 請求対象月 = ディレクトリの月 + 1
        billing_month = dir_month + 1
        billing_year = dir_year
        if billing_month > 12:
            billing_month = 1
            billing_year += 1

        # 最新のCSVを使用（同じディレクトリに複数ある場合）
        csv_path = str(sorted(csvs, key=lambda p: p.stat().st_mtime)[-1])
        result[(billing_year, billing_month)] = csv_path

    return result


def discover_excel_files(school: SchoolConfig) -> dict[tuple[int, int], str]:
    """
    校舎のExcelファイルを自動検出。ファイル名から年月を抽出。
    返り値: {(year, month): excel_path}
    """
    result = {}
    excel_dir = Path(school.excel_dir)

    if not excel_dir.exists():
        return result

    for f in excel_dir.iterdir():
        if not f.is_file():
            continue
        if not (f.suffix == ".xlsm" or f.suffix == ".xlsx"):
            continue

        # ファイル名から (YYYY.MM) を抽出
        m = re.search(r"\((\d{4})\.(\d{2})\)", f.name)
        if not m:
            continue

        year = int(m.group(1))
        month = int(m.group(2))
        result[(year, month)] = str(f)

    return result


def match_months(
    csv_files: dict[tuple[int, int], str],
    excel_files: dict[tuple[int, int], str],
) -> list[MonthPair]:
    """CSV と Excel を月でマッチング"""
    pairs = []
    for ym in sorted(set(csv_files.keys()) & set(excel_files.keys())):
        pairs.append(MonthPair(
            year=ym[0],
            month=ym[1],
            csv_path=csv_files[ym],
            excel_path=excel_files[ym],
        ))
    return pairs


# ====================================================================
# データ読み込み・集計（コアロジック）
# ====================================================================

def parse_number(val) -> float:
    if val is None or val == "" or val == 0:
        return 0.0
    try:
        return float(str(val).replace(",", ""))
    except (ValueError, TypeError):
        return 0.0


def read_billing_csv(csv_path: str) -> dict[str, list[tuple[str, str, float]]]:
    """請求CSVを読み込み → {生徒ID: [(ブランド名, カテ��リ名, 月額料金), ...]}"""
    result = {}
    with open(csv_path, encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        for row in reader:
            sid = row["生徒ID"].strip()
            brand = row["ブランド名"].strip()
            category = row["請求カテゴリ名"].strip()
            amount = parse_number(row["月額料金"])
            if sid not in result:
                result[sid] = []
            result[sid].append((brand, category, amount))
    return result


def aggregate_csv_for_student(
    entries: list[tuple[str, str, float]],
) -> dict[str, float]:
    """CSV明細をExcel列に対応する形で集計 → {列レター: 金額}"""
    col_totals: dict[str, float] = {}

    for brand, category, amount in entries:
        if amount == 0:
            continue

        target_col = _map_to_column(brand, category)

        if target_col:
            col_totals[target_col] = col_totals.get(target_col, 0) + amount

    return col_totals


def _map_to_column(brand: str, category: str) -> str | None:
    """ブランド名＋カテゴリ名 → Excel列レターを決定"""
    if category == "設備費":
        return "D"
    if category == "入会金":
        return "X"
    if category == "模試代":
        return "H"
    if category in ("0", "1"):
        return "Y"
    if category in ("入会時教材費", "入会時授業料1", "入会時授業料2",
                     "入会時授業料3", "入会時授業料A", "入会時月会費",
                     "入会時設備費"):
        return "Y"
    if category in ("家族割", "割引", "過不足金", "諸経費", "家���",
                     "総合指導管理費"):
        return "Y"

    if brand in BRAND_COLUMN_MAP:
        jugyou_col, getsukai_col = BRAND_COLUMN_MAP[brand]
        if category in ("授業料", "追加授業料"):
            return jugyou_col
        if category == "月会費":
            return getsukai_col
        if category in ("講習��費", "必須講座", "必須講習会", "テスト対策"):
            if brand.startswith("アンそろばん"):
                return "Q"
            if brand == "アンプログラミングクラブ":
                return "L"
            if brand == "アン美文字クラブ":
                return "T"
            if brand == "アン将棋クラブ":
                return "W"
            return "I"
        return "Y"

    if brand in JUKU_BRANDS:
        if category in ("授業料", "追加授業料"):
            return "E"
        if category == "月会費":
            return "F"
        if category in ("講習会費", "必須講座", "必須講習会", "テスト対策"):
            return "I"
        return "Y"

    if "そろばん検定" in brand:
        return "Y"

    return "Y"


def read_excel_sales(
    excel_path: str,
    sheet_index: int = 2,
) -> dict[str, dict]:
    """
    売上明細Excelを読み込む。
    返り値: {生徒ID: {"name": 生徒名, "cols": {列レター: 金額}, "row": 行番号}}
    """
    wb = openpyxl.load_workbook(excel_path, read_only=True, data_only=True)

    if sheet_index >= len(wb.sheetnames):
        print(f"  警告: シートindex {sheet_index} が存在しません: {excel_path}")
        wb.close()
        return {}

    ws = wb[wb.sheetnames[sheet_index]]

    result = {}
    for row in ws.iter_rows(min_row=5, max_row=ws.max_row, values_only=False):
        a_val = row[0].value
        if a_val is not None and isinstance(a_val, str) and "計" in a_val:
            break

        sid = row[1].value
        if sid is None:
            continue

        sid = str(int(sid)) if isinstance(sid, (int, float)) else str(sid)
        name = row[2].value or ""

        cols = {}
        for cell in row:
            if not hasattr(cell, "column_letter"):
                continue
            letter = cell.column_letter
            if letter in ("A", "B", "C"):
                continue
            val = parse_number(cell.value)
            if val != 0:
                cols[letter] = val

        result[sid] = {"name": str(name), "cols": cols, "row": row[0].row}

    wb.close()
    return result


# ====================================================================
# 照合チェック
# ====================================================================

def compare_student(
    excel_cols: dict[str, float],
    csv_cols: dict[str, float],
) -> list[tuple[str, float, float, float]]:
    """1人の生徒をExcel列ごとにCSVと比較 → [(列, Excel値, CSV値, 差額)]"""
    diffs = []
    all_cols = sorted(set(list(excel_cols.keys()) + list(csv_cols.keys())))

    for col in all_cols:
        if col == "Z":
            continue
        excel_val = excel_cols.get(col, 0)
        csv_val = csv_cols.get(col, 0)
        if isinstance(excel_val, str):
            continue
        if abs(excel_val - csv_val) >= 1:
            diffs.append((col, excel_val, csv_val, excel_val - csv_val))

    return diffs


def run_check(
    school_name: str,
    month_label: str,
    csv_path: str,
    excel_path: str,
    sheet_index: int = 2,
) -> list[CheckResult]:
    """1校舎・1���月分のチェックを実行"""
    print(f"\n{'='*80}")
    print(f"  [{school_name}] {month_label} 照合チェック")
    print(f"  CSV:   {os.path.basename(csv_path)}")
    print(f"  Excel: {os.path.basename(excel_path)}")
    print(f"{'='*80}")

    billing = read_billing_csv(csv_path)
    sales = read_excel_sales(excel_path, sheet_index)

    ok_count = 0
    col_only_count = 0
    total_diff_count = 0
    not_in_csv_count = 0
    results: list[CheckResult] = []

    for sid, sdata in sorted(sales.items(), key=lambda x: x[1]["row"]):
        name = sdata["name"]
        excel_cols = sdata["cols"]
        row_num = sdata["row"]

        if sid not in billing:
            total = excel_cols.get("Z", 0)
            if total != 0 and not isinstance(total, str):
                not_in_csv_count += 1
                results.append(CheckResult(
                    school=school_name,
                    month_label=month_label,
                    result_type="NOT_IN_CSV",
                    sid=sid,
                    name=name,
                    row=row_num,
                    excel_total=total,
                ))
            continue

        csv_agg = aggregate_csv_for_student(billing[sid])
        diffs = compare_student(excel_cols, csv_agg)

        if diffs:
            excel_total = excel_cols.get("Z", 0)
            if isinstance(excel_total, str):
                excel_total = 0
            csv_total = sum(csv_agg.values())
            total_match = abs(excel_total - csv_total) < 1

            if total_match:
                col_only_count += 1
            else:
                total_diff_count += 1

            results.append(CheckResult(
                school=school_name,
                month_label=month_label,
                result_type="TOTAL_MATCH" if total_match else "TOTAL_DIFF",
                sid=sid,
                name=name,
                row=row_num,
                diffs=diffs,
                excel_total=excel_total,
                csv_total=csv_total,
            ))
        else:
            ok_count += 1

    # コンソール出力
    print(f"\n■ 結果サマリー")
    print(f"  Excel生徒数:       {len(sales)}")
    print(f"  完全一致 (OK):     {ok_count}")
    print(f"  列配分違い(合計一致): {col_only_count}")
    print(f"  合計不一致(要確認):  {total_diff_count}")
    print(f"  CSV未登録:         {not_in_csv_count}")

    _print_details(results)

    return results


def _print_details(results: list[CheckResult]) -> None:
    """差異詳細をコンソール出力"""
    critical = [r for r in results if r.result_type == "TOTAL_DIFF"]
    not_in_csv = [r for r in results if r.result_type == "NOT_IN_CSV"]
    col_only = [r for r in results if r.result_type == "TOTAL_MATCH"]

    if critical:
        print(f"\n{'='*80}")
        print(f"  ★★★ 合計不一致（要確認）: {len(critical)}件 ★★★")
        print(f"{'='*80}")
        for r in critical:
            tdiff = r.excel_total - r.csv_total
            sign = "+" if tdiff > 0 else ""
            print(f"\n  生徒ID={r.sid} {r.name} (Row {r.row})"
                  f"  合計差額={sign}{tdiff:,.0f}")
            for col, ev, cv, diff in r.diffs:
                disp = COL_DISPLAY.get(col, col)
                s = "+" if diff > 0 else ""
                print(f"    {disp:20s}  Excel={ev:>10,.0f}  "
                      f"CSV={cv:>10,.0f}  差={s}{diff:,.0f}")
            print(f"    {'─'*60}")
            print(f"    {'合計':20s}  Excel={r.excel_total:>10,.0f}  "
                  f"CSV={r.csv_total:>10,.0f}  差={sign}{tdiff:,.0f}")

    if not_in_csv:
        print(f"\n{'─'*80}")
        print(f"  CSV未登録: {len(not_in_csv)}件")
        print(f"{'─'*80}")
        for r in not_in_csv:
            print(f"  生徒ID={r.sid} {r.name} "
                  f"(Row {r.row}) Excel売上={r.excel_total:,.0f}")

    if col_only:
        print(f"\n{'─'*80}")
        print(f"  列配分違い（合計は一致）: {len(col_only)}件")
        print(f"{'─'*80}")
        for r in col_only:
            print(f"  生徒ID={r.sid} {r.name} (Row {r.row})")
            for col, ev, cv, diff in r.diffs:
                disp = COL_DISPLAY.get(col, col)
                s = "+" if diff > 0 else ""
                print(f"    {disp:20s}  Excel={ev:>10,.0f}  "
                      f"CSV={cv:>10,.0f}  差={s}{diff:,.0f}")


# ====================================================================
# CSV出力
# ====================================================================

def write_results_csv(
    output_path: str,
    all_results: list[CheckResult],
) -> None:
    """照合結果をCSVに出力"""
    os.makedirs(os.path.dirname(output_path), exist_ok=True)

    with open(output_path, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f)
        writer.writerow([
            "校舎", "月", "重要度", "種別", "生徒ID", "生徒名",
            "行番号", "項目", "Excel金額", "CSV金額", "差額", "合計差額",
        ])

        for r in all_results:
            if r.result_type == "NOT_IN_CSV":
                writer.writerow([
                    r.school, r.month_label, "★要確認", "CSV未登録",
                    r.sid, r.name, r.row, "売上合計",
                    r.excel_total, 0, r.excel_total, r.excel_total,
                ])
            elif r.diffs:
                total_diff = r.excel_total - r.csv_total
                total_match = abs(total_diff) < 1
                severity = "列配分違い" if total_match else "★要確認"
                kind = "合計一致" if total_match else "合計不一致"

                for col, ev, cv, diff in r.diffs:
                    disp = COL_DISPLAY.get(col, col)
                    writer.writerow([
                        r.school, r.month_label, severity, kind,
                        r.sid, r.name, r.row, disp,
                        ev, cv, diff, total_diff,
                    ])


# ====================================================================
# メイン
# ====================================================================

def main():
    config = load_config()

    csv_base_dir = config["csv_base_dir"]
    output_dir = config.get("output_dir", "C:/Users/USER/Documents/照合結果")
    schools = [SchoolConfig(**s) for s in config["schools"]]

    print("=" * 80)
    print("  売上明細 vs 請求データ 照合チェック（全校舎対応）")
    print(f"  実行日時: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"  対象校舎: {', '.join(s.name for s in schools)}")
    print("=" * 80)

    # 請求CSV自動検出
    csv_files = discover_csv_files(csv_base_dir)
    print(f"\n  検出された請求CSV: {len(csv_files)}ヶ月分")
    for (y, m), path in sorted(csv_files.items()):
        print(f"    {y}年{m}月 ← {os.path.basename(path)}")

    all_results: list[CheckResult] = []

    for school in schools:
        print(f"\n{'#'*80}")
        print(f"  校舎: {school.name}")
        print(f"  Excel: {school.excel_dir}")
        print(f"{'#'*80}")

        excel_files = discover_excel_files(school)
        if not excel_files:
            print(f"  ⚠ Excelファイルが見つかりません: {school.excel_dir}")
            continue

        print(f"  検出されたExcel: {len(excel_files)}ヶ月分")
        for (y, m), path in sorted(excel_files.items()):
            print(f"    {y}年{m}月 ← {os.path.basename(path)}")

        pairs = match_months(csv_files, excel_files)
        if not pairs:
            print(f"  ⚠ CSVとExcelで一致する月がありません")
            continue

        print(f"  マッチした月: {len(pairs)}ヶ月")

        for pair in pairs:
            results = run_check(
                school_name=school.name,
                month_label=pair.label,
                csv_path=pair.csv_path,
                excel_path=pair.excel_path,
                sheet_index=school.sheet_index,
            )
            all_results.extend(results)

    # 結果CSV出力
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_path = os.path.join(output_dir, f"照合結果_{timestamp}.csv")
    write_results_csv(output_path, all_results)

    # 全校舎サマリー
    print(f"\n{'='*80}")
    print(f"  全校舎サマリー")
    print(f"{'='*80}")
    school_names = sorted(set(r.school for r in all_results))
    for sn in school_names:
        school_results = [r for r in all_results if r.school == sn]
        critical = sum(1 for r in school_results if r.result_type == "TOTAL_DIFF")
        col_only = sum(1 for r in school_results if r.result_type == "TOTAL_MATCH")
        not_csv = sum(1 for r in school_results if r.result_type == "NOT_IN_CSV")
        print(f"  {sn}: 合計不一致={critical}件, "
              f"列配分違い={col_only}件, CSV未登録={not_csv}件")

    print(f"\n■ 照合結果CSVを出力しました: {output_path}")
    print()


if __name__ == "__main__":
    main()
