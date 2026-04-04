"""
小幡売上明細 vs 請求CSV 照合チェックツール

売上申請（Excel）に記載された金額が、請求データ（CSV）と一致しているか検証する。
"""

import csv
import sys
import os
from pathlib import Path
from datetime import datetime

try:
    import openpyxl
except ImportError:
    print("openpyxl が必要です: pip install openpyxl")
    sys.exit(1)


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

# 学習塾系ブランド → E(授業料), F(月会費)
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

# ブランド → (授業料列, 月会費列) のマッピング
BRAND_COLUMN_MAP = {
    # アン英語系 → M, N
    "アンイングリッシュクラブ": ("M", "N"),
    "アンさんこくキッズ": ("M", "N"),
    # そろばん → O, P
    "アンそろばんクラブ": ("O", "P"),
    "アンそろばんクラブ【選択講習会】": ("O", "P"),
    # プログラミング → J, K
    "アンプログラミングクラブ": ("J", "K"),
    # 筆っこ（美文字） → R, S
    "アン美文字クラブ": ("R", "S"),
    # 将棋 → U, V
    "アン将棋クラブ": ("U", "V"),
}

# 設備費列は全ブランド共通で D列


def col_letter_to_index(letter):
    """A=0, B=1, ... Z=25"""
    return ord(letter) - ord("A")


def parse_number(val):
    """数値をfloatに変換。空・Noneは0"""
    if val is None or val == "" or val == 0:
        return 0.0
    try:
        return float(str(val).replace(",", ""))
    except (ValueError, TypeError):
        return 0.0


def read_billing_csv(csv_path):
    """
    請求CSVを読み込み、生徒ID別にカテゴリ・ブランド別の月額料金を集計。
    返り値: {生徒ID: [(ブランド名, カテゴリ名, 月額料金), ...]}
    """
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


def aggregate_csv_for_student(entries):
    """
    CSV明細をExcel列に対応する形で集計する。
    返り値: {列レター: 金額}
    """
    col_totals = {}

    for brand, category, amount in entries:
        if amount == 0:
            continue

        target_col = None

        # 設備費は全ブランド共通 → D列
        if category == "設備費":
            target_col = "D"
        # 入会金 → X列
        elif category == "入会金":
            target_col = "X"
        # 模試代 → H列
        elif category == "模試代":
            target_col = "H"
        # その他系（カテゴリが数字の場合も含む）
        elif category in ("0", "1"):
            target_col = "Y"  # その他
        elif category in ("入会時教材費", "入会時授業料1", "入会時授業料2",
                          "入会時授業料3", "入会時授業料A", "入会時月会費",
                          "入会時設備費"):
            target_col = "Y"  # 入会時系はその他
        # 家族割・割引・過不足金・諸経費 → その他
        elif category in ("家族割", "割引", "過不足金", "諸経費", "家賃",
                          "総合指導管理費"):
            target_col = "Y"
        # ブランド別のマッピング
        elif brand in BRAND_COLUMN_MAP:
            jugyou_col, getsukai_col = BRAND_COLUMN_MAP[brand]
            if category == "授業料" or category == "追加授業料":
                target_col = jugyou_col
            elif category == "月会費":
                target_col = getsukai_col
            elif category in ("講習会費", "必須講座", "必須講習会",
                              "テスト対策"):
                # 講習会費列はブランドにより異なる
                if brand.startswith("アンそろばん"):
                    target_col = "Q"
                elif brand == "アンプログラミングクラブ":
                    target_col = "L"
                elif brand == "アン美文字クラブ":
                    target_col = "T"
                elif brand == "アン将棋クラブ":
                    target_col = "W"
                else:
                    target_col = "I"  # デフォルト講習会費
            elif category == "教材費":
                target_col = "Y"
            elif category == "休会費":
                target_col = "Y"
            else:
                target_col = "Y"
        elif brand in JUKU_BRANDS:
            if category == "授業料" or category == "追加授業料":
                target_col = "E"
            elif category == "月会費":
                target_col = "F"
            elif category in ("講習会費", "必須講座", "必須講習会",
                              "テスト対策"):
                target_col = "I"
            elif category == "教材費":
                target_col = "Y"
            elif category == "休会費":
                target_col = "Y"
            else:
                target_col = "Y"
        # そろばん検定
        elif "そろばん検定" in brand:
            target_col = "Y"
        else:
            # 未知のブランド → その他
            target_col = "Y"

        if target_col:
            col_totals[target_col] = col_totals.get(target_col, 0) + amount

    return col_totals


def read_excel_sales(excel_path):
    """
    売上明細Excelを読み込む。
    返り値: {生徒ID: {"name": 生徒名, "cols": {列レター: 金額}, "row": 行番号}}
    """
    wb = openpyxl.load_workbook(excel_path, read_only=True, data_only=True)
    # 売上明細シートは3番目（index 2）
    ws = wb[wb.sheetnames[2]]

    result = {}
    for row in ws.iter_rows(min_row=5, max_row=ws.max_row, values_only=False):
        # 集計行で終了
        a_val = row[0].value
        if a_val is not None and isinstance(a_val, str) and "計" in a_val:
            break

        sid = row[1].value  # B列 = 生徒ID
        if sid is None:
            continue

        sid = str(int(sid)) if isinstance(sid, (int, float)) else str(sid)
        name = row[2].value or ""  # C列 = 生徒名

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


# 列レター → カテゴリ表示名
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


def compare_student(sid, excel_cols, csv_cols, student_name):
    """
    1人の生徒について、Excel列ごとにCSVと比較。
    返り値: 差異リスト [(列, Excel値, CSV値, 差額)]
    """
    diffs = []
    all_cols = sorted(set(list(excel_cols.keys()) + list(csv_cols.keys())))

    for col in all_cols:
        if col == "Z":
            continue  # 合計は別途チェック
        excel_val = excel_cols.get(col, 0)
        csv_val = csv_cols.get(col, 0)

        # 文字列が入っている場合（休会メモ等）はスキップ
        if isinstance(excel_val, str):
            continue

        if abs(excel_val - csv_val) >= 1:  # 1円以上の差異
            diffs.append((col, excel_val, csv_val, excel_val - csv_val))

    return diffs


def run_check(month_label, csv_path, excel_path):
    """1ヶ月分のチェックを実行"""
    print(f"\n{'='*80}")
    print(f"  {month_label} 照合チェック")
    print(f"  CSV:   {os.path.basename(csv_path)}")
    print(f"  Excel: {os.path.basename(excel_path)}")
    print(f"{'='*80}")

    billing = read_billing_csv(csv_path)
    sales = read_excel_sales(excel_path)

    ok_count = 0
    col_only_count = 0
    total_diff_count = 0
    not_in_csv_count = 0
    results = []

    for sid, sdata in sorted(sales.items(), key=lambda x: x[1]["row"]):
        name = sdata["name"]
        excel_cols = sdata["cols"]
        row_num = sdata["row"]

        if sid not in billing:
            # CSVに存在しない生徒
            total = excel_cols.get("Z", 0)
            if total != 0 and not isinstance(total, str):
                not_in_csv_count += 1
                results.append({
                    "type": "NOT_IN_CSV",
                    "sid": sid,
                    "name": name,
                    "row": row_num,
                    "excel_total": total,
                })
            continue

        csv_agg = aggregate_csv_for_student(billing[sid])
        diffs = compare_student(sid, excel_cols, csv_agg, name)

        if diffs:
            # 合計チェック
            excel_total = excel_cols.get("Z", 0)
            if isinstance(excel_total, str):
                excel_total = 0
            csv_total = sum(csv_agg.values())
            total_match = abs(excel_total - csv_total) < 1

            if total_match:
                col_only_count += 1
            else:
                total_diff_count += 1

            results.append({
                "type": "TOTAL_MATCH" if total_match else "TOTAL_DIFF",
                "sid": sid,
                "name": name,
                "row": row_num,
                "diffs": diffs,
                "excel_total": excel_total,
                "csv_total": csv_total,
            })
        else:
            ok_count += 1

    # 結果表示
    print(f"\n■ 結果サマリー")
    print(f"  Excel生徒数:       {len(sales)}")
    print(f"  完全一致 (OK):     {ok_count}")
    print(f"  列配分違い(合計一致): {col_only_count}")
    print(f"  合計不一致(要確認):  {total_diff_count}")
    print(f"  CSV未登録:         {not_in_csv_count}")

    # === 合計不一致（要確認） ===
    critical = [r for r in results if r["type"] == "TOTAL_DIFF"]
    not_in_csv = [r for r in results if r["type"] == "NOT_IN_CSV"]
    col_only = [r for r in results if r["type"] == "TOTAL_MATCH"]

    if critical:
        print(f"\n{'='*80}")
        print(f"  ★★★ 合計不一致（要確認）: {len(critical)}件 ★★★")
        print(f"{'='*80}")
        for r in critical:
            et = r["excel_total"]
            ct = r["csv_total"]
            tdiff = et - ct
            sign = "+" if tdiff > 0 else ""
            print(f"\n  生徒ID={r['sid']} {r['name']} (Row {r['row']})"
                  f"  合計差額={sign}{tdiff:,.0f}")
            for col, ev, cv, diff in r["diffs"]:
                disp = COL_DISPLAY.get(col, col)
                sign2 = "+" if diff > 0 else ""
                print(f"    {disp:20s}  Excel={ev:>10,.0f}  "
                      f"CSV={cv:>10,.0f}  差={sign2}{diff:,.0f}")
            print(f"    {'─'*60}")
            print(f"    {'合計':20s}  Excel={et:>10,.0f}  "
                  f"CSV={ct:>10,.0f}  差={sign}{tdiff:,.0f}")

    if not_in_csv:
        print(f"\n{'─'*80}")
        print(f"  CSV未登録: {len(not_in_csv)}件")
        print(f"{'─'*80}")
        for r in not_in_csv:
            print(f"  生徒ID={r['sid']} {r['name']} "
                  f"(Row {r['row']}) Excel売上={r['excel_total']:,.0f}")

    if col_only:
        print(f"\n{'─'*80}")
        print(f"  列配分違い（合計は一致）: {len(col_only)}件")
        print(f"{'─'*80}")
        for r in col_only:
            print(f"  生徒ID={r['sid']} {r['name']} (Row {r['row']})")
            for col, ev, cv, diff in r["diffs"]:
                disp = COL_DISPLAY.get(col, col)
                sign = "+" if diff > 0 else ""
                print(f"    {disp:20s}  Excel={ev:>10,.0f}  "
                      f"CSV={cv:>10,.0f}  差={sign}{diff:,.0f}")

    return results


def main():
    # ファイルパス定義
    base_csv = Path("Y:/_★20170701作業用/100 真由美/野田より/◆日にち別お仕事◆")
    base_excel = Path("C:/Users/USER/Documents/小幡売上")

    months = [
        {
            "label": "2026年1月",
            "csv": base_csv / "20251217_送信日" / "AC_5_保護者別請求合計結果(全体)_202512171411_UTF8.csv",
            "excel": base_excel / "小幡売上明細書(2026.01).xlsm",
        },
        {
            "label": "2026年2月",
            "csv": base_csv / "20260114_送信日" / "AC_5_保護者別請求合計結果(全体)_202601141831_UTF8送信ちょっと後.csv",
            "excel": base_excel / "小幡売上明細書(2026.02).xlsm",
        },
        {
            "label": "2026年3月",
            "csv": base_csv / "20260213_送信日" / "AC_5_保護者別請求合計結果(全体)_202602132141_UTF8_3月月謝確定直後.csv",
            "excel": base_excel / "小幡売上明細書(2026.03).xlsm",
        },
    ]

    print("=" * 80)
    print("  小幡売上明細 vs 請求データ 照合チェック")
    print(f"  実行日時: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 80)

    all_results = []
    for m in months:
        if not m["csv"].exists():
            print(f"\n⚠ CSV not found: {m['csv']}")
            continue
        if not m["excel"].exists():
            print(f"\n⚠ Excel not found: {m['excel']}")
            continue
        results = run_check(m["label"], str(m["csv"]), str(m["excel"]))
        all_results.extend(results)

    # CSV出力
    output_path = Path("C:/Users/USER/Documents/小幡売上/照合結果.csv")
    with open(output_path, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(["月", "重要度", "種別", "生徒ID", "生徒名", "行番号",
                         "項目", "Excel金額", "CSV金額", "差額", "合計差額"])
        for m in months:
            if not m["csv"].exists() or not m["excel"].exists():
                continue
            billing = read_billing_csv(str(m["csv"]))
            sales = read_excel_sales(str(m["excel"]))

            for sid, sdata in sorted(sales.items(), key=lambda x: x[1]["row"]):
                name = sdata["name"]
                excel_cols = sdata["cols"]
                row_num = sdata["row"]

                if sid not in billing:
                    total = excel_cols.get("Z", 0)
                    if total != 0 and not isinstance(total, str):
                        writer.writerow([m["label"], "★要確認", "CSV未登録",
                                         sid, name, row_num, "売上合計",
                                         total, 0, total, total])
                    continue

                csv_agg = aggregate_csv_for_student(billing[sid])
                diffs = compare_student(sid, excel_cols, csv_agg, name)
                if not diffs:
                    continue

                excel_total = excel_cols.get("Z", 0)
                if isinstance(excel_total, str):
                    excel_total = 0
                csv_total = sum(csv_agg.values())
                total_diff = excel_total - csv_total
                total_match = abs(total_diff) < 1

                severity = "列配分違い" if total_match else "★要確認"
                kind = "合計一致" if total_match else "合計不一致"

                for col, ev, cv, diff in diffs:
                    disp = COL_DISPLAY.get(col, col)
                    writer.writerow([m["label"], severity, kind, sid, name,
                                     row_num, disp, ev, cv, diff, total_diff])

    print(f"\n■ 照合結果CSVを出力しました: {output_path}")
    print()


if __name__ == "__main__":
    main()
