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
    school_keywords: tuple[str, ...] = ()  # クラス情報の校舎名マッチ用


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
    # 前後月の引落額: {月ラベル: {列レター: 金額}}
    monthly_billing: dict = field(default_factory=dict)
    # 備考: {列レター: コメント文字列}
    remarks: dict = field(default_factory=dict)
    grade: str = ""
    # 対象月の前後月ラベルリスト（表示順）
    month_columns: list = field(default_factory=list)


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


BRAND_COLUMN_MAP = {
    "アンイングリッシュクラブ": ("M", "N"),
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
# ユーザークラス情報（校舎別受講フィルタ）
# ====================================================================

def discover_class_info_files(
    class_info_dir: str,
) -> dict[tuple[int, int], str]:
    """
    クラス情報Excelを自動検出。ファイル名から月を抽出。
    返り値: {(year, month): path}
    """
    result = {}
    d = Path(class_info_dir)
    if not d.exists():
        return result

    month_map = {
        "1月": 1, "2月": 2, "3月": 3, "4月": 4, "5月": 5, "6月": 6,
        "7月": 7, "8月": 8, "9月": 9, "10月": 10, "11月": 11, "12月": 12,
    }

    for f in d.iterdir():
        if not f.is_file() or f.suffix not in (".xlsx", ".xlsm"):
            continue
        # "1月末.xlsx", "11月分.xlsx" etc.
        m = re.search(r"(\d{1,2})月", f.name)
        if not m:
            continue
        month = int(m.group(1))
        # 年は修正日時から推定（ファイル名に年がないため）
        year = datetime.fromtimestamp(f.stat().st_mtime).year
        result[(year, month)] = str(f)

    return result


def read_class_info(
    class_info_path: str,
    school_keywords: tuple[str, ...],
) -> dict[str, set[str]]:
    """
    クラス情報を読み込み、指定校舎に在籍する生徒のブランド一覧を返す。
    返り値: {生徒ID: {この校舎で受講しているブランド名のset}}

    校舎名にschool_keywordsのいずれかが含まれていれば、その校舎のブランドとみなす。
    """
    wb = openpyxl.load_workbook(class_info_path, read_only=True, data_only=True)
    ws = wb[wb.sheetnames[0]]

    sid_brands: dict[str, set[str]] = {}

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, values_only=False):
        sid_val = row[6].value   # G=生徒ID
        brand = row[25].value    # Z=ブランド名
        school = row[31].value   # AF=校舎名

        if sid_val is None or school is None:
            continue

        sid = str(int(sid_val)) if isinstance(sid_val, (int, float)) else str(sid_val)
        school_str = str(school)

        # この校舎に該当するか
        if any(kw in school_str for kw in school_keywords):
            if sid not in sid_brands:
                sid_brands[sid] = set()
            if brand:
                sid_brands[sid].add(str(brand))

    wb.close()
    return sid_brands


def read_withdrawn_students(
    class_info_path: str,
    school_keywords: tuple[str, ...],
    target_year: int,
    target_month: int,
) -> set[str]:
    """
    対象月の初日時点で既に退会済みの生徒IDを返す。
    退会日 < 対象月の初日 → 退会者とみなす。
    """
    from datetime import date

    target_first = date(target_year, target_month, 1)

    wb = openpyxl.load_workbook(class_info_path, read_only=True, data_only=True)
    ws = wb[wb.sheetnames[0]]

    # 生徒ごとに全校舎・全ブランドの退会日を集める
    # この校舎のクラスが全て退会済みなら退会者とみなす
    sid_classes: dict[str, list[tuple[bool, bool]]] = {}  # {sid: [(is_this_school, is_withdrawn)]}

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, values_only=False):
        sid_val = row[6].value   # G=生徒ID
        school = row[31].value   # AF=校舎名
        taikai = row[21].value   # V=退会日

        if sid_val is None or school is None:
            continue

        sid = str(int(sid_val)) if isinstance(sid_val, (int, float)) else str(sid_val)
        school_str = str(school)

        is_this_school = any(kw in school_str for kw in school_keywords)
        if not is_this_school:
            continue

        # 退会日の判定
        is_withdrawn = False
        if taikai is not None:
            try:
                if hasattr(taikai, "date"):
                    taikai_date = taikai.date()
                elif hasattr(taikai, "year"):
                    taikai_date = taikai
                else:
                    taikai_date = datetime.strptime(
                        str(taikai).split()[0], "%Y-%m-%d"
                    ).date()
                is_withdrawn = taikai_date < target_first
            except (ValueError, AttributeError):
                pass

        if sid not in sid_classes:
            sid_classes[sid] = []
        sid_classes[sid].append(is_withdrawn)

    wb.close()

    # この校舎の全クラスが退会済みの生徒
    withdrawn = set()
    for sid, classes in sid_classes.items():
        if classes and all(classes):
            withdrawn.add(sid)

    return withdrawn


def filter_billing_by_school(
    billing_entries: list[tuple[str, str, float]],
    school_brands: set[str] | None,
) -> list[tuple[str, str, float]]:
    """
    請求明細から、この校舎で受講していないブランドの項目を除外する。
    school_brands が None の場合はフィルタなし（全項目を返す）。

    設備費・入会金・家族割等のブランド横断項目はフィルタしない。
    """
    if school_brands is None:
        return billing_entries

    # フィルタ不要のカテゴリ（ブランドに依存しない費目）
    PASSTHROUGH_CATEGORIES = {
        "設備費", "入会金", "模試代", "家族割", "割引", "過不足金",
        "諸経費", "家賃", "総合指導管理費", "0", "1",
        "入会時教材費", "入会時授業料1", "入会時授業料2",
        "入会時授業料3", "入会時授業料A", "入会時月会費", "入会時設備費",
    }

    # BRAND_COLUMN_MAP にあるブランドのみ校舎フィルタ対象
    # それ以外（学習塾系）はクラス情報に載っていなくても通す
    FILTERED_BRANDS = set(BRAND_COLUMN_MAP.keys())

    filtered = []
    for brand, category, amount in billing_entries:
        # ブランド横断の費目はそのまま通す
        if category in PASSTHROUGH_CATEGORIES:
            filtered.append((brand, category, amount))
            continue
        # ブランドが空の場合（社割等）はそのまま通す
        if not brand:
            filtered.append((brand, category, amount))
            continue
        # BRAND_COLUMN_MAP に含まれるブランド → 校舎フィルタで判定
        if brand in FILTERED_BRANDS:
            if brand in school_brands:
                filtered.append((brand, category, amount))
            elif any(sb in brand or brand in sb for sb in school_brands):
                filtered.append((brand, category, amount))
        else:
            # 学習塾系ブランド → クラス情報になくても通す
            filtered.append((brand, category, amount))

    return filtered


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
    """請求CSVを読み込み → {生徒ID: [(ブランド名, カテゴリ名, 月額料金), ...]}"""
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


# 学年情報キャッシュ
_grade_cache: dict[str, dict[str, str]] = {}


def read_grades_from_csv(csv_path: str) -> dict[str, str]:
    """請求CSVから生徒IDと学年の対応を取得 → {生徒ID: 学年}"""
    if csv_path in _grade_cache:
        return _grade_cache[csv_path]
    grades = {}
    with open(csv_path, encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        for row in reader:
            sid = row["生徒ID"].strip()
            grade = row.get("学年", "").strip()
            if sid and grade and sid not in grades:
                grades[sid] = grade
    _grade_cache[csv_path] = grades
    return grades


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

    # 上記BRAND_COLUMN_MAP以外のブランドは全て学習塾として扱う
    if category in ("授業料", "追加授業料"):
        return "E"
    if category == "月会費":
        return "F"
    if category in ("講習会費", "必須講座", "必須講習会", "テスト対策"):
        return "I"
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


def _find_nearest_class_info(
    year: int,
    month: int,
    class_info_files: dict[tuple[int, int], str],
) -> str | None:
    """対象月に最も近いクラス情報ファイルを返す"""
    if not class_info_files:
        return None
    target = year * 12 + month
    best_path = None
    best_dist = 9999
    for (cy, cm), path in class_info_files.items():
        dist = abs((cy * 12 + cm) - target)
        if dist < best_dist:
            best_dist = dist
            best_path = path
    return best_path


def _load_billing_by_priority(
    csv_paths_by_priority: list[str],
) -> dict[str, list[tuple[str, str, float]]]:
    """
    複数月のCSVを優先度順に読み込む。
    csv_paths_by_priority[0] = 対象月（最優先）
    csv_paths_by_priority[1:] = 前後月（対象月にない生徒のみ補完）

    各生徒の明細は1つの月からのみ取得（合算しない）。
    対象月にいれば対象月のデータ、いなければ前後月で見つかったデータを使う。
    """
    result: dict[str, list[tuple[str, str, float]]] = {}

    for path in csv_paths_by_priority:
        billing = read_billing_csv(path)
        for sid, entries in billing.items():
            if sid not in result:
                result[sid] = entries

    return result


def _get_adjacent_months(
    year: int, month: int, offset: int,
) -> list[tuple[int, int]]:
    """対象月の前後offset月を含むリストを返す"""
    months = []
    for delta in range(-offset, offset + 1):
        m = month + delta
        y = year
        while m < 1:
            m += 12
            y -= 1
        while m > 12:
            m -= 12
            y += 1
        months.append((y, m))
    return months


def _preload_all_billings(
    csv_paths_with_labels: list[tuple[str, str]],
) -> dict[str, dict[str, list[tuple[str, str, float]]]]:
    """
    全月のCSVを一括読み込み。
    返り値: {月ラベル: {生徒ID: [(ブランド名, カテゴリ名, 月額料金), ...]}}
    """
    all_billings: dict[str, dict[str, list[tuple[str, str, float]]]] = {}
    for path, label in csv_paths_with_labels:
        all_billings[label] = read_billing_csv(path)
    return all_billings


def _compute_monthly_billing(
    sid: str,
    all_billings: dict[str, dict[str, list[tuple[str, str, float]]]],
    month_col_labels: list[str],
    school_brands: dict[str, set[str]] | None,
) -> dict[str, dict[str, float]]:
    """
    事前読み込み済みのデータから生徒1人の各月別引落額を計算。
    返り値: {月ラベル: {列レター: 金額}}
    """
    monthly: dict[str, dict[str, float]] = {}

    for label in month_col_labels:
        billing = all_billings.get(label, {})
        entries = billing.get(sid, [])
        if school_brands is not None:
            student_brands = school_brands.get(sid)
            entries = filter_billing_by_school(entries, student_brands)
        agg = aggregate_csv_for_student(entries)
        monthly[label] = agg

    return monthly


def _find_similar_billing(
    sid: str,
    excel_amount: float,
    all_billings: dict[str, dict[str, list[tuple[str, str, float]]]],
    month_col_labels: list[str],
) -> str:
    """
    売上あり請求なしの項目について、前後月の全請求から
    同額または近い金額の請求を探してコメントを返す。
    """
    hints = []
    for label in month_col_labels:
        billing = all_billings.get(label, {})
        entries = billing.get(sid, [])
        for brand, category, amount in entries:
            if abs(amount - excel_amount) < 1 and amount != 0:
                hints.append(f"{label}「{brand}/{category}」{amount:,.0f}")
            elif abs(amount) >= abs(excel_amount) * 0.5 and abs(amount - excel_amount) < abs(excel_amount) * 0.3 and amount != 0:
                hints.append(f"{label}「{brand}/{category}」{amount:,.0f}(近似)")
    if hints:
        return " / ".join(hints[:3])
    return ""


def run_check(
    school_name: str,
    month_label: str,
    csv_paths: list[str],
    csv_paths_with_labels: list[tuple[str, str]],
    excel_path: str,
    sheet_index: int = 2,
    school_brands: dict[str, set[str]] | None = None,
    withdrawn_sids: set[str] | None = None,
) -> list[CheckResult]:
    """
    1校舎・1ヶ月分のチェックを実行。
    csv_paths: 対象月±2ヶ月分のCSVパスリスト（優先度順）
    csv_paths_with_labels: [(path, "11月"), (path, "12月"), ...] 表示順
    school_brands: {生徒ID: {この校舎で受講しているブランド}} or None
    """
    csv_names = [os.path.basename(p) for p in csv_paths]
    print(f"\n{'='*80}")
    print(f"  [{school_name}] {month_label} 照合チェック")
    print(f"  CSV:   {csv_names[0]} 他{len(csv_paths)-1}ヶ月分" if len(csv_paths) > 1
          else f"  CSV:   {csv_names[0]}")
    print(f"  Excel: {os.path.basename(excel_path)}")
    if school_brands is not None:
        print(f"  校舎フィルタ: 有効（クラス情報から{len(school_brands)}生徒分）")
    print(f"{'='*80}")

    billing = _load_billing_by_priority(csv_paths)
    sales = read_excel_sales(excel_path, sheet_index)
    grades = read_grades_from_csv(csv_paths[0])  # 対象月CSVから学年取得

    month_col_labels = [label for _, label in csv_paths_with_labels]

    # 月別引落額計算用: 全CSVを事前に一括読み込み
    all_billings = _preload_all_billings(csv_paths_with_labels)

    ok_count = 0
    col_only_count = 0
    total_diff_count = 0
    not_in_csv_count = 0
    no_billing_count = 0
    withdrawn_count = 0
    results: list[CheckResult] = []

    _withdrawn = withdrawn_sids or set()

    for sid, sdata in sorted(sales.items(), key=lambda x: x[1]["row"]):
        name = sdata["name"]
        excel_cols = sdata["cols"]
        row_num = sdata["row"]

        # 退会者なのに売上がある
        excel_total_raw = excel_cols.get("Z", 0)
        if isinstance(excel_total_raw, str):
            excel_total_raw = 0
        if sid in _withdrawn and excel_total_raw > 0:
            withdrawn_count += 1
            results.append(CheckResult(
                school=school_name,
                month_label=month_label,
                result_type="WITHDRAWN",
                sid=sid,
                name=name,
                row=row_num,
                excel_total=excel_total_raw,
                month_columns=month_col_labels,
                grade=grades.get(sid, ""),
            ))
            continue

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
                    month_columns=month_col_labels,
                ))
            continue

        entries = billing[sid]
        if school_brands is not None:
            student_brands = school_brands.get(sid)
            entries = filter_billing_by_school(entries, student_brands)
        csv_agg = aggregate_csv_for_student(entries)
        diffs = compare_student(excel_cols, csv_agg)

        if diffs:
            excel_total = excel_cols.get("Z", 0)
            if isinstance(excel_total, str):
                excel_total = 0
            csv_total = sum(csv_agg.values())
            total_match = abs(excel_total - csv_total) < 1

            # 各月別の引落額を計算（事前読み込み済みデータを使用）
            monthly = _compute_monthly_billing(
                sid, all_billings, month_col_labels, school_brands,
            )

            # 項目レベルで「売上あり請求なし」と「その他の差異」を分離
            # 売上>0 かつ 請求額と不一致 → 売上あり請求なし（Y列は対象外）
            unbilled_diffs = [
                d for d in diffs
                if d[1] > 0 and d[0] != "Y"
            ]
            other_diffs = [
                d for d in diffs
                if not (d[1] > 0 and d[0] != "Y")
            ]

            if unbilled_diffs:
                no_billing_count += 1
                # 備考: 前後月で同額/近似額の請求を探す
                remarks = {}
                for col, ev, cv, diff in unbilled_diffs:
                    hint = _find_similar_billing(
                        sid, ev, all_billings, month_col_labels,
                    )
                    if hint:
                        remarks[col] = hint
                results.append(CheckResult(
                    school=school_name,
                    month_label=month_label,
                    result_type="NO_BILLING",
                    sid=sid, name=name, row=row_num,
                    diffs=unbilled_diffs,
                    excel_total=excel_total,
                    csv_total=csv_total,
                    monthly_billing=monthly,
                    month_columns=month_col_labels,
                    remarks=remarks,
                    grade=grades.get(sid, ""),
                ))

            if other_diffs:
                if total_match:
                    col_only_count += 1
                    rtype = "TOTAL_MATCH"
                else:
                    total_diff_count += 1
                    rtype = "TOTAL_DIFF"
                results.append(CheckResult(
                    school=school_name,
                    month_label=month_label,
                    result_type=rtype,
                    sid=sid, name=name, row=row_num,
                    diffs=other_diffs,
                    excel_total=excel_total,
                    csv_total=csv_total,
                    monthly_billing=monthly,
                    month_columns=month_col_labels,
                    grade=grades.get(sid, ""),
                ))
            elif not unbilled_diffs:
                # unbilledもotherもない（ありえないが安全策）
                pass
        else:
            ok_count += 1

    # コンソール出力
    print(f"\n■ 結果サマリー")
    print(f"  Excel生徒数:       {len(sales)}")
    print(f"  完全一致 (OK):     {ok_count}")
    print(f"  列配分違い(合計一致): {col_only_count}")
    print(f"  合計不一致(要確認):  {total_diff_count}")
    print(f"  売上あり請求なし:    {no_billing_count}")
    print(f"  退会者売上あり:     {withdrawn_count}")
    print(f"  月謝未計上:        {not_in_csv_count}")

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
        print(f"  月謝未計上: {len(not_in_csv)}件")
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
    """照合結果をCSVに出力（月別引落額付き）"""
    os.makedirs(os.path.dirname(output_path), exist_ok=True)

    # 全結果から月カラムのユニオンを取得
    all_month_cols: list[str] = []
    for r in all_results:
        for mc in r.month_columns:
            if mc not in all_month_cols:
                all_month_cols.append(mc)
    all_month_cols.sort()

    with open(output_path, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f)
        header = [
            "校舎", "月", "重要度", "種別", "生徒ID", "生徒名",
            "行番号", "項目", "売上",
        ]
        for mc in all_month_cols:
            header.append(f"{mc}引落")
        header.extend(["差額", "合計差額"])
        writer.writerow(header)

        for r in all_results:
            if r.result_type == "NOT_IN_CSV":
                row_data = [
                    r.school, r.month_label, "★要確認", "月謝未計上",
                    r.sid, r.name, r.row, "売上合計",
                    r.excel_total,
                ]
                row_data.extend([0] * len(all_month_cols))
                row_data.extend([r.excel_total, r.excel_total])
                writer.writerow(row_data)
            elif r.diffs:
                total_diff = r.excel_total - r.csv_total
                total_match = abs(total_diff) < 1

                # 売上あり・請求なし判定
                if r.csv_total == 0 and r.excel_total > 0:
                    severity = "★売上あり請求なし"
                    kind = "請求漏れ"
                elif total_match:
                    severity = "列配分違い"
                    kind = "合計一致"
                else:
                    severity = "★要確認"
                    kind = "合計不一致"

                for col, ev, cv, diff in r.diffs:
                    disp = COL_DISPLAY.get(col, col)
                    row_data = [
                        r.school, r.month_label, severity, kind,
                        r.sid, r.name, r.row, disp, ev,
                    ]
                    for mc in all_month_cols:
                        monthly_agg = r.monthly_billing.get(mc, {})
                        row_data.append(monthly_agg.get(col, 0))
                    row_data.extend([diff, total_diff])
                    writer.writerow(row_data)


# ====================================================================
# メイン
# ====================================================================

def main():
    config = load_config()

    csv_base_dir = config["csv_base_dir"]
    class_info_dir = config.get("class_info_dir", "")
    output_dir = config.get("output_dir", "C:/Users/USER/Documents/照合結果")
    schools = []
    for s in config["schools"]:
        s = dict(s)  # コピーして元configを壊さない
        kw = s.pop("school_keywords", [])
        schools.append(SchoolConfig(**s, school_keywords=tuple(kw)))

    print("=" * 80)
    print("  売上明細 vs 請求データ 照合チェック（全校舎対応）")
    print(f"  実行日時: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"  対象校舎: {', '.join(s.name for s in schools)}")
    print(f"  チェック範囲: 対象月 ±2ヶ月")
    print("=" * 80)

    # 請求CSV自動検出
    csv_files = discover_csv_files(csv_base_dir)
    print(f"\n  検出された請求CSV: {len(csv_files)}ヶ月分")
    for (y, m), path in sorted(csv_files.items()):
        print(f"    {y}年{m}月 ← {os.path.basename(path)}")

    # クラス情報自動検出
    class_info_files = {}
    if class_info_dir:
        class_info_files = discover_class_info_files(class_info_dir)
        if class_info_files:
            print(f"\n  検出されたクラス情報: {len(class_info_files)}ヶ月分")
            for (y, m), path in sorted(class_info_files.items()):
                print(f"    {y}年{m}月 ← {os.path.basename(path)}")
        else:
            print(f"\n  ⚠ クラス情報が見つかりません: {class_info_dir}")

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
            # 対象月を先頭に、±2ヶ月分のCSVパスを優先度順に収集
            csv_paths = [pair.csv_path]  # 対象月が最優先
            adjacent = _get_adjacent_months(pair.year, pair.month, 2)
            for ym in adjacent:
                if ym in csv_files and csv_files[ym] != pair.csv_path:
                    csv_paths.append(csv_files[ym])

            # 表示用: 月ラベル付きCSVリスト（時系列順）
            csv_paths_with_labels = []
            for ym in adjacent:
                if ym in csv_files:
                    csv_paths_with_labels.append(
                        (csv_files[ym], f"{ym[0]}年{ym[1]}月")
                    )

            # クラス情報から校舎別ブランドフィルタ＋退会者リストを構築
            school_brands = None
            withdrawn_sids = None
            if school.school_keywords and class_info_files:
                best_ci = _find_nearest_class_info(
                    pair.year, pair.month, class_info_files,
                )
                if best_ci:
                    school_brands = read_class_info(
                        best_ci, school.school_keywords,
                    )
                    withdrawn_sids = read_withdrawn_students(
                        best_ci, school.school_keywords,
                        pair.year, pair.month,
                    )

            results = run_check(
                school_name=school.name,
                month_label=pair.label,
                csv_paths=csv_paths,
                csv_paths_with_labels=csv_paths_with_labels,
                excel_path=pair.excel_path,
                sheet_index=school.sheet_index,
                school_brands=school_brands,
                withdrawn_sids=withdrawn_sids,
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
              f"列配分違い={col_only}件, 月謝未計上={not_csv}件")

    print(f"\n■ 照合結果CSVを出力しました: {output_path}")
    print()


if __name__ == "__main__":
    main()
