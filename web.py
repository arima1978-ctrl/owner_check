"""
売上照合チェック Webアプリ

事務員がブラウザからアクセスし、照合結果を確認・CSVダウンロードできる。
社内ネットワークで http://<サーバーIP>:5050 でアクセス。
"""

import io
import csv
import os
from datetime import datetime
from pathlib import Path

from flask import Flask, render_template_string, send_file, redirect, url_for

from owner_check import (
    load_config,
    SchoolConfig,
    CheckResult,
    discover_csv_files,
    discover_excel_files,
    discover_class_info_files,
    match_months,
    read_class_info,
    filter_billing_by_school,
    _load_billing_by_priority,
    _get_adjacent_months,
    _find_nearest_class_info,
    read_excel_sales,
    aggregate_csv_for_student,
    compare_student,
    COL_DISPLAY,
)

app = Flask(__name__)

# ====================================================================
# 照合ロジック（Web用にコンソール出力なし）
# ====================================================================

def run_check_silent(
    school_name: str,
    month_label: str,
    csv_paths: list[str],
    excel_path: str,
    sheet_index: int = 2,
    school_brands: dict[str, set[str]] | None = None,
) -> tuple[list[CheckResult], dict]:
    """コンソール出力なしで照合チェックを実行"""
    billing = _load_billing_by_priority(csv_paths)
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
                    sid=sid, name=name, row=row_num,
                    excel_total=total,
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

            if total_match:
                col_only_count += 1
            else:
                total_diff_count += 1

            results.append(CheckResult(
                school=school_name,
                month_label=month_label,
                result_type="TOTAL_MATCH" if total_match else "TOTAL_DIFF",
                sid=sid, name=name, row=row_num,
                diffs=diffs,
                excel_total=excel_total,
                csv_total=csv_total,
            ))
        else:
            ok_count += 1

    summary = {
        "total": len(sales),
        "ok": ok_count,
        "col_only": col_only_count,
        "total_diff": total_diff_count,
        "not_in_csv": not_in_csv_count,
    }

    return results, summary


def run_all_checks() -> tuple[list[CheckResult], list[dict]]:
    """全校舎・全月の照合を実行"""
    config = load_config()
    csv_base_dir = config["csv_base_dir"]
    class_info_dir = config.get("class_info_dir", "")
    schools = []
    for s in config["schools"]:
        s = dict(s)
        kw = s.pop("school_keywords", [])
        schools.append(SchoolConfig(**s, school_keywords=tuple(kw)))

    csv_files = discover_csv_files(csv_base_dir)
    class_info_files = {}
    if class_info_dir:
        class_info_files = discover_class_info_files(class_info_dir)

    all_results: list[CheckResult] = []
    all_summaries: list[dict] = []

    for school in schools:
        excel_files = discover_excel_files(school)
        if not excel_files:
            continue

        pairs = match_months(csv_files, excel_files)
        for pair in pairs:
            # ±2ヶ月分のCSV（対象月を先頭に）
            csv_paths = [pair.csv_path]
            adjacent = _get_adjacent_months(pair.year, pair.month, 2)
            for ym in adjacent:
                if ym in csv_files and csv_files[ym] != pair.csv_path:
                    csv_paths.append(csv_files[ym])

            # 校舎フィルタ
            school_brands = None
            if school.school_keywords and class_info_files:
                best_ci = _find_nearest_class_info(
                    pair.year, pair.month, class_info_files,
                )
                if best_ci:
                    school_brands = read_class_info(
                        best_ci, school.school_keywords,
                    )

            results, summary = run_check_silent(
                school_name=school.name,
                month_label=pair.label,
                csv_paths=csv_paths,
                excel_path=pair.excel_path,
                sheet_index=school.sheet_index,
                school_brands=school_brands,
            )
            all_results.extend(results)
            all_summaries.append({
                "school": school.name,
                "month": pair.label,
                **summary,
            })

    return all_results, all_summaries


def results_to_csv_bytes(results: list[CheckResult]) -> bytes:
    """結果をCSVバイト列に変換"""
    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow([
        "校舎", "月", "重要度", "種別", "生徒ID", "生徒名",
        "行番号", "項目", "Excel金額", "CSV金額", "差額", "合計差額",
    ])

    for r in results:
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

    return output.getvalue().encode("utf-8-sig")


# ====================================================================
# HTMLテンプレート
# ====================================================================

HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="ja">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>売上照合チェック</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: 'Yu Gothic', 'Meiryo', sans-serif; background: #f5f5f5; color: #333; }
        .header { background: #1a5276; color: white; padding: 16px 24px; }
        .header h1 { font-size: 20px; font-weight: 500; }
        .header .sub { font-size: 12px; color: #aed6f1; margin-top: 4px; }
        .container { max-width: 1400px; margin: 0 auto; padding: 20px; }

        .actions { display: flex; gap: 12px; margin-bottom: 20px; align-items: center; }
        .btn { display: inline-block; padding: 10px 20px; border: none; border-radius: 6px;
               font-size: 14px; cursor: pointer; text-decoration: none; font-weight: 500; }
        .btn-primary { background: #2980b9; color: white; }
        .btn-primary:hover { background: #2471a3; }
        .btn-success { background: #27ae60; color: white; }
        .btn-success:hover { background: #229954; }
        .btn-loading { background: #95a5a6; color: white; pointer-events: none; }

        .summary-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(320px, 1fr));
                        gap: 16px; margin-bottom: 24px; }
        .summary-card { background: white; border-radius: 8px; padding: 16px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); }
        .summary-card h3 { font-size: 15px; color: #1a5276; margin-bottom: 12px;
                           border-bottom: 2px solid #eee; padding-bottom: 8px; }
        .summary-row { display: flex; justify-content: space-between; padding: 4px 0; font-size: 13px; }
        .summary-row .label { color: #666; }
        .summary-row .value { font-weight: 600; }
        .value.ok { color: #27ae60; }
        .value.warn { color: #e67e22; }
        .value.critical { color: #e74c3c; }

        .tabs { display: flex; gap: 4px; margin-bottom: 0; }
        .tab { padding: 10px 20px; background: #ddd; border: none; border-radius: 8px 8px 0 0;
               cursor: pointer; font-size: 13px; font-weight: 500; }
        .tab.active { background: white; color: #e74c3c; font-weight: 700; }
        .tab.tab-col { }
        .tab.tab-col.active { color: #e67e22; }
        .tab.tab-csv { }
        .tab.tab-csv.active { color: #7f8c8d; }

        .table-wrap { background: white; border-radius: 0 8px 8px 8px; box-shadow: 0 1px 3px rgba(0,0,0,0.1);
                      overflow-x: auto; }
        table { width: 100%; border-collapse: collapse; font-size: 13px; }
        th { background: #f8f9fa; padding: 10px 12px; text-align: left; font-weight: 600;
             border-bottom: 2px solid #dee2e6; white-space: nowrap; position: sticky; top: 0; }
        td { padding: 8px 12px; border-bottom: 1px solid #eee; }
        tr:hover { background: #f8f9fa; }
        .num { text-align: right; font-family: 'Consolas', monospace; }
        .positive { color: #e74c3c; }
        .negative { color: #2980b9; }
        .badge { display: inline-block; padding: 2px 8px; border-radius: 4px; font-size: 11px; font-weight: 600; }
        .badge-critical { background: #fde8e8; color: #e74c3c; }
        .badge-warn { background: #fef3e2; color: #e67e22; }
        .badge-info { background: #eaf2f8; color: #2980b9; }

        .filter-bar { display: flex; gap: 12px; margin-bottom: 16px; align-items: center; flex-wrap: wrap; }
        .filter-bar label { font-size: 13px; font-weight: 500; }
        .filter-bar select { padding: 6px 10px; border: 1px solid #ccc; border-radius: 4px; font-size: 13px; }

        .empty { text-align: center; padding: 40px; color: #999; }
        .loading { text-align: center; padding: 60px; }
        .loading .spinner { display: inline-block; width: 40px; height: 40px; border: 4px solid #eee;
                            border-top: 4px solid #2980b9; border-radius: 50%; animation: spin 0.8s linear infinite; }
        @keyframes spin { to { transform: rotate(360deg); } }
        .timestamp { font-size: 12px; color: #999; }
    </style>
</head>
<body>
    <div class="header">
        <h1>���上明細 照合チェック</h1>
        <div class="sub">売上申請（Excel）と請求データ（CSV）の自動照合システム</div>
    </div>

    <div class="container">
        <div class="actions">
            <a href="/run" class="btn btn-primary" id="runBtn"
               onclick="this.textContent='チェック実行中...'; this.classList.add('btn-loading');">
                照合チェック実行
            </a>
            {% if results %}
            <a href="/download" class="btn btn-success">CSV ダウンロード</a>
            {% endif %}
            <span class="timestamp">
                {% if timestamp %}最終実行: {{ timestamp }}{% endif %}
            </span>
        </div>

        {% if summaries %}
        <div class="summary-grid">
            {% for s in summaries %}
            <div class="summary-card">
                <h3>{{ s.school }} — {{ s.month }}</h3>
                <div class="summary-row">
                    <span class="label">Excel生徒数</span>
                    <span class="value">{{ s.total }}人</span>
                </div>
                <div class="summary-row">
                    <span class="label">完全一致 (OK)</span>
                    <span class="value ok">{{ s.ok }}人</span>
                </div>
                <div class="summary-row">
                    <span class="label">列配分違い（合計一致）</span>
                    <span class="value warn">{{ s.col_only }}人</span>
                </div>
                <div class="summary-row">
                    <span class="label">合計不一致（要確認）</span>
                    <span class="value critical">{{ s.total_diff }}人</span>
                </div>
                <div class="summary-row">
                    <span class="label">CSV未登録</span>
                    <span class="value">{{ s.not_in_csv }}人</span>
                </div>
            </div>
            {% endfor %}
        </div>
        {% endif %}

        {% if results %}
        <div class="filter-bar">
            <label>校舎:</label>
            <select id="filterSchool" onchange="filterTable()">
                <option value="">すべて</option>
                {% for s in school_names %}<option>{{ s }}</option>{% endfor %}
            </select>
            <label>月:</label>
            <select id="filterMonth" onchange="filterTable()">
                <option value="">すべて</option>
                {% for m in month_names %}<option>{{ m }}</option>{% endfor %}
            </select>
        </div>

        <div class="tabs">
            <button class="tab active" onclick="showTab('critical', this)">
                合計不一致（{{ count_critical }}件）
            </button>
            <button class="tab tab-csv" onclick="showTab('not_csv', this)">
                CSV未登録（{{ count_not_csv }}件）
            </button>
            <button class="tab tab-col" onclick="showTab('col_only', this)">
                列配分違い（{{ count_col_only }}件）
            </button>
        </div>

        <div class="table-wrap">
            <!-- 合計不一致 -->
            <table id="table-critical">
                <thead>
                    <tr>
                        <th>校舎</th><th>月</th><th>生徒ID</th><th>生徒名</th>
                        <th>行</th><th>項目</th>
                        <th>Excel金額</th><th>CSV金額</th><th>差額</th><th>合計差額</th>
                    </tr>
                </thead>
                <tbody>
                {% for r in critical_rows %}
                    <tr data-school="{{ r.school }}" data-month="{{ r.month }}">
                        <td>{{ r.school }}</td>
                        <td>{{ r.month }}</td>
                        <td>{{ r.sid }}</td>
                        <td>{{ r.name }}</td>
                        <td class="num">{{ r.row }}</td>
                        <td>{{ r.item }}</td>
                        <td class="num">{{ r.excel }}</td>
                        <td class="num">{{ r.csv }}</td>
                        <td class="num {{ 'positive' if r.diff_val > 0 else 'negative' if r.diff_val < 0 else '' }}">{{ r.diff }}</td>
                        <td class="num {{ 'positive' if r.total_diff_val > 0 else 'negative' if r.total_diff_val < 0 else '' }}">{{ r.total_diff }}</td>
                    </tr>
                {% endfor %}
                {% if not critical_rows %}
                    <tr><td colspan="10" class="empty">合計不一致はありません</td></tr>
                {% endif %}
                </tbody>
            </table>

            <!-- CSV未登録 -->
            <table id="table-not_csv" style="display:none;">
                <thead>
                    <tr>
                        <th>校舎</th><th>月</th><th>生徒ID</th><th>生徒名</th>
                        <th>行</th><th>Excel売��</th>
                    </tr>
                </thead>
                <tbody>
                {% for r in not_csv_rows %}
                    <tr data-school="{{ r.school }}" data-month="{{ r.month }}">
                        <td>{{ r.school }}</td>
                        <td>{{ r.month }}</td>
                        <td>{{ r.sid }}</td>
                        <td>{{ r.name }}</td>
                        <td class="num">{{ r.row }}</td>
                        <td class="num">{{ r.excel }}</td>
                    </tr>
                {% endfor %}
                {% if not not_csv_rows %}
                    <tr><td colspan="6" class="empty">CSV未登録はありません</td></tr>
                {% endif %}
                </tbody>
            </table>

            <!-- 列配分違い -->
            <table id="table-col_only" style="display:none;">
                <thead>
                    <tr>
                        <th>校舎</th><th>月</th><th>生徒ID</th><th>生徒名</th>
                        <th>行</th><th>項目</th>
                        <th>Excel金額</th><th>CSV金額</th><th>差額</th>
                    </tr>
                </thead>
                <tbody>
                {% for r in col_only_rows %}
                    <tr data-school="{{ r.school }}" data-month="{{ r.month }}">
                        <td>{{ r.school }}</td>
                        <td>{{ r.month }}</td>
                        <td>{{ r.sid }}</td>
                        <td>{{ r.name }}</td>
                        <td class="num">{{ r.row }}</td>
                        <td>{{ r.item }}</td>
                        <td class="num">{{ r.excel }}</td>
                        <td class="num">{{ r.csv }}</td>
                        <td class="num {{ 'positive' if r.diff_val > 0 else 'negative' if r.diff_val < 0 else '' }}">{{ r.diff }}</td>
                    </tr>
                {% endfor %}
                {% if not col_only_rows %}
                    <tr><td colspan="9" class="empty">列配分違いはありません</td></tr>
                {% endif %}
                </tbody>
            </table>
        </div>
        {% elif not summaries %}
        <div class="empty" style="margin-top: 60px;">
            <p style="font-size: 16px; margin-bottom: 8px;">「照合チェック実行」をクリックしてください</p>
            <p>売上Excelと請求CSVを自動検出して照合します</p>
        </div>
        {% endif %}
    </div>

    <script>
    function showTab(name, btn) {
        document.querySelectorAll('table[id^="table-"]').forEach(t => t.style.display = 'none');
        document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
        document.getElementById('table-' + name).style.display = '';
        btn.classList.add('active');
        filterTable();
    }

    function filterTable() {
        const school = document.getElementById('filterSchool').value;
        const month = document.getElementById('filterMonth').value;
        document.querySelectorAll('.table-wrap table:not([style*="display: none"]) tbody tr').forEach(tr => {
            if (tr.querySelector('.empty')) return;
            const s = tr.getAttribute('data-school') || '';
            const m = tr.getAttribute('data-month') || '';
            const show = (!school || s === school) && (!month || m === month);
            tr.style.display = show ? '' : 'none';
        });
    }
    </script>
</body>
</html>
"""


# ====================================================================
# ルーティング
# ====================================================================

# グローバルに結果を保持（シンプルなインメモリキャッシュ）
_cache: dict = {"results": None, "summaries": None, "timestamp": None}


def _format_number(val: float) -> str:
    if val == 0:
        return "0"
    return f"{val:,.0f}"


def _build_template_data() -> dict:
    results = _cache["results"]
    summaries = _cache["summaries"]
    timestamp = _cache["timestamp"]

    if results is None:
        return {"results": None, "summaries": None, "timestamp": None}

    critical_rows = []
    not_csv_rows = []
    col_only_rows = []

    for r in results:
        if r.result_type == "NOT_IN_CSV":
            not_csv_rows.append({
                "school": r.school,
                "month": r.month_label,
                "sid": r.sid,
                "name": r.name,
                "row": r.row,
                "excel": _format_number(r.excel_total),
            })
        elif r.diffs:
            total_diff = r.excel_total - r.csv_total
            total_match = abs(total_diff) < 1
            target = col_only_rows if total_match else critical_rows

            for col, ev, cv, diff in r.diffs:
                disp = COL_DISPLAY.get(col, col)
                target.append({
                    "school": r.school,
                    "month": r.month_label,
                    "sid": r.sid,
                    "name": r.name,
                    "row": r.row,
                    "item": disp,
                    "excel": _format_number(ev),
                    "csv": _format_number(cv),
                    "diff": _format_number(diff),
                    "diff_val": diff,
                    "total_diff": _format_number(total_diff),
                    "total_diff_val": total_diff,
                })

    school_names = sorted(set(r.school for r in results))
    month_names = sorted(set(r.month_label for r in results))

    return {
        "results": results,
        "summaries": summaries,
        "timestamp": timestamp,
        "critical_rows": critical_rows,
        "not_csv_rows": not_csv_rows,
        "col_only_rows": col_only_rows,
        "count_critical": len(critical_rows),
        "count_not_csv": len(not_csv_rows),
        "count_col_only": len(col_only_rows),
        "school_names": school_names,
        "month_names": month_names,
    }


@app.route("/")
def index():
    data = _build_template_data()
    return render_template_string(HTML_TEMPLATE, **data)


@app.route("/run")
def run():
    results, summaries = run_all_checks()
    _cache["results"] = results
    _cache["summaries"] = summaries
    _cache["timestamp"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    return redirect(url_for("index"))


@app.route("/download")
def download():
    if _cache["results"] is None:
        return redirect(url_for("index"))

    csv_bytes = results_to_csv_bytes(_cache["results"])
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"照合結果_{timestamp}.csv"

    return send_file(
        io.BytesIO(csv_bytes),
        mimetype="text/csv",
        as_attachment=True,
        download_name=filename,
    )


if __name__ == "__main__":
    print("=" * 60)
    print("  売上照合チェック Webアプリ")
    print("  ブラウザで http://localhost:3006 にアクセス")
    print("  社内LAN: http://<このPCのIP>:3006")
    print("  停止: Ctrl+C")
    print("=" * 60)
    app.run(host="0.0.0.0", port=3006, debug=False, load_dotenv=False)
