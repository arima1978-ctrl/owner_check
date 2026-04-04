# -*- coding: utf-8 -*-
"""
売上照合チェック Webアプリ

事務員がブラウザからアクセスし、照合結果を確認・CSVダウンロードできる。
社内ネットワークで http://<サーバーIP>:3006 でアクセス。
"""

import io
import csv
import os
import shutil
from datetime import datetime
from pathlib import Path

from flask import (
    Flask, render_template_string, send_file, redirect,
    url_for, request, flash,
)

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
    _preload_all_billings,
    _compute_monthly_billing,
    read_excel_sales,
    aggregate_csv_for_student,
    compare_student,
    COL_DISPLAY,
)

app = Flask(__name__)
app.secret_key = "owner_check_secret_key"

UPLOAD_SALES_DIR = Path("C:/Users/USER/Documents")
UPLOAD_CSV_DIR = Path("Y:/_★20170701作業用/100 真由美/野田より/◆日にち別お仕事◆")


# ====================================================================
# 照合ロジック（Web用）
# ====================================================================

def run_check_silent(
    school_name: str,
    month_label: str,
    csv_paths: list[str],
    csv_paths_with_labels: list[tuple[str, str]],
    excel_path: str,
    sheet_index: int = 2,
    school_brands: dict[str, set[str]] | None = None,
) -> tuple[list[CheckResult], dict]:
    """コンソール出力なしで照合チェック"""
    billing = _load_billing_by_priority(csv_paths)
    sales = read_excel_sales(excel_path, sheet_index)

    month_col_labels = [label for _, label in csv_paths_with_labels]
    all_billings = _preload_all_billings(csv_paths_with_labels)

    ok_count = 0
    col_only_count = 0
    total_diff_count = 0
    not_in_csv_count = 0
    no_billing_count = 0
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
                ))
        else:
            ok_count += 1

    summary = {
        "total": len(sales),
        "ok": ok_count,
        "col_only": col_only_count,
        "total_diff": total_diff_count,
        "not_in_csv": not_in_csv_count,
        "no_billing": no_billing_count,
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
            csv_paths = [pair.csv_path]
            adjacent = _get_adjacent_months(pair.year, pair.month, 2)
            for ym in adjacent:
                if ym in csv_files and csv_files[ym] != pair.csv_path:
                    csv_paths.append(csv_files[ym])

            csv_paths_with_labels = []
            for ym in adjacent:
                if ym in csv_files:
                    csv_paths_with_labels.append(
                        (csv_files[ym], f"{ym[1]}月")
                    )

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
                csv_paths_with_labels=csv_paths_with_labels,
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
    """結果をCSVバイト列に変換（月別引落額付き）"""
    all_month_cols = []
    for r in results:
        for mc in r.month_columns:
            if mc not in all_month_cols:
                all_month_cols.append(mc)

    output = io.StringIO()
    writer = csv.writer(output)
    header = ["校舎", "月", "重要度", "種別", "生徒ID", "生徒名",
              "行番号", "項目", "売上"]
    for mc in all_month_cols:
        header.append(f"{mc}引落")
    header.extend(["差額", "合計差額"])
    writer.writerow(header)

    for r in results:
        if r.result_type == "NOT_IN_CSV":
            row_data = [r.school, r.month_label, "★要確認", "月謝未計上",
                        r.sid, r.name, r.row, "売上合計", r.excel_total]
            row_data.extend([0] * len(all_month_cols))
            row_data.extend([r.excel_total, r.excel_total])
            writer.writerow(row_data)
        elif r.diffs:
            total_diff = r.excel_total - r.csv_total
            total_match = abs(total_diff) < 1
            if r.result_type == "NO_BILLING":
                severity, kind = "★売上あり請求なし", "請求漏れ"
            elif total_match:
                severity, kind = "列配分違い", "合計一致"
            else:
                severity, kind = "★要確認", "合計不一致"

            for col, ev, cv, diff in r.diffs:
                disp = COL_DISPLAY.get(col, col)
                row_data = [r.school, r.month_label, severity, kind,
                            r.sid, r.name, r.row, disp, ev]
                for mc in all_month_cols:
                    monthly_agg = r.monthly_billing.get(mc, {})
                    row_data.append(monthly_agg.get(col, 0))
                row_data.extend([diff, total_diff])
                writer.writerow(row_data)

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
        .header { background: #1a5276; color: white; padding: 16px 24px; display: flex;
                  justify-content: space-between; align-items: center; }
        .header h1 { font-size: 20px; font-weight: 500; }
        .header .sub { font-size: 12px; color: #aed6f1; margin-top: 4px; }
        .header-right a { color: #aed6f1; font-size: 13px; text-decoration: none; margin-left: 16px; }
        .header-right a:hover { color: white; }
        .container { max-width: 1600px; margin: 0 auto; padding: 20px; }

        .actions { display: flex; gap: 12px; margin-bottom: 20px; align-items: center; flex-wrap: wrap; }
        .btn { display: inline-block; padding: 10px 20px; border: none; border-radius: 6px;
               font-size: 14px; cursor: pointer; text-decoration: none; font-weight: 500; }
        .btn-primary { background: #2980b9; color: white; }
        .btn-primary:hover { background: #2471a3; }
        .btn-success { background: #27ae60; color: white; }
        .btn-success:hover { background: #229954; }
        .btn-loading { background: #95a5a6; color: white; pointer-events: none; }

        .summary-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(340px, 1fr));
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
        .value.nobill { color: #8e44ad; }

        .tabs { display: flex; gap: 4px; margin-bottom: 0; flex-wrap: wrap; }
        .tab { padding: 10px 16px; background: #ddd; border: none; border-radius: 8px 8px 0 0;
               cursor: pointer; font-size: 13px; font-weight: 500; }
        .tab.active { background: white; font-weight: 700; }

        .table-wrap { background: white; border-radius: 0 8px 8px 8px; box-shadow: 0 1px 3px rgba(0,0,0,0.1);
                      overflow-x: auto; max-height: 70vh; overflow-y: auto; }
        table { width: 100%; border-collapse: collapse; font-size: 12px; }
        th { background: #f8f9fa; padding: 8px 10px; text-align: left; font-weight: 600;
             border-bottom: 2px solid #dee2e6; white-space: nowrap; position: sticky; top: 0; z-index: 1; }
        td { padding: 6px 10px; border-bottom: 1px solid #eee; white-space: nowrap; }
        tr:hover { background: #f0f7fc; }
        .num { text-align: right; font-family: 'Consolas', monospace; }
        .positive { color: #e74c3c; }
        .negative { color: #2980b9; }
        .month-col { background: #fafafa; }
        .month-col.has-value { background: #e8f8f5; font-weight: 600; }

        .filter-bar { display: flex; gap: 12px; margin-bottom: 16px; align-items: center; flex-wrap: wrap; }
        .filter-bar label { font-size: 13px; font-weight: 500; }
        .filter-bar select { padding: 6px 10px; border: 1px solid #ccc; border-radius: 4px; font-size: 13px; }

        .empty { text-align: center; padding: 40px; color: #999; }
        .timestamp { font-size: 12px; color: #999; }
        .flash { background: #d4edda; color: #155724; padding: 10px 16px; border-radius: 6px;
                 margin-bottom: 16px; font-size: 13px; }
        .flash-error { background: #f8d7da; color: #721c24; }

        .upload-section { background: white; border-radius: 8px; padding: 20px; margin-bottom: 20px;
                          box-shadow: 0 1px 3px rgba(0,0,0,0.1); }
        .upload-section h3 { font-size: 15px; color: #1a5276; margin-bottom: 12px; }
        .upload-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; }
        .upload-box { border: 2px dashed #ccc; border-radius: 8px; padding: 16px; text-align: center; }
        .upload-box h4 { font-size: 14px; margin-bottom: 8px; }
        .upload-box p { font-size: 12px; color: #888; margin-bottom: 8px; }
        .upload-box input[type=file] { font-size: 12px; }
        .upload-box .btn { margin-top: 8px; font-size: 12px; padding: 6px 16px; }
    </style>
</head>
<body>
    <div class="header">
        <div>
            <h1>売上照合チェック</h1>
            <div class="sub">売上申請と請求データの自動照合 / 前後2ヶ月引落参照 / 校舎フィルタ適用</div>
        </div>
        <div class="header-right">
            <a href="/upload">データ登録</a>
        </div>
    </div>

    <div class="container">
        {% for msg in get_flashed_messages() %}
        <div class="flash">{{ msg }}</div>
        {% endfor %}

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
                <h3>{{ s.school }} ― {{ s.month }}</h3>
                <div class="summary-row">
                    <span class="label">生徒数</span>
                    <span class="value">{{ s.total }}人</span>
                </div>
                <div class="summary-row">
                    <span class="label">完全一致 (OK)</span>
                    <span class="value ok">{{ s.ok }}人</span>
                </div>
                <div class="summary-row">
                    <span class="label">その他の差異</span>
                    <span class="value warn">{{ s.col_only }}人</span>
                </div>
                <div class="summary-row">
                    <span class="label">売上あり請求なし</span>
                    <span class="value nobill">{{ s.no_billing }}人</span>
                </div>
                <div class="summary-row">
                    <span class="label">月謝未計上</span>
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
            <button class="tab active" id="tab-no_billing" data-tab="no_billing" onclick="showTab('no_billing', this)" style="color:#8e44ad">
                売上あり請求なし（<span class="tab-count">{{ count_no_billing }}</span>件）
            </button>
            <button class="tab" id="tab-not_csv" data-tab="not_csv" onclick="showTab('not_csv', this)" style="color:#7f8c8d">
                月謝未計上（<span class="tab-count">{{ count_not_csv }}</span>件）
            </button>
            <button class="tab" id="tab-col_only" data-tab="col_only" onclick="showTab('col_only', this)" style="color:#e67e22">
                その他の差異（<span class="tab-count">{{ count_col_only }}</span>件）
            </button>
        </div>

        <div class="table-wrap">
            {% for tab_name, tab_rows, tab_type in [
                ('no_billing', no_billing_rows, 'full'),
                ('col_only', col_only_rows, 'full'),
            ] %}
            <table id="table-{{ tab_name }}" {% if not loop.first %}style="display:none;"{% endif %}>
                <thead>
                    <tr>
                        <th>校舎</th><th>月</th><th>生徒ID</th><th>生徒名</th>
                        <th>行</th><th>項目</th><th>売上</th>
                        {% for mc in all_month_cols %}<th>{{ mc }}引落</th>{% endfor %}
                        <th>差額</th><th>合計差額</th>
                    </tr>
                </thead>
                <tbody>
                {% for r in tab_rows %}
                    <tr data-school="{{ r.school }}" data-month="{{ r.month }}">
                        <td>{{ r.school }}</td>
                        <td>{{ r.month }}</td>
                        <td>{{ r.sid }}</td>
                        <td>{{ r.name }}</td>
                        <td class="num">{{ r.row }}</td>
                        <td>{{ r.item }}</td>
                        <td class="num">{{ r.excel }}</td>
                        {% for mv in r.month_vals %}
                        <td class="num month-col {{ 'has-value' if mv != '0' else '' }}">{{ mv }}</td>
                        {% endfor %}
                        <td class="num {{ 'positive' if r.diff_val > 0 else 'negative' if r.diff_val < 0 else '' }}">{{ r.diff }}</td>
                        <td class="num {{ 'positive' if r.total_diff_val > 0 else 'negative' if r.total_diff_val < 0 else '' }}">{{ r.total_diff }}</td>
                    </tr>
                {% endfor %}
                {% if not tab_rows %}
                    <tr><td colspan="{{ 9 + all_month_cols|length }}" class="empty">該当なし</td></tr>
                {% endif %}
                </tbody>
            </table>
            {% endfor %}

            <table id="table-not_csv" style="display:none;">
                <thead>
                    <tr>
                        <th>校舎</th><th>月</th><th>生徒ID</th><th>生徒名</th>
                        <th>行</th><th>売上</th>
                    </tr>
                </thead>
                <tbody>
                {% for r in not_csv_rows %}
                    <tr data-school="{{ r.school }}" data-month="{{ r.month }}">
                        <td>{{ r.school }}</td><td>{{ r.month }}</td>
                        <td>{{ r.sid }}</td><td>{{ r.name }}</td>
                        <td class="num">{{ r.row }}</td>
                        <td class="num">{{ r.excel }}</td>
                    </tr>
                {% endfor %}
                {% if not not_csv_rows %}
                    <tr><td colspan="6" class="empty">該当なし</td></tr>
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

        // 全テーブルにフィルタ適用 & 件数カウント
        const tabNames = ['no_billing', 'not_csv', 'col_only'];
        tabNames.forEach(function(name) {
            const table = document.getElementById('table-' + name);
            if (!table) return;
            let count = 0;
            table.querySelectorAll('tbody tr').forEach(function(tr) {
                if (tr.querySelector('.empty')) return;
                const s = tr.getAttribute('data-school') || '';
                const m = tr.getAttribute('data-month') || '';
                const show = (!school || s === school) && (!month || m === month);
                tr.style.display = show ? '' : 'none';
                if (show) count++;
            });
            // タブの件数を更新
            const tab = document.getElementById('tab-' + name);
            if (tab) {
                const span = tab.querySelector('.tab-count');
                if (span) span.textContent = count;
            }
        });
    }
    </script>
</body>
</html>
"""

UPLOAD_TEMPLATE = """
<!DOCTYPE html>
<html lang="ja">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>データ登録 - 売上照合チェック</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: 'Yu Gothic', 'Meiryo', sans-serif; background: #f5f5f5; color: #333; }
        .header { background: #1a5276; color: white; padding: 16px 24px; display: flex;
                  justify-content: space-between; align-items: center; }
        .header h1 { font-size: 20px; font-weight: 500; }
        .header-right a { color: #aed6f1; font-size: 13px; text-decoration: none; }
        .header-right a:hover { color: white; }
        .container { max-width: 960px; margin: 0 auto; padding: 20px; }
        .section { background: white; border-radius: 8px; padding: 24px; margin-bottom: 20px;
                   box-shadow: 0 1px 3px rgba(0,0,0,0.1); }
        .section h3 { font-size: 16px; color: #1a5276; margin-bottom: 4px; }
        .section .desc { font-size: 12px; color: #888; margin-bottom: 16px; }
        .form-row { margin-bottom: 12px; }
        .form-row label { display: block; font-size: 13px; font-weight: 500; margin-bottom: 4px; }
        .form-row select, .form-row input[type=file], .form-row input[type=text] {
            font-size: 13px; padding: 6px 10px; border: 1px solid #ccc; border-radius: 4px; }
        .form-row input[type=text] { width: 300px; }
        .form-inline { display: flex; gap: 10px; align-items: end; flex-wrap: wrap; }
        .form-inline .form-row { margin-bottom: 0; }
        .btn { display: inline-block; padding: 10px 20px; border: none; border-radius: 6px;
               font-size: 14px; cursor: pointer; font-weight: 500; }
        .btn-sm { padding: 7px 14px; font-size: 12px; }
        .btn-primary { background: #2980b9; color: white; }
        .btn-primary:hover { background: #2471a3; }
        .btn-outline { background: white; color: #2980b9; border: 1px solid #2980b9; }
        .btn-outline:hover { background: #eaf2f8; }
        .flash { background: #d4edda; color: #155724; padding: 10px 16px; border-radius: 6px;
                 margin-bottom: 16px; font-size: 13px; }
        .file-list { margin-top: 12px; font-size: 12px; color: #555; }
        .file-list table { border-collapse: collapse; width: 100%; }
        .file-list th { text-align: left; padding: 4px 8px; background: #f8f9fa; font-weight: 600; }
        .file-list td { padding: 4px 8px; border-bottom: 1px solid #eee; }
        .badge { display: inline-block; padding: 1px 6px; border-radius: 3px; font-size: 11px; }
        .badge-ok { background: #d4edda; color: #155724; }
    </style>
</head>
<body>
    <div class="header">
        <h1>データ登録</h1>
        <div class="header-right">
            <a href="/">← 照合チェックに戻る</a>
        </div>
    </div>
    <div class="container">
        {% for msg in get_flashed_messages() %}
        <div class="flash">{{ msg }}</div>
        {% endfor %}

        <!-- 売上明細アップロード -->
        <div class="section">
            <h3>売上明細の登録</h3>
            <p class="desc">校舎の売上明細書（Excel）を登録します。校舎フォルダに自動配置されます。</p>
            <form action="/upload/sales" method="post" enctype="multipart/form-data">
                <div class="form-inline">
                    <div class="form-row">
                        <label>校舎:</label>
                        <select name="school">
                            {% for s in schools %}<option value="{{ s.name }}">{{ s.name }}</option>{% endfor %}
                        </select>
                    </div>
                    <div class="form-row">
                        <label>売上明細Excel:</label>
                        <input type="file" name="files" accept=".xlsm,.xlsx" multiple required>
                    </div>
                    <button type="submit" class="btn btn-primary">登録</button>
                </div>
            </form>

            {% if sales_files %}
            <div class="file-list">
                <p style="font-weight:600; margin-bottom:4px;">登録済み売上ファイル:</p>
                <table>
                    <tr><th>校舎</th><th>ファイル</th></tr>
                    {% for sf in sales_files %}
                    <tr><td>{{ sf.school }}</td><td>{{ sf.name }}</td></tr>
                    {% endfor %}
                </table>
            </div>
            {% endif %}
        </div>

        <!-- 引落CSV アップロード -->
        <div class="section">
            <h3>引落結果（請求CSV）の登録</h3>
            <p class="desc">請求CSVファイルを登録します。対象月を選択すると送信日フォルダが自動決定されます。</p>
            <form action="/upload/csv" method="post" enctype="multipart/form-data">
                <div class="form-inline">
                    <div class="form-row">
                        <label>対象月（引落月）:</label>
                        <select name="year">
                            {% for y in csv_years %}<option value="{{ y }}" {{ 'selected' if y == current_year else '' }}>{{ y }}年</option>{% endfor %}
                        </select>
                        <select name="month">
                            {% for m in range(1,13) %}<option value="{{ m }}" {{ 'selected' if m == current_month else '' }}>{{ m }}月</option>{% endfor %}
                        </select>
                    </div>
                    <div class="form-row">
                        <label>請求CSV:</label>
                        <input type="file" name="files" accept=".csv" multiple required>
                    </div>
                    <button type="submit" class="btn btn-primary">登録</button>
                </div>
            </form>

            {% if csv_months %}
            <div class="file-list">
                <p style="font-weight:600; margin-bottom:4px;">登録済み請求CSV:</p>
                <table>
                    <tr><th>対象月</th><th>ファイル</th></tr>
                    {% for cm in csv_months %}
                    <tr><td>{{ cm.month }}</td><td>{{ cm.name }}</td></tr>
                    {% endfor %}
                </table>
            </div>
            {% endif %}
        </div>

        <!-- 新校舎追加 -->
        <div class="section">
            <h3>新しい校舎を追加</h3>
            <p class="desc">照合対象の校舎を追加します。売上フォルダが自動作成されます。</p>
            <form action="/upload/add_school" method="post">
                <div class="form-inline">
                    <div class="form-row">
                        <label>校舎名:</label>
                        <input type="text" name="school_name" placeholder="例: 千種" required>
                    </div>
                    <button type="submit" class="btn btn-outline btn-sm">追加</button>
                </div>
            </form>
        </div>
    </div>
</body>
</html>
"""


# ====================================================================
# ルーティング
# ====================================================================

_cache: dict = {"results": None, "summaries": None, "timestamp": None}


def _format_number(val: float) -> str:
    if val == 0:
        return "0"
    return f"{val:,.0f}"


def _build_row(r, col, ev, cv, diff, total_diff, all_month_cols):
    """1行分のテンプレートデータを構築"""
    disp = COL_DISPLAY.get(col, col)
    month_vals = []
    for mc in all_month_cols:
        monthly_agg = r.monthly_billing.get(mc, {})
        v = monthly_agg.get(col, 0)
        month_vals.append(_format_number(v))

    return {
        "school": r.school,
        "month": r.month_label,
        "sid": r.sid,
        "name": r.name,
        "row": r.row,
        "item": disp,
        "excel": _format_number(ev),
        "month_vals": month_vals,
        "diff": _format_number(diff),
        "diff_val": diff,
        "total_diff": _format_number(total_diff),
        "total_diff_val": total_diff,
    }


def _build_template_data() -> dict:
    results = _cache["results"]
    summaries = _cache["summaries"]
    timestamp = _cache["timestamp"]

    if results is None:
        return {"results": None, "summaries": None, "timestamp": None,
                "all_month_cols": []}

    # 全結果から月カラム収集
    all_month_cols = []
    for r in results:
        for mc in r.month_columns:
            if mc not in all_month_cols:
                all_month_cols.append(mc)

    no_billing_rows = []
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

            if r.result_type == "NO_BILLING":
                target = no_billing_rows
            else:
                target = col_only_rows

            for col, ev, cv, diff in r.diffs:
                target.append(
                    _build_row(r, col, ev, cv, diff, total_diff, all_month_cols)
                )

    school_names = sorted(set(r.school for r in results))
    month_names = sorted(set(r.month_label for r in results))

    return {
        "results": results,
        "summaries": summaries,
        "timestamp": timestamp,
        "all_month_cols": all_month_cols,
        "no_billing_rows": no_billing_rows,
        "not_csv_rows": not_csv_rows,
        "col_only_rows": col_only_rows,
        "count_no_billing": len(no_billing_rows),
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
    return send_file(
        io.BytesIO(csv_bytes),
        mimetype="text/csv",
        as_attachment=True,
        download_name=f"照合結果_{timestamp}.csv",
    )


@app.route("/upload")
def upload_page():
    config = load_config()
    schools = []
    for s in config["schools"]:
        schools.append({"name": s["name"], "excel_dir": s["excel_dir"]})

    # 登録済み売上ファイル一覧
    sales_files = []
    for s in config["schools"]:
        d = Path(s["excel_dir"])
        if d.exists():
            for f in sorted(d.glob("*.xls*")):
                sales_files.append({"school": s["name"], "name": f.name})

    # 登録済み請求CSV一覧
    csv_files = discover_csv_files(config["csv_base_dir"])
    csv_months = []
    for (y, m), path in sorted(csv_files.items()):
        csv_months.append({
            "month": f"{y}年{m}月",
            "name": os.path.basename(path),
        })

    now = datetime.now()
    csv_years = list(range(now.year - 1, now.year + 2))

    return render_template_string(
        UPLOAD_TEMPLATE,
        schools=schools,
        sales_files=sales_files,
        csv_months=csv_months,
        csv_years=csv_years,
        current_year=now.year,
        current_month=now.month,
    )


@app.route("/upload/sales", methods=["POST"])
def upload_sales():
    config = load_config()
    school_name = request.form.get("school")
    files = request.files.getlist("files")
    if not files or not school_name:
        flash("ファイルまたは校舎が選択されていません")
        return redirect(url_for("upload_page"))

    school_cfg = None
    for s in config["schools"]:
        if s["name"] == school_name:
            school_cfg = s
            break
    if not school_cfg:
        flash(f"校舎 '{school_name}' が見つかりません")
        return redirect(url_for("upload_page"))

    dest_dir = Path(school_cfg["excel_dir"])
    dest_dir.mkdir(parents=True, exist_ok=True)
    saved = []
    for file in files:
        if file.filename:
            dest = dest_dir / file.filename
            file.save(str(dest))
            saved.append(file.filename)
    flash(f"売上明細を{len(saved)}件登録しました（{school_name}）: {', '.join(saved)}")
    return redirect(url_for("upload_page"))


@app.route("/upload/csv", methods=["POST"])
def upload_csv():
    files = request.files.getlist("files")
    year = request.form.get("year", "")
    month = request.form.get("month", "")
    if not files or not year or not month:
        flash("対象月またはファイルが選択されていません")
        return redirect(url_for("upload_page"))

    # 対象月 → 送信日フォルダ名を自動生成（対象月-1ヶ月の日付）
    y, m = int(year), int(month)
    prev_m = m - 1
    prev_y = y
    if prev_m < 1:
        prev_m = 12
        prev_y -= 1
    today = datetime.now().strftime("%d")
    folder_name = f"{prev_y}{prev_m:02d}{today}_送信日"

    dest_dir = UPLOAD_CSV_DIR / folder_name
    dest_dir.mkdir(parents=True, exist_ok=True)
    saved = []
    for file in files:
        if file.filename:
            dest = dest_dir / file.filename
            file.save(str(dest))
            saved.append(file.filename)
    flash(f"{y}年{m}月分の請求CSVを{len(saved)}件登録しました（{folder_name}/）: {', '.join(saved)}")
    return redirect(url_for("upload_page"))


@app.route("/upload/add_school", methods=["POST"])
def add_school():
    school_name = request.form.get("school_name", "").strip()
    if not school_name:
        flash("校舎名を入力してください")
        return redirect(url_for("upload_page"))

    config_path = str(Path(__file__).parent / "config.yaml")
    config = load_config(config_path)

    # 既存チェック
    for s in config["schools"]:
        if s["name"] == school_name:
            flash(f"校舎 '{school_name}' は既に登録されています")
            return redirect(url_for("upload_page"))

    # config.yaml に追加
    new_school = {
        "name": school_name,
        "excel_dir": f"C:/Users/USER/Documents/{school_name}売上",
        "excel_pattern": "{name}売上明細書({year}.{month}).xlsm",
        "sheet_index": 2,
        "school_keywords": [school_name],
    }
    config["schools"].append(new_school)

    import yaml
    with open(config_path, "w", encoding="utf-8") as f:
        yaml.dump(config, f, allow_unicode=True, default_flow_style=False)

    # フォルダ作成
    Path(new_school["excel_dir"]).mkdir(parents=True, exist_ok=True)

    flash(f"校舎 '{school_name}' を追加しました。売上フォルダ: {new_school['excel_dir']}")
    return redirect(url_for("upload_page"))


if __name__ == "__main__":
    print("=" * 60)
    print("  売上照合チェック Webアプリ")
    print("  ブラウザで http://localhost:3006 にアクセス")
    print("  社内LAN: http://<このPCのIP>:3006")
    print("  停止: Ctrl+C")
    print("=" * 60)
    app.run(host="0.0.0.0", port=3006, debug=False, load_dotenv=False)
