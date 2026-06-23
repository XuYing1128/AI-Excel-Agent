"""Deterministic builders for complex business workbooks with known semantics."""

from __future__ import annotations

import re
from pathlib import Path
from typing import Any

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from .io_utils import ensure_output_path
from .style_library import apply_print_settings, set_reasonable_column_widths


THIN = Side(style="thin", color="B7C2D0")
HEADER_FILL = PatternFill("solid", fgColor="003366")
HEADER_FONT = Font(name="Microsoft YaHei", bold=True, color="FFFFFF")
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")
FORMULA_FILL = PatternFill("solid", fgColor="E2F0D9")
SUBTOTAL_FILL = PatternFill("solid", fgColor="D9EAF7")
TOTAL_FILL = PatternFill("solid", fgColor="B4C6E7")


def can_build_performance_compensation(
    prompt: str,
    plan: dict[str, Any],
) -> bool:
    text = str(prompt or "")
    tables = list(plan.get("inline_tables") or [])
    required = ("绩效评估", "薪酬调整")
    header_sets = [
        {str(item) for item in table.get("columns", [])}
        for table in tables
        if isinstance(table, dict)
    ]
    return (
        all(word in text for word in required)
        and any({"指标", "权重"}.issubset(headers) for headers in header_sets)
        and any({"等级", "最低分", "最高分"}.issubset(headers) for headers in header_sets)
        and any("员工编号" in headers and "当前月薪(元)" in headers for headers in header_sets)
    )


def build_performance_compensation_workbook(
    plan: dict[str, Any],
    prompt: str,
    output: str | Path,
) -> Path:
    """Build a parameter-driven performance and salary adjustment workbook."""

    tables = [dict(item) for item in plan.get("inline_tables", [])]
    weights = _find_table(tables, {"指标", "权重"})
    grades = _find_table(tables, {"等级", "最低分", "最高分"})
    adjustments = _find_table(tables, {"等级", "调整比例"})
    deductions = _find_table(tables, {"扣分值"})
    employees = max(tables, key=lambda item: len(item.get("columns", [])))
    if not all((weights, grades, adjustments, deductions, employees)):
        raise ValueError("绩效薪酬模型缺少必要参数表或员工基础数据。")

    output_path = ensure_output_path(output, "员工绩效评估及薪酬调整表.xlsx")
    wb = Workbook()
    wb.remove(wb.active)
    _write_instructions(wb.create_sheet("说明"), plan)
    parameter_refs = _write_parameter_sheet(
        wb.create_sheet("参数表"),
        weights,
        grades,
        adjustments,
        deductions,
    )
    _write_performance_detail(
        wb.create_sheet("明细表"),
        employees,
        parameter_refs,
        str(plan.get("title") or "员工绩效评估及薪酬调整明细表"),
        prompt,
    )
    wb.save(output_path)
    return output_path


def _write_instructions(ws, plan: dict[str, Any]) -> None:
    ws["A1"] = str(plan.get("title") or "绩效薪酬模型使用说明")
    ws["A1"].font = Font(name="Microsoft YaHei", size=16, bold=True, color="FFFFFF")
    ws["A1"].fill = PatternFill("solid", fgColor="24588A")
    ws.merge_cells("A1:D1")
    rows = [
        ("工作簿结构", "参数表保存权重、等级、调薪和缺勤扣分参数；明细表通过公式引用参数表。"),
        ("输入区域", "明细表浅黄色单元格为原始输入；隐藏辅助列保存缺勤天数和未合并部门。"),
        ("计算区域", "浅绿色单元格为 Excel 公式，包括考勤得分、加权总分、等级、排名和调整后薪资。"),
        ("维护方式", "修改参数表后，使用 Excel、WPS 或 LibreOffice 打开并重算即可更新结果。"),
        ("风险提醒", "绩效、工资和薪酬调整结果需要人工复核。"),
    ]
    for row, (label, value) in enumerate(rows, start=3):
        ws.cell(row, 1, label).font = Font(name="Microsoft YaHei", bold=True)
        ws.cell(row, 2, value)
    set_reasonable_column_widths(ws)
    apply_print_settings(ws)


def _write_parameter_sheet(
    ws,
    weights: dict[str, Any],
    grades: dict[str, Any],
    adjustments: dict[str, Any],
    deductions: dict[str, Any],
) -> dict[str, Any]:
    ws["A1"] = "绩效评估参数表"
    ws.merge_cells("A1:M1")
    ws["A1"].font = Font(name="Microsoft YaHei", size=16, bold=True, color="FFFFFF")
    ws["A1"].fill = PatternFill("solid", fgColor="24588A")
    ws["A1"].alignment = Alignment(horizontal="center")

    weight_rows = _write_block(ws, 3, 1, "指标权重", weights)
    grade_records = sorted(
        grades.get("records", []),
        key=lambda item: _number(item.get("最低分")),
    )
    grade_table = {**grades, "records": grade_records}
    grade_rows = _write_block(ws, 3, 4, "绩效等级判定标准", grade_table)
    adjustment_rows = _write_block(ws, 3, 8, "薪酬调整系数", adjustments)

    ws.cell(3, 11, "缺勤扣分对照表")
    ws.merge_cells(start_row=3, start_column=11, end_row=3, end_column=13)
    _style_section_title(ws.cell(3, 11))
    for col, value in enumerate(["起始缺勤天数", "缺勤天数范围（天）", "扣分值"], 11):
        ws.cell(4, col, value)
        _style_header(ws.cell(4, col))
    deduction_rows = []
    for row, record in enumerate(deductions.get("records", []), start=5):
        range_value = next(
            (
                value
                for key, value in record.items()
                if "缺勤天数" in str(key) or "范围" in str(key)
            ),
            "",
        )
        lower = _range_lower_bound(range_value)
        deduction = next(
            (value for key, value in record.items() if "扣分" in str(key)),
            0,
        )
        ws.cell(row, 11, lower)
        ws.cell(row, 12, range_value)
        ws.cell(row, 13, _number(deduction))
        deduction_rows.append(row)
    _style_block_body(ws, 5, 4 + len(deduction_rows), 11, 13)

    for row in weight_rows:
        ws.cell(row, 2).number_format = "0%"
    for row in adjustment_rows:
        ws.cell(row, 9).number_format = "0%"
    ws.freeze_panes = "A4"
    ws.sheet_view.showGridLines = False
    set_reasonable_column_widths(ws)
    apply_print_settings(ws)

    weight_map = {
        str(ws.cell(row, 1).value): f"'参数表'!$B${row}"
        for row in weight_rows
    }
    weight_values = {
        str(ws.cell(row, 1).value): _number(ws.cell(row, 2).value)
        for row in weight_rows
    }
    return {
        "weights": weight_map,
        "weight_values": weight_values,
        "grade_range": (
            f"'参数表'!$E${min(grade_rows)}:$E${max(grade_rows)}",
            f"'参数表'!$D${min(grade_rows)}:$D${max(grade_rows)}",
        ),
        "adjustment_range": (
            f"'参数表'!$H${min(adjustment_rows)}:$H${max(adjustment_rows)}",
            f"'参数表'!$I${min(adjustment_rows)}:$I${max(adjustment_rows)}",
        ),
        "deduction_range": (
            f"'参数表'!$K${min(deduction_rows)}:$K${max(deduction_rows)}",
            f"'参数表'!$M${min(deduction_rows)}:$M${max(deduction_rows)}",
        ),
        "deduction_rules": [
            (_number(ws.cell(row, 11).value), _number(ws.cell(row, 13).value))
            for row in deduction_rows
        ],
    }


def _write_block(
    ws,
    start_row: int,
    start_col: int,
    title: str,
    table: dict[str, Any],
) -> list[int]:
    columns = [str(item) for item in table.get("columns", [])]
    records = [dict(item) for item in table.get("records", [])]
    end_col = start_col + len(columns) - 1
    ws.cell(start_row, start_col, title)
    if end_col > start_col:
        ws.merge_cells(
            start_row=start_row,
            start_column=start_col,
            end_row=start_row,
            end_column=end_col,
        )
    _style_section_title(ws.cell(start_row, start_col))
    for offset, name in enumerate(columns):
        _style_header(ws.cell(start_row + 1, start_col + offset, name))
    rows: list[int] = []
    for row, record in enumerate(records, start=start_row + 2):
        rows.append(row)
        for offset, name in enumerate(columns):
            ws.cell(row, start_col + offset, record.get(name, ""))
    _style_block_body(ws, start_row + 2, start_row + 1 + len(records), start_col, end_col)
    return rows


def _write_performance_detail(
    ws,
    employees: dict[str, Any],
    refs: dict[str, Any],
    title: str,
    prompt: str,
) -> None:
    del prompt
    source_records = [dict(item) for item in employees.get("records", [])]
    records = sorted(
        source_records,
        key=lambda item: (
            _department_order(str(item.get("部门", ""))),
            -_estimated_score(item, refs),
        ),
    )
    headers = [
        "部门",
        "员工编号",
        "姓名",
        "当前月薪(元)",
        "工作质量",
        "工作效率",
        "团队协作",
        "考勤",
        "加权总分",
        "绩效等级",
        "部门内排名",
        "薪酬调整比例",
        "调整后薪资(元)",
        "备注",
    ]
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(headers))
    ws["A1"] = title
    ws["A1"].font = Font(name="Microsoft YaHei", size=16, bold=True, color="FFFFFF")
    ws["A1"].fill = PatternFill("solid", fgColor="24588A")
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws["A2"] = "需要人工复核：绩效等级、排名和薪酬调整均由公式生成。"
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(headers))

    grouped_headers = [
        ("部门", 1, 1),
        ("员工编号", 2, 2),
        ("姓名", 3, 3),
        ("当前月薪(元)", 4, 4),
        ("指标评分（百分制）", 5, 8),
        ("加权总分", 9, 9),
        ("绩效等级", 10, 10),
        ("部门内排名", 11, 11),
        ("薪酬调整比例", 12, 12),
        ("调整后薪资(元)", 13, 13),
        ("备注", 14, 14),
    ]
    for label, start, end in grouped_headers:
        if start == end and start not in range(5, 9):
            ws.merge_cells(start_row=3, start_column=start, end_row=4, end_column=end)
        elif end > start:
            ws.merge_cells(start_row=3, start_column=start, end_row=3, end_column=end)
        cell = ws.cell(3, start, label)
        _style_header(cell)
    for col, label in enumerate(["工作质量(40%)", "工作效率(30%)", "团队协作(20%)", "考勤(10%)"], start=5):
        _style_header(ws.cell(4, col, label))
    ws.cell(4, 15, "缺勤天数")
    ws.cell(4, 16, "原部门")
    for row in range(3, 5):
        for col in range(1, len(headers) + 1):
            _style_header(ws.cell(row, col))

    data_rows: list[int] = []
    current_row = 5
    for department, group in _group_records(records):
        rows: list[int] = []
        for record in group:
            row = current_row
            rows.append(row)
            data_rows.append(row)
            values = [
                department,
                record.get("员工编号", ""),
                record.get("姓名", ""),
                _number(record.get("当前月薪(元)")),
                _number(record.get("工作质量")),
                _number(record.get("工作效率")),
                _number(record.get("团队协作")),
            ]
            for col, value in enumerate(values, start=1):
                ws.cell(row, col, value)
            ws.cell(row, 15, _number(record.get("缺勤天数")))
            ws.cell(row, 16, department)
            deduction_keys, deduction_values = refs["deduction_range"]
            ws.cell(
                row,
                8,
                f"=IFERROR(10-LOOKUP(O{row},{deduction_keys},{deduction_values}),0)",
            )
            weights = refs["weights"]
            ws.cell(
                row,
                9,
                "=IFERROR("
                f"E{row}*{_weight_ref(weights, '工作质量')}+"
                f"F{row}*{_weight_ref(weights, '工作效率')}+"
                f"G{row}*{_weight_ref(weights, '团队协作')}+"
                f"H{row}*{_weight_ref(weights, '考勤')}"
                ",0)",
            )
            grade_scores, grade_labels = refs["grade_range"]
            ws.cell(
                row,
                10,
                f'=IFERROR(LOOKUP(I{row},{grade_scores},{grade_labels}),"")',
            )
            adjustment_labels, adjustment_values = refs["adjustment_range"]
            ws.cell(
                row,
                12,
                f'=IFERROR(INDEX({adjustment_values},MATCH(J{row},{adjustment_labels},0)),0)',
            )
            ws.cell(row, 13, f"=IFERROR(D{row}*(1+L{row}),0)")
            ws.cell(row, 14, f'=IF(I{row}<60,"⚠️警告：需改进","")')
            current_row += 1

        subtotal_row = current_row
        ws.merge_cells(
            start_row=subtotal_row,
            start_column=1,
            end_row=subtotal_row,
            end_column=3,
        )
        ws.cell(subtotal_row, 1, f"{department}汇总")
        ws.cell(subtotal_row, 4, f"=SUM(D{rows[0]}:D{rows[-1]})")
        ws.cell(
            subtotal_row,
            9,
            f'="平均："&TEXT(AVERAGE(I{rows[0]}:I{rows[-1]}),"0.0")'
            f'&" / 最高："&TEXT(MAX(I{rows[0]}:I{rows[-1]}),"0.0")'
            f'&" / 最低："&TEXT(MIN(I{rows[0]}:I{rows[-1]}),"0.0")',
        )
        ws.cell(subtotal_row, 13, f"=SUM(M{rows[0]}:M{rows[-1]})")
        _style_summary_row(ws, subtotal_row, 14, SUBTOTAL_FILL)
        if len(rows) > 1:
            ws.merge_cells(
                start_row=rows[0],
                start_column=1,
                end_row=rows[-1],
                end_column=1,
            )
        current_row += 1

    last_data_bound = current_row - 1
    for row in data_rows:
        ws.cell(
            row,
            11,
            f"=1+SUMPRODUCT(($P$5:$P${last_data_bound}=P{row})*"
            f"($I$5:$I${last_data_bound}>I{row}))",
        )
    total_row = current_row
    ws.merge_cells(start_row=total_row, start_column=1, end_row=total_row, end_column=3)
    ws.cell(total_row, 1, "总计")
    salary_ranges = ",".join(f"D{row}" for row in data_rows)
    score_ranges = ",".join(f"I{row}" for row in data_rows)
    adjusted_ranges = ",".join(f"M{row}" for row in data_rows)
    ws.cell(total_row, 4, f"=SUM({salary_ranges})")
    ws.cell(
        total_row,
        9,
        f'="平均："&TEXT(AVERAGE({score_ranges}),"0.0")'
        f'&" / 最高："&TEXT(MAX({score_ranges}),"0.0")'
        f'&" / 最低："&TEXT(MIN({score_ranges}),"0.0")',
    )
    ws.cell(total_row, 13, f"=SUM({adjusted_ranges})")
    _style_summary_row(ws, total_row, 14, TOTAL_FILL)

    formula_columns = {8, 9, 10, 11, 12, 13, 14}
    for row in data_rows:
        for col in range(1, 15):
            cell = ws.cell(row, col)
            cell.fill = FORMULA_FILL if col in formula_columns else INPUT_FILL
            cell.border = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.font = Font(name="Microsoft YaHei", size=10)
        ws.cell(row, 4).number_format = '¥#,##0;[Red]-¥#,##0'
        ws.cell(row, 9).number_format = "0.0"
        ws.cell(row, 12).number_format = "0%"
        ws.cell(row, 13).number_format = '¥#,##0;[Red]-¥#,##0'
    first, last = min(data_rows), max(data_rows)
    ws.conditional_formatting.add(
        f"J{first}:J{last}",
        FormulaRule(
            formula=[f'J{first}="A"'],
            fill=PatternFill("solid", fgColor="C6EFCE"),
        ),
    )
    ws.conditional_formatting.add(
        f"J{first}:J{last}",
        FormulaRule(
            formula=[f'OR(J{first}="D",J{first}="E")'],
            fill=PatternFill("solid", fgColor="FFC7CE"),
        ),
    )
    ws.conditional_formatting.add(
        f"N{first}:N{last}",
        FormulaRule(
            formula=[f'N{first}<>""'],
            font=Font(color="C00000", bold=True),
        ),
    )

    ws.column_dimensions["O"].hidden = True
    ws.column_dimensions["P"].hidden = True
    for col in range(1, 15):
        ws.column_dimensions[get_column_letter(col)].width = (
            18 if col in {9, 13, 14} else 14
        )
    ws.freeze_panes = "A5"
    ws.auto_filter.ref = f"B4:N{total_row - 1}"
    ws.sheet_view.showGridLines = False
    ws.print_title_rows = "1:4"
    ws.print_area = f"A1:N{total_row}"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    apply_print_settings(ws)

    # A comparison chart of real input values (names vs current salary) so the
    # workbook satisfies the user's chart request even on this fallback path.
    if data_rows:
        first_data, last_data = min(data_rows), max(data_rows)
        chart = BarChart()
        chart.type = "col"
        chart.title = "各员工当前月薪对比"
        chart.style = 10
        chart.height = 8
        chart.width = 18
        chart.gapWidth = 80
        chart.add_data(
            Reference(ws, min_col=4, min_row=first_data, max_row=last_data),
            titles_from_data=False,
        )
        chart.set_categories(
            Reference(ws, min_col=3, min_row=first_data, max_row=last_data)
        )
        chart.legend = None
        chart.y_axis.numFmt = "#,##0"
        ws.add_chart(chart, f"B{total_row + 2}")


def _find_table(
    tables: list[dict[str, Any]],
    required_headers: set[str],
) -> dict[str, Any] | None:
    for table in tables:
        headers = {str(item) for item in table.get("columns", [])}
        if required_headers.issubset(headers):
            return table
    return None


def _style_header(cell) -> None:
    cell.fill = HEADER_FILL
    cell.font = HEADER_FONT
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def _style_section_title(cell) -> None:
    cell.fill = PatternFill("solid", fgColor="D9EAF7")
    cell.font = Font(name="Microsoft YaHei", bold=True, color="17324D")
    cell.alignment = Alignment(horizontal="center")


def _style_block_body(ws, min_row: int, max_row: int, min_col: int, max_col: int) -> None:
    for row in range(min_row, max_row + 1):
        for col in range(min_col, max_col + 1):
            cell = ws.cell(row, col)
            cell.border = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
            cell.alignment = Alignment(horizontal="center", vertical="center")


def _style_summary_row(ws, row: int, max_col: int, fill) -> None:
    for col in range(1, max_col + 1):
        cell = ws.cell(row, col)
        cell.fill = fill
        cell.font = Font(name="Microsoft YaHei", bold=True, color="17324D")
        cell.border = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.cell(row, 4).number_format = '¥#,##0;[Red]-¥#,##0'
    ws.cell(row, 13).number_format = '¥#,##0;[Red]-¥#,##0'


def _range_lower_bound(value: Any) -> int:
    text = str(value)
    match = re.search(r"\d+", text)
    return int(match.group()) if match else 0


def _number(value: Any) -> float:
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return float(value)
    text = str(value or "").strip().replace(",", "").replace("¥", "").replace("￥", "")
    if text.endswith("%"):
        text = text[:-1]
        try:
            return float(text.lstrip("+")) / 100
        except ValueError:
            return 0.0
    try:
        return float(text)
    except ValueError:
        return 0.0


def _department_order(value: str) -> tuple[int, str]:
    order = {"技术部": 0, "市场部": 1, "行政部": 2}
    return order.get(value, len(order)), value


def _estimated_score(record: dict[str, Any], refs: dict[str, Any]) -> float:
    absence = _number(record.get("缺勤天数"))
    deduction = 0.0
    for lower_bound, value in sorted(refs.get("deduction_rules", [])):
        if absence >= lower_bound:
            deduction = value
    attendance = 10 - deduction
    weights = refs.get("weight_values", {})
    return (
        _number(record.get("工作质量")) * _weight_value(weights, "工作质量")
        + _number(record.get("工作效率")) * _weight_value(weights, "工作效率")
        + _number(record.get("团队协作")) * _weight_value(weights, "团队协作")
        + attendance * _weight_value(weights, "考勤")
    )


def _group_records(records: list[dict[str, Any]]) -> list[tuple[str, list[dict[str, Any]]]]:
    groups: list[tuple[str, list[dict[str, Any]]]] = []
    for record in records:
        department = str(record.get("部门", ""))
        if not groups or groups[-1][0] != department:
            groups.append((department, [record]))
        else:
            groups[-1][1].append(record)
    return groups


def _weight_ref(weights: dict[str, str], name: str) -> str:
    return next(
        (reference for label, reference in weights.items() if name in label),
        "0",
    )


def _weight_value(weights: dict[str, float], name: str) -> float:
    return next(
        (float(value) for label, value in weights.items() if name in label),
        0.0,
    )
