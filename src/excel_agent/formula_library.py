"""Small helpers that generate Excel formula strings.

The functions intentionally return strings with a leading "=" so callers can
write them directly into openpyxl cells.
"""

from __future__ import annotations

import re
from typing import Any, Sequence, Tuple


CellOrRange = str
CriteriaPair = Tuple[CellOrRange, Any]


def strip_equals(formula: str) -> str:
    return formula[1:] if formula.startswith("=") else formula


def ensure_formula(expr: str) -> str:
    return expr if expr.startswith("=") else f"={expr}"


def excel_literal(value: Any) -> str:
    """Return a value formatted for formula criteria arguments."""

    if isinstance(value, (int, float)):
        return str(value)
    text = str(value)
    if text.startswith('"') and text.endswith('"'):
        return text
    if re.match(r"^[$]?[A-Za-z]{1,3}[$]?\d+$", text):
        return text
    if text.startswith((">", "<", "=")):
        return f'"{text}"'
    return f'"{text}"'


def _fallback_literal(value: Any) -> str:
    if isinstance(value, str):
        return value if value.startswith('"') else excel_literal(value)
    return str(value)


def _operand(expr: CellOrRange) -> str:
    text = str(expr)
    if text.startswith("(") and text.endswith(")"):
        return text
    if re.search(r"[+\-*/]", text):
        return f"({text})"
    return text


def excel_sum(range_ref: CellOrRange) -> str:
    return f"=SUM({range_ref})"


def average(range_ref: CellOrRange) -> str:
    return f"=AVERAGE({range_ref})"


def countif(range_ref: CellOrRange, criteria: Any) -> str:
    return f"=COUNTIF({range_ref},{excel_literal(criteria)})"


def countifs(*criteria_pairs: CriteriaPair) -> str:
    parts: list[str] = []
    for range_ref, criteria in criteria_pairs:
        parts.extend([range_ref, excel_literal(criteria)])
    return f"=COUNTIFS({','.join(parts)})"


def sumif(range_ref: CellOrRange, criteria: Any, sum_range: CellOrRange) -> str:
    return f"=SUMIF({range_ref},{excel_literal(criteria)},{sum_range})"


def sumifs(sum_range: CellOrRange, *criteria_pairs: CriteriaPair) -> str:
    parts = [sum_range]
    for range_ref, criteria in criteria_pairs:
        parts.extend([range_ref, excel_literal(criteria)])
    return f"=SUMIFS({','.join(parts)})"


def xlookup(
    lookup_value: CellOrRange,
    lookup_array: CellOrRange,
    return_array: CellOrRange,
    if_not_found: str = '""',
) -> str:
    return f"=XLOOKUP({lookup_value},{lookup_array},{return_array},{if_not_found})"


def index_match_lookup(
    lookup_value: CellOrRange,
    lookup_array: CellOrRange,
    return_array: CellOrRange,
) -> str:
    return f"=INDEX({return_array},MATCH({lookup_value},{lookup_array},0))"


def iferror(formula: str, fallback: Any = 0) -> str:
    return f"=IFERROR({strip_equals(formula)},{_fallback_literal(fallback)})"


def gross_margin(revenue_cell: CellOrRange, cost_cell: CellOrRange) -> str:
    revenue = _operand(revenue_cell)
    cost = _operand(cost_cell)
    return f"=IFERROR(({revenue}-{cost})/{revenue},0)"


def gross_margin_from_profit(profit_cell: CellOrRange, revenue_cell: CellOrRange) -> str:
    return f"=IFERROR({profit_cell}/{revenue_cell},0)"


def growth_rate(current_cell: CellOrRange, previous_cell: CellOrRange) -> str:
    return f"=IFERROR(({current_cell}-{previous_cell})/{previous_cell},0)"


def share(part_cell: CellOrRange, total_cell: CellOrRange) -> str:
    return f"=IFERROR({part_cell}/{total_cell},0)"


def yoy(current_cell: CellOrRange, prior_year_cell: CellOrRange) -> str:
    return growth_rate(current_cell, prior_year_cell)


def mom(current_cell: CellOrRange, prior_month_cell: CellOrRange) -> str:
    return growth_rate(current_cell, prior_month_cell)


def inventory_turnover(cogs_cell: CellOrRange, avg_inventory_cell: CellOrRange) -> str:
    return f"=IFERROR({cogs_cell}/{avg_inventory_cell},0)"


def average_order_value(revenue_cell: CellOrRange, order_count_cell: CellOrRange) -> str:
    return f"=IFERROR({revenue_cell}/{order_count_cell},0)"


def gmv(quantity_cell: CellOrRange, unit_price_cell: CellOrRange) -> str:
    return f"={quantity_cell}*{unit_price_cell}"


def roi(gain_cell: CellOrRange, cost_cell: CellOrRange) -> str:
    gain = _operand(gain_cell)
    cost = _operand(cost_cell)
    return f"=IFERROR(({gain}-{cost})/{cost},0)"


def project_completion_rate(done_cell: CellOrRange, total_cell: CellOrRange) -> str:
    return f"=IFERROR({done_cell}/{total_cell},0)"


def attendance_rate(actual_cell: CellOrRange, expected_cell: CellOrRange) -> str:
    return f"=IFERROR({actual_cell}/{expected_cell},0)"


def guarded(required_cells: Sequence[CellOrRange], formula: str, blank_value: str = '""') -> str:
    """Wrap a formula so unused template rows stay blank."""

    checks = ",".join(required_cells)
    return f'=IF(COUNTA({checks})=0,{blank_value},{strip_equals(formula)})'


def range_ref(sheet: str, column: str, start_row: int, end_row: int, absolute: bool = True) -> str:
    col = f"${column}" if absolute else column
    if absolute:
        return f"'{sheet}'!{col}${start_row}:{col}${end_row}"
    return f"'{sheet}'!{column}{start_row}:{column}{end_row}"


__all__ = [
    "average",
    "average_order_value",
    "attendance_rate",
    "countif",
    "countifs",
    "excel_sum",
    "gmv",
    "gross_margin",
    "gross_margin_from_profit",
    "growth_rate",
    "guarded",
    "iferror",
    "index_match_lookup",
    "inventory_turnover",
    "mom",
    "project_completion_rate",
    "range_ref",
    "roi",
    "share",
    "sumif",
    "sumifs",
    "xlookup",
    "yoy",
]
