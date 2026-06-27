"""锁住两处误判修复：校验器不再把横向求和当成纵向少覆盖；reviewer 摘要带上冻结/筛选。"""

from openpyxl import Workbook

from excel_agent.services.subjective_review_service import _compact_workbook_summary
from excel_agent.validators import validate_workbook


def test_horizontal_sum_not_flagged_as_short_coverage(tmp_path):
    """各行的 =SUM(B2:E2)（横向算全年合计）不该被报 formula_range_short。"""
    path = tmp_path / "h.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "销售明细"
    ws.append(["产品线", "一季度", "二季度", "三季度", "四季度", "全年合计"])
    for idx, name in enumerate(["智能手机", "笔记本", "平板", "配件"], start=2):
        ws.append([name, 1, 2, 3, 4, f"=SUM(B{idx}:E{idx})"])
    ws.auto_filter.ref = "A1:F5"
    ws.freeze_panes = "A2"
    wb.save(path)

    report = validate_workbook(path)
    checks = [w.get("check") for w in report["warnings"]]
    assert "formula_range_short" not in checks


def test_vertical_short_range_still_flagged(tmp_path):
    """纵向范围确实少覆盖（同列、未到末行）仍应报，确保修复没把规则关死。"""
    path = tmp_path / "v.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "数据明细"
    ws.append(["名称", "数值", "累计"])
    for idx in range(2, 7):  # 数据到第 6 行
        ws.append([f"x{idx}", idx, f"=SUM($B$2:$B$3)"])  # 纵向只覆盖到第 3 行
    ws.auto_filter.ref = "A1:C6"
    wb.save(path)

    report = validate_workbook(path)
    checks = [w.get("check") for w in report["warnings"]]
    assert "formula_range_short" in checks


def test_compact_summary_includes_freeze_and_filter():
    summary = {
        "sheet_count": 1,
        "sheets": [
            {
                "name": "销售明细",
                "max_row": 5,
                "max_column": 6,
                "hidden": False,
                "title": "销售明细",
                "headers": ["产品线"],
                "formula_columns": ["全年合计"],
                "chart_count": 3,
                "freeze_panes": "A2",
                "auto_filter": "A1:F5",
            }
        ],
    }
    compact = _compact_workbook_summary(summary)
    sheet = compact["sheets"][0]
    assert sheet["freeze_panes"] == "A2"
    assert sheet["auto_filter"] == "A1:F5"
