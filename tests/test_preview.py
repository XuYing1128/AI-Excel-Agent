from excel_agent.preview import (
    sheet_preview_dataframe,
    workbook_chart_preview,
    workbook_preview,
)
from excel_agent.workbook_builder import create_workbook


def test_preview_uses_real_headers_and_hides_empty_formula_template_rows(tmp_path):
    workbook = create_workbook("generic_table", tmp_path / "preview.xlsx")
    preview = workbook_preview(workbook)
    data_sheet = next(item for item in preview["sheets"] if item["name"] == "Data")
    frame = sheet_preview_dataframe(data_sheet)

    assert list(frame.columns) == ["日期", "类别", "事项", "数量", "单价", "金额", "备注"]
    assert len(frame) == 2
    assert frame.iloc[0]["金额"] == "自动计算"


def test_chart_preview_is_available_for_chart_workbook(tmp_path):
    workbook = create_workbook("sales_report", tmp_path / "sales.xlsx")
    preview = workbook_preview(workbook)
    chart = workbook_chart_preview(preview)
    assert chart is not None
    frame, kind = chart
    assert not frame.empty
    assert kind == "line"
