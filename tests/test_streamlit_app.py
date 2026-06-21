from pathlib import Path

from streamlit.testing.v1 import AppTest


def test_streamlit_app_initial_page_loads():
    app_path = Path(__file__).resolve().parents[1] / "app.py"
    app = AppTest.from_file(str(app_path), default_timeout=15).run()
    assert not app.exception
    assert any("AI-Excel-Agent 本地表格生成工具" in item.value for item in app.title)
    assert any(button.label == "分析需求" for button in app.button)


def test_streamlit_confirmed_task_generates_and_validates(tmp_path, monkeypatch):
    monkeypatch.setenv("AI_EXCEL_OUTPUTS_DIR", str(tmp_path / "outputs"))
    app_path = Path(__file__).resolve().parents[1] / "app.py"
    app = AppTest.from_file(str(app_path), default_timeout=30).run()

    app.text_area[0].input("帮我生成个人月度收支预算表")
    next(button for button in app.button if button.label == "分析需求").click()
    app.run()
    assert any(button.label == "确认并生成" for button in app.button)

    next(button for button in app.button if button.label == "确认并生成").click()
    app.run()
    assert not app.exception
    task_outputs = list((tmp_path / "outputs" / "tasks").glob("*/output/result.xlsx"))
    validation_reports = list(
        (tmp_path / "outputs" / "tasks").glob("*/reports/validation.json")
    )
    assert len(task_outputs) == 1
    assert len(validation_reports) == 1
    assert (tmp_path / "outputs" / "manifest.json").exists()
