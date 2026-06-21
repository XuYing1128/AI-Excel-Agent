from pathlib import Path

from streamlit.testing.v1 import AppTest

from excel_agent.api_settings import load_api_settings


def test_streamlit_app_initial_page_loads():
    app_path = Path(__file__).resolve().parents[1] / "app.py"
    app = AppTest.from_file(str(app_path), default_timeout=15).run()
    assert not app.exception
    assert any("本地表格助手" in item.value for item in app.markdown)
    assert any(button.label == "检查并完善需求" for button in app.button)
    assert any(button.label == "接口设置" for button in app.button)
    assert any(uploader.label == "上传数据文件" for uploader in app.file_uploader)
    assert any(uploader.label == "上传模板文件" for uploader in app.file_uploader)
    assert any(item.label == "模板约束方式" for item in app.radio)
    assert not any("无需账号" in item.value for item in app.markdown)


def test_streamlit_confirmed_task_generates_and_validates(tmp_path, monkeypatch):
    monkeypatch.setenv("AI_EXCEL_OUTPUTS_DIR", str(tmp_path / "outputs"))
    monkeypatch.setenv(
        "AI_EXCEL_API_SETTINGS_FILE",
        str(tmp_path / "private" / "api_settings.json"),
    )
    app_path = Path(__file__).resolve().parents[1] / "app.py"
    app = AppTest.from_file(str(app_path), default_timeout=30).run()

    app.text_area[0].input("帮我生成个人月度收支预算表")
    next(button for button in app.button if button.label == "检查并完善需求").click()
    app.run()
    assert any(button.label == "确认并生成" for button in app.button)

    next(button for button in app.button if button.label == "确认并生成").click()
    app.run()
    assert not app.exception
    task_outputs = list((tmp_path / "outputs" / "tasks").glob("*/output/*.xlsx"))
    validation_reports = list(
        (tmp_path / "outputs" / "tasks").glob("*/reports/validation.json")
    )
    assert len(task_outputs) == 1
    assert len(validation_reports) == 1
    assert (tmp_path / "outputs" / "manifest.json").exists()


def test_streamlit_can_save_local_custom_api_settings(tmp_path, monkeypatch):
    settings_path = tmp_path / "private" / "api_settings.json"
    monkeypatch.setenv("AI_EXCEL_API_SETTINGS_FILE", str(settings_path))
    app_path = Path(__file__).resolve().parents[1] / "app.py"
    app = AppTest.from_file(str(app_path), default_timeout=20).run()

    next(button for button in app.button if button.label == "接口设置").click()
    app.run()
    next(button for button in app.button if button.label == "编辑接口设置").click()
    app.run()
    next(item for item in app.text_input if item.label == "接口名称").input("测试模型")
    next(item for item in app.text_input if item.label == "接口地址").input(
        "https://example.com/v1"
    )
    next(item for item in app.text_input if item.label == "模型名称").input("model")
    next(item for item in app.text_input if item.label == "接口密钥").input("secret")
    next(item for item in app.checkbox if item.label == "启用这个接口").check()
    next(button for button in app.button if button.label == "保存设置").click()
    app.run()

    assert not app.exception
    saved = load_api_settings(settings_path)
    assert saved.configured is True
    assert saved.provider_name == "测试模型"
