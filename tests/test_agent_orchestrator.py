import json

from openpyxl import load_workbook

from excel_agent.model_registry import ModelSettings, ProviderConfig, save_model_settings
from excel_agent.services.agent import orchestrator
from excel_agent.services.custom_api_service import ToolCall, ToolChatResult
from excel_agent.task_paths import create_task_paths
from excel_agent.task_spec import TaskSpec


def _write_model_settings(path):
    save_model_settings(
        ModelSettings(
            providers=[
                ProviderConfig(
                    id="builder",
                    name="构建模型",
                    base_url="https://example.com/v1",
                    api_key="secret",
                    model="builder-model",
                )
            ],
            roles={"builder": "builder"},
            agent_enabled=True,
            run_python_enabled=True,
        ),
        path,
    )


def test_agent_orchestrator_builds_workbook_with_mock_tools(tmp_path, monkeypatch):
    model_path = tmp_path / "model_settings.json"
    _write_model_settings(model_path)
    monkeypatch.setenv("AI_EXCEL_MODEL_SETTINGS_FILE", str(model_path))
    blueprint = {
        "title": "项目清单",
        "sheet_name": "任务",
        "columns": [
            {"key": "task", "label": "任务", "type": "text"},
            {"key": "owner", "label": "负责人", "type": "text"},
            {"key": "progress", "label": "完成率", "type": "percentage"},
        ],
        "records": [
            {"task": "整理需求", "owner": "张三", "progress": 0.5},
            {"task": "生成表格", "owner": "李四", "progress": 0.8},
        ],
    }

    def fake_chat_with_tools(*args, **kwargs):
        return ToolChatResult(
            success=True,
            content="",
            tool_calls=[
                ToolCall("call_build", "build_workbook", {"blueprint": blueprint}),
                ToolCall("call_finish", "finish_task", {"summary": "完成"}),
            ],
            error=None,
            status_code=200,
            latency_ms=10,
            message={"role": "assistant", "content": "", "tool_calls": []},
        )

    monkeypatch.setattr(orchestrator.model_registry, "chat_with_tools", fake_chat_with_tools)
    paths = create_task_paths("project_plan", tmp_path / "tasks", output_name="项目清单.xlsx")
    result = orchestrator.run_agent(
        TaskSpec(task_type="project_plan", user_goal="生成项目清单", output_name="项目清单.xlsx"),
        paths,
        max_steps=3,
    )

    assert result.success is True
    assert paths.output_file.exists()
    assert result.tool_calls == 2
    wb = load_workbook(paths.output_file)
    assert "任务" in wb.sheetnames
    assert wb["任务"]["A1"].value == "项目清单"
    blueprint_payload = json.loads((paths.task_dir / "agent_workbook_blueprint.json").read_text(encoding="utf-8"))
    assert blueprint_payload["title"] == "项目清单"


def test_agent_orchestrator_returns_failure_when_model_stalls(tmp_path, monkeypatch):
    model_path = tmp_path / "model_settings.json"
    _write_model_settings(model_path)
    monkeypatch.setenv("AI_EXCEL_MODEL_SETTINGS_FILE", str(model_path))

    def fake_chat_with_tools(*args, **kwargs):
        return ToolChatResult(
            success=True,
            content="我先解释一下",
            tool_calls=[],
            error=None,
            status_code=200,
            latency_ms=10,
            message={"role": "assistant", "content": "我先解释一下"},
        )

    monkeypatch.setattr(orchestrator.model_registry, "chat_with_tools", fake_chat_with_tools)
    paths = create_task_paths("generic_table", tmp_path / "tasks")
    result = orchestrator.run_agent(
        TaskSpec(task_type="generic_table", user_goal="生成一个表"),
        paths,
        max_steps=2,
    )

    assert result.success is False
    assert "工具" in (result.error or "")

