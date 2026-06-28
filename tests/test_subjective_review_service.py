import json

from excel_agent.api_settings import ApiSettings
from excel_agent.model_registry import ModelSettings, ProviderConfig, save_model_settings
from excel_agent.services import subjective_review_service as review_service
from excel_agent.services.custom_api_service import ApiCallResult
from excel_agent.task_paths import create_task_paths
from excel_agent.task_spec import TaskSpec


def test_custom_subjective_review_is_non_blocking_and_uses_safe_summary(
    tmp_path, monkeypatch
):
    captured = {}

    def fake_chat(settings, **kwargs):
        captured.update(kwargs)
        return ApiCallResult(
            success=True,
            content=(
                '{"status":"pass","fit_to_user_goal":"pass",'
                '"over_design_risk":"low","concerns":[],"suggestions":[]}'
            ),
            error=None,
            status_code=200,
            latency_ms=12,
        )

    monkeypatch.setattr(review_service, "chat_completion", fake_chat)
    paths = create_task_paths("personal_budget", tmp_path / "tasks")
    spec = TaskSpec(
        task_type="personal_budget",
        user_goal="制作预算表",
        input_files=[str(tmp_path / "private" / "budget.xlsx")],
    )
    settings = ApiSettings(
        enabled=True,
        base_url="https://example.com/v1",
        api_key="secret",
        model="model",
        provider_name="测试模型",
    )
    result = review_service.run_subjective_review(
        spec,
        {"status": "pass", "error_count": 0, "warning_count": 0},
        {"sheet_count": 3, "sheets": [{"name": "Data", "max_row": 20}]},
        {"mode": "standard_template", "message": "完成"},
        paths,
        settings,
    )

    assert result["enabled"] is True
    assert result["reviews"][0]["model"] == "测试模型"
    assert str(tmp_path) not in captured["user_prompt"]
    saved = json.loads(paths.subjective_review_report.read_text(encoding="utf-8"))
    assert saved["reviews"][0]["status"] == "pass"


def test_review_cannot_recommend_removing_user_confirmed_chart(tmp_path, monkeypatch):
    def fake_chat(settings, **kwargs):
        return ApiCallResult(
            success=True,
            content=(
                '{"status":"warn","fit_to_user_goal":"warn",'
                '"over_design_risk":"medium",'
                '"concerns":["用户未要求图表"],'
                '"suggestions":["删除图表，保持表格纯净"]}'
            ),
            error=None,
            status_code=200,
            latency_ms=12,
        )

    monkeypatch.setattr(review_service, "chat_completion", fake_chat)
    paths = create_task_paths("sales_report", tmp_path / "tasks")
    spec = TaskSpec(
        task_type="sales_report",
        user_goal="制作销售表",
        include_charts=True,
    )
    settings = ApiSettings(
        enabled=True,
        base_url="https://example.com/v1",
        api_key="secret",
        model="model",
    )
    result = review_service.run_subjective_review(
        spec,
        {"status": "pass"},
        {"sheet_count": 1, "sheets": [{"name": "销售业绩", "chart_count": 1}]},
        {"mode": "llm_tool_agent", "message": "完成"},
        paths,
        settings,
    )
    assert result["reviews"][0]["status"] == "pass"
    assert result["reviews"][0]["concerns"] == []
    assert result["reviews"][0]["suggestions"] == []


def test_subjective_review_prefers_reviewer_role_over_legacy_api_settings(
    tmp_path,
    monkeypatch,
):
    model_path = tmp_path / "private" / "model_settings.json"
    monkeypatch.setenv("AI_EXCEL_MODEL_SETTINGS_FILE", str(model_path))
    save_model_settings(
        ModelSettings(
            providers=[
                ProviderConfig(
                    id="reviewer",
                    name="审查模型",
                    base_url="https://reviewer.example/v1",
                    api_key="secret",
                    model="reviewer-model",
                )
            ],
            roles={"reviewer": "reviewer"},
        ),
        model_path,
    )
    used = {}

    def fake_chat(settings, **kwargs):
        used["provider"] = settings.provider_name
        return ApiCallResult(
            success=True,
            content='{"status":"pass","issues":[],"suggestions":[]}',
            error=None,
            status_code=200,
            latency_ms=12,
        )

    monkeypatch.setattr(review_service, "chat_completion", fake_chat)
    paths = create_task_paths("sales_report", tmp_path / "tasks")
    legacy = ApiSettings(
        enabled=True,
        base_url="https://legacy.example/v1",
        api_key="legacy",
        model="legacy-model",
        provider_name="旧单模型",
    )

    result = review_service.run_subjective_review(
        TaskSpec(task_type="sales_report", user_goal="做销售表"),
        {"status": "pass"},
        {"sheet_count": 1, "sheets": [{"name": "明细"}]},
        {"mode": "domain_compiler", "message": "完成"},
        paths,
        legacy,
    )

    assert result["enabled"] is True
    assert used["provider"] == "审查模型"
