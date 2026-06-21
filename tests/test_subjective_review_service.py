import json

from excel_agent.api_settings import ApiSettings
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
