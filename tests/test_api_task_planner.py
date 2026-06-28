from excel_agent.api_settings import ApiSettings
from excel_agent.services import api_task_planner as planner
from excel_agent.services.custom_api_service import ApiCallResult
from excel_agent.task_spec_builder import build_task_spec_draft


def test_api_planner_enhances_draft_without_file_contents(monkeypatch):
    captured = {}

    def fake_chat(settings, **kwargs):
        captured.update(kwargs)
        return ApiCallResult(
            success=True,
            content=(
                '{"task_type":"ecommerce_analysis","confidence":0.92,'
                '"goal_summary":"分析订单表现","clarifying_questions":[],'
                '"include_charts":true,"include_summary":true}'
            ),
            error=None,
            status_code=200,
            latency_ms=20,
        )

    monkeypatch.setattr(planner, "chat_completion", fake_chat)
    draft = build_task_spec_draft("分析这个订单文件", ["orders.csv"])
    settings = ApiSettings(
        enabled=True,
        base_url="https://example.com/v1",
        api_key="secret",
        model="model",
    )
    result = planner.enhance_task_spec_draft(
        draft,
        user_prompt="分析这个订单文件",
        input_file_names=["orders.csv"],
        settings=settings,
    )

    assert result.used_api is True
    assert result.draft.task_spec.task_type == "ecommerce_analysis"
    assert "orders.csv" in captured["user_prompt"]
    assert "文件内容" not in captured["user_prompt"]


def test_api_planner_cannot_remove_explicit_chart_request(monkeypatch):
    def fake_chat(settings, **kwargs):
        return ApiCallResult(
            success=True,
            content=(
                '{"task_type":"sales_report","confidence":0.9,'
                '"goal_summary":"销售表","clarifying_questions":[],'
                '"include_charts":false,"include_summary":false}'
            ),
            error=None,
            status_code=200,
            latency_ms=20,
        )

    monkeypatch.setattr(planner, "chat_completion", fake_chat)
    draft = build_task_spec_draft("生成销售表，并添加柱状图。", [])
    settings = ApiSettings(
        enabled=True,
        base_url="https://example.com/v1",
        api_key="secret",
        model="model",
    )

    result = planner.enhance_task_spec_draft(
        draft,
        user_prompt="生成销售表，并添加柱状图。",
        input_file_names=[],
        settings=settings,
    )

    assert result.draft.task_spec.include_charts is True
    assert result.draft.task_spec.options["chart_requirements"]["required"] is True
