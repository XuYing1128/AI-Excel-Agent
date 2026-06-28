from excel_agent.api_settings import ApiSettings
from excel_agent.services.custom_api_service import ApiCallResult
from excel_agent.services.review.structured_review import run_structured_reviews


def test_structured_reviews_return_two_reviewers():
    calls = []

    def fake_chat(settings, **kwargs):
        calls.append(kwargs["system_prompt"])
        return ApiCallResult(
            success=True,
            content='{"status":"pass","issues":[],"suggestions":[],"requires_user_confirmation":false}',
            error=None,
            status_code=200,
            latency_ms=1,
        )

    reviews = run_structured_reviews(
        {"任务方案": {"include_charts": True}},
        ApiSettings(
            enabled=True,
            base_url="https://example.com/v1",
            api_key="secret",
            model="model",
            provider_name="审查模型",
        ),
        chat_func=fake_chat,
    )

    assert [item["reviewer"] for item in reviews] == [
        "requirement_review",
        "excel_usability_review",
    ]
    assert all(item["status"] == "pass" for item in reviews)
    assert len(calls) == 2

