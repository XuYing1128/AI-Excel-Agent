from types import SimpleNamespace

from excel_agent.api_settings import ApiSettings
from excel_agent.services import custom_api_service as service


class FakeResponse:
    ok = True
    status_code = 200
    text = ""
    elapsed = SimpleNamespace(total_seconds=lambda: 0.12)

    def json(self):
        return {"choices": [{"message": {"content": "连接成功"}}]}


def test_custom_api_connection_uses_compatible_endpoint(monkeypatch):
    captured = {}

    def fake_post(url, headers, json, timeout):
        captured.update(
            {"url": url, "headers": headers, "json": json, "timeout": timeout}
        )
        return FakeResponse()

    monkeypatch.setattr(service.requests, "post", fake_post)
    settings = ApiSettings(
        enabled=True,
        base_url="https://example.com/v1",
        api_key="secret",
        model="example-model",
    )
    result = service.test_api_connection(settings)

    assert result.success is True
    assert captured["url"] == "https://example.com/v1/chat/completions"
    assert captured["json"]["model"] == "example-model"
    assert captured["headers"]["Authorization"] == "Bearer secret"


def test_parse_json_object_accepts_fenced_text():
    result = service.parse_json_object('说明文字\\n{"status":"pass"}\\n结束')
    assert result["status"] == "pass"
