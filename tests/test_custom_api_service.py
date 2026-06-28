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


class FakeToolResponse(FakeResponse):
    def json(self):
        return {
            "choices": [
                {
                    "message": {
                        "content": None,
                        "tool_calls": [
                            {
                                "id": "call_1",
                                "type": "function",
                                "function": {
                                    "name": "build_workbook",
                                    "arguments": '{"blueprint":{"title":"测试"}}',
                                },
                            }
                        ],
                    }
                }
            ]
        }


class FakeTruncatedToolResponse(FakeResponse):
    def json(self):
        return {
            "choices": [
                {
                    "finish_reason": "length",
                    "message": {
                        "content": None,
                        "tool_calls": [
                            {
                                "id": "call_1",
                                "type": "function",
                                "function": {
                                    "name": "build_workbook",
                                    "arguments": '{"blueprint":',
                                },
                            }
                        ],
                    },
                }
            ]
        }


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


def test_custom_api_supports_openai_compatible_tool_calls(monkeypatch):
    captured = {}

    def fake_post(url, headers, json, timeout):
        captured["json"] = json
        return FakeToolResponse()

    monkeypatch.setattr(service.requests, "post", fake_post)
    settings = ApiSettings(
        enabled=True,
        base_url="https://example.com/v1",
        api_key="secret",
        model="example-model",
    )
    result = service.chat_completion_with_tools(
        settings,
        messages=[{"role": "user", "content": "test"}],
        tools=[
            {
                "type": "function",
                "function": {
                    "name": "build_workbook",
                    "parameters": {"type": "object"},
                },
            }
        ],
    )
    assert result.success is True
    assert result.tool_calls[0].name == "build_workbook"
    assert result.tool_calls[0].arguments["blueprint"]["title"] == "测试"
    assert captured["json"]["tools"]


def test_custom_api_retries_truncated_tool_arguments(monkeypatch):
    responses = iter([FakeTruncatedToolResponse(), FakeToolResponse()])
    calls = []

    def fake_post(url, headers, json, timeout):
        calls.append(url)
        return next(responses)

    monkeypatch.setattr(service.requests, "post", fake_post)
    monkeypatch.setattr(service.time, "sleep", lambda _seconds: None)
    settings = ApiSettings(
        enabled=True,
        base_url="https://example.com/v1",
        api_key="secret",
        model="example-model",
    )
    result = service.chat_completion_with_tools(
        settings,
        messages=[{"role": "user", "content": "test"}],
        tools=[],
    )

    assert result.success is True
    assert len(calls) == 2


def test_custom_api_retries_direct_when_proxy_fails(monkeypatch):
    calls = []

    def fake_post(url, headers, json, timeout):
        calls.append(("proxy", url))
        raise service.requests.exceptions.ProxyError("bad proxy")

    class FakeSession:
        trust_env = True

        def post(self, url, headers, json, timeout):
            calls.append(("direct", url, self.trust_env))
            return FakeResponse()

    monkeypatch.setattr(service.requests, "post", fake_post)
    monkeypatch.setattr(service.requests, "Session", FakeSession)
    settings = ApiSettings(
        enabled=True,
        base_url="https://example.com/v1",
        api_key="secret",
        model="example-model",
    )

    result = service.chat_completion(
        settings,
        system_prompt="s",
        user_prompt="u",
    )

    assert result.success is True
    assert calls[0][0] == "proxy"
    assert calls[1] == ("direct", "https://example.com/v1/chat/completions", False)
