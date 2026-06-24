"""Small OpenAI-compatible chat client used only for text understanding."""

from __future__ import annotations

import json
import re
import time
from dataclasses import dataclass
from typing import Any

import requests

from ..api_settings import ApiSettings

try:
    from json_repair import repair_json
except ImportError:  # pragma: no cover - dependency is declared for normal installs
    repair_json = None


JSON_BLOCK_RE = re.compile(r"\{.*\}", re.DOTALL)


TRUNCATION_HINT = (
    "模型输出被截断：思考占满了输出长度，没能产出完整内容。"
    "请在接口设置中调大“等待时间”，或改用响应更快的非推理模型（例如 deepseek-chat）。"
)


@dataclass
class ApiCallResult:
    success: bool
    content: str
    error: str | None
    status_code: int | None
    latency_ms: int | None
    finish_reason: str | None = None


@dataclass
class ToolCall:
    id: str
    name: str
    arguments: dict[str, Any]


@dataclass
class ToolChatResult:
    success: bool
    content: str
    tool_calls: list[ToolCall]
    error: str | None
    status_code: int | None
    latency_ms: int | None
    message: dict[str, Any] | None


def chat_completion(
    settings: ApiSettings,
    *,
    system_prompt: str,
    user_prompt: str,
    temperature: float = 0.1,
    max_tokens: int = 800,
    json_mode: bool = False,
) -> ApiCallResult:
    if not settings.configured:
        return ApiCallResult(False, "", "接口配置不完整。", None, None)

    endpoint = completion_endpoint(settings.base_url)
    payload = {
        "model": settings.model,
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt},
        ],
        "temperature": temperature,
        "max_tokens": max_tokens,
        "stream": False,
    }
    if json_mode:
        payload["response_format"] = {"type": "json_object"}
    headers = {
        "Authorization": f"Bearer {settings.api_key}",
        "Content-Type": "application/json",
        "Accept": "application/json",
    }
    try:
        response = requests.post(
            endpoint,
            headers=headers,
            json=payload,
            timeout=settings.timeout_seconds,
        )
    except requests.RequestException as exc:
        return ApiCallResult(False, "", f"连接失败：{exc}", None, None)

    elapsed_ms = int(response.elapsed.total_seconds() * 1000)
    if not response.ok and json_mode and response.status_code in {400, 404, 422}:
        payload.pop("response_format", None)
        try:
            response = requests.post(
                endpoint,
                headers=headers,
                json=payload,
                timeout=settings.timeout_seconds,
            )
            elapsed_ms = int(response.elapsed.total_seconds() * 1000)
        except requests.RequestException as exc:
            return ApiCallResult(False, "", f"连接失败：{exc}", None, None)
    if not response.ok:
        detail = _response_error_message(response)
        return ApiCallResult(
            False,
            "",
            f"接口返回 {response.status_code}：{detail}",
            response.status_code,
            elapsed_ms,
        )

    try:
        data = response.json()
        choice = data["choices"][0]
        content = choice["message"].get("content")
        finish_reason = choice.get("finish_reason")
    except (ValueError, KeyError, IndexError, TypeError) as exc:
        return ApiCallResult(
            False,
            "",
            f"接口响应格式无法识别：{exc}",
            response.status_code,
            elapsed_ms,
        )
    text = str(content or "").strip()
    if not text and finish_reason == "length":
        return ApiCallResult(
            False, "", TRUNCATION_HINT, response.status_code, elapsed_ms, finish_reason
        )
    return ApiCallResult(
        True,
        text,
        None,
        response.status_code,
        elapsed_ms,
        finish_reason,
    )


def chat_completion_with_tools(
    settings: ApiSettings,
    *,
    messages: list[dict[str, Any]],
    tools: list[dict[str, Any]],
    tool_choice: str | dict[str, Any] = "auto",
    temperature: float = 0.1,
    max_tokens: int = 3000,
) -> ToolChatResult:
    """Call an OpenAI-compatible endpoint with function tools."""

    if not settings.configured:
        return ToolChatResult(False, "", [], "接口配置不完整。", None, None, None)
    endpoint = completion_endpoint(settings.base_url)
    payload = {
        "model": settings.model,
        "messages": messages,
        "tools": tools,
        "tool_choice": tool_choice,
        "temperature": temperature,
        "max_tokens": max_tokens,
        "stream": False,
    }
    headers = {
        "Authorization": f"Bearer {settings.api_key}",
        "Content-Type": "application/json",
        "Accept": "application/json",
    }
    response = None
    last_exception: requests.RequestException | None = None
    last_parse_error: Exception | None = None
    last_finish_reason = ""
    for attempt in range(3):
        try:
            response = requests.post(
                endpoint,
                headers=headers,
                json=payload,
                timeout=settings.timeout_seconds,
            )
        except requests.RequestException as exc:
            last_exception = exc
            if attempt < 2:
                time.sleep(1.5 * (attempt + 1))
            continue

        elapsed_ms = int(response.elapsed.total_seconds() * 1000)
        if not response.ok:
            if response.status_code >= 500 and attempt < 2:
                time.sleep(1.5 * (attempt + 1))
                continue
            return ToolChatResult(
                False,
                "",
                [],
                f"接口返回 {response.status_code}：{_response_error_message(response)}",
                response.status_code,
                elapsed_ms,
                None,
            )
        try:
            data = response.json()
            choice = data["choices"][0]
            last_finish_reason = str(choice.get("finish_reason") or "")
            message = choice["message"]
            content = str(message.get("content") or "").strip()
            parsed_calls = []
            for raw in message.get("tool_calls") or []:
                function = raw.get("function") or {}
                arguments = function.get("arguments") or "{}"
                if isinstance(arguments, str):
                    arguments = (
                        json.loads(arguments)
                        if last_finish_reason == "length"
                        else parse_json_object(arguments)
                    )
                if not isinstance(arguments, dict):
                    raise TypeError("tool arguments 不是 JSON 对象")
                parsed_calls.append(
                    ToolCall(
                        id=str(raw.get("id") or f"call_{len(parsed_calls) + 1}"),
                        name=str(function.get("name") or ""),
                        arguments=arguments,
                    )
                )
        except (ValueError, KeyError, IndexError, TypeError, json.JSONDecodeError) as exc:
            last_parse_error = exc
            if last_finish_reason == "length":
                # Truncation (the model's reasoning ate the output budget). Give
                # one quick retry in case it was transient, then fail fast so the
                # caller switches to JSON mode / local generation instead of
                # burning a third slow call (which just truncates again).
                if attempt == 0:
                    continue
                return ToolChatResult(
                    False, "", [], TRUNCATION_HINT, response.status_code, elapsed_ms, None
                )
            if attempt < 2:
                time.sleep(1.5 * (attempt + 1))
                continue
            return ToolChatResult(
                False,
                "",
                [],
                (
                    f"接口工具调用响应无法识别，已重试 3 次：{exc}"
                    + (
                        f"；finish_reason={last_finish_reason}"
                        if last_finish_reason
                        else ""
                    )
                ),
                response.status_code,
                elapsed_ms,
                None,
            )
        if not parsed_calls and not content.strip() and last_finish_reason == "length":
            return ToolChatResult(
                False,
                "",
                [],
                TRUNCATION_HINT,
                response.status_code,
                elapsed_ms,
                None,
            )
        return ToolChatResult(
            True,
            content,
            parsed_calls,
            None,
            response.status_code,
            elapsed_ms,
            message,
        )

    return ToolChatResult(
        False,
        "",
        [],
        (
            f"连接失败，已重试 3 次：{last_exception}"
            if last_exception
            else f"接口响应无法解析：{last_parse_error}"
        ),
        response.status_code if response is not None else None,
        (
            int(response.elapsed.total_seconds() * 1000)
            if response is not None
            else None
        ),
        None,
    )


def test_api_connection(settings: ApiSettings) -> ApiCallResult:
    # Give reasoning ("thinking") models enough room to both think and reply;
    # max_tokens=16 made every reasoning model look "broken" (the reasoning ate
    # the budget -> empty content -> truncation).
    result = chat_completion(
        settings,
        system_prompt="你是连接测试助手。只按要求回复，不要解释。",
        user_prompt="请只回复四个字：连接成功",
        temperature=0,
        max_tokens=2048,
    )
    if result.success:
        return result
    # A 200 that only truncated means the endpoint, key and model are all valid —
    # it is simply a thinking model. Report success so the user is not misled.
    if result.finish_reason == "length" or result.status_code == 200:
        return ApiCallResult(
            True,
            "连接正常（这是推理型模型，思考较多，正式使用时建议把等待时间调大或用于审查类角色）。",
            None,
            result.status_code,
            result.latency_ms,
            result.finish_reason,
        )
    return result


def parse_json_object(content: str) -> dict[str, Any]:
    text = str(content or "").strip()
    try:
        parsed = json.loads(text)
    except json.JSONDecodeError:
        match = JSON_BLOCK_RE.search(text)
        candidate = match.group(0) if match else text[text.find("{") :] if "{" in text else ""
        if not candidate:
            raise ValueError("模型没有返回 JSON 对象。") from None
        try:
            parsed = json.loads(candidate)
        except json.JSONDecodeError as exc:
            if repair_json is None:
                raise ValueError(f"模型返回的 JSON 无法解析：{exc}") from exc
            try:
                parsed = repair_json(candidate, return_objects=True)
            except Exception as repair_exc:
                raise ValueError(
                    f"模型返回的 JSON 无法解析，自动修复也失败：{repair_exc}"
                ) from repair_exc
    if not isinstance(parsed, dict):
        raise ValueError("模型返回的 JSON 顶层不是对象。")
    return parsed


def completion_endpoint(base_url: str) -> str:
    """Build the OpenAI-compatible chat-completions URL from any base shape.

    Handles every provider we support:
      - already full: .../chat/completions                -> as-is
      - Responses API given by mistake: .../responses     -> swap to chat
      - ends in a version segment (/v1, /v3, /api/paas/v4,
        /compatible-mode/v1 ...)                           -> append /chat/completions
      - bare host (https://api.deepseek.com)               -> append /v1/chat/completions
    """

    url = str(base_url or "").strip().rstrip("/")
    if not url:
        return url
    if url.endswith("/chat/completions"):
        return url
    if url.endswith("/responses"):
        return url[: -len("/responses")] + "/chat/completions"
    last_segment = url.rsplit("/", 1)[-1].lower()
    if re.fullmatch(r"v\d+(?:beta)?", last_segment):
        return f"{url}/chat/completions"
    return f"{url}/v1/chat/completions"


def _response_error_message(response: requests.Response) -> str:
    try:
        data = response.json()
        if isinstance(data, dict):
            error = data.get("error")
            if isinstance(error, dict) and error.get("message"):
                return str(error["message"])[:300]
            if error:
                return str(error)[:300]
            if data.get("message"):
                return str(data["message"])[:300]
    except ValueError:
        pass
    return response.text.strip()[:300] or "没有错误详情"
