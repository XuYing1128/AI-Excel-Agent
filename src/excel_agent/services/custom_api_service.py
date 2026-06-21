"""Small OpenAI-compatible chat client used only for text understanding."""

from __future__ import annotations

import json
import re
from dataclasses import dataclass
from typing import Any

import requests

from ..api_settings import ApiSettings


JSON_BLOCK_RE = re.compile(r"\{.*\}", re.DOTALL)


@dataclass
class ApiCallResult:
    success: bool
    content: str
    error: str | None
    status_code: int | None
    latency_ms: int | None


def chat_completion(
    settings: ApiSettings,
    *,
    system_prompt: str,
    user_prompt: str,
    temperature: float = 0.1,
    max_tokens: int = 800,
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
        content = data["choices"][0]["message"]["content"]
    except (ValueError, KeyError, IndexError, TypeError) as exc:
        return ApiCallResult(
            False,
            "",
            f"接口响应格式无法识别：{exc}",
            response.status_code,
            elapsed_ms,
        )
    return ApiCallResult(
        True,
        str(content).strip(),
        None,
        response.status_code,
        elapsed_ms,
    )


def test_api_connection(settings: ApiSettings) -> ApiCallResult:
    return chat_completion(
        settings,
        system_prompt="你是连接测试助手。只按要求回复，不要解释。",
        user_prompt="请只回复四个字：连接成功",
        temperature=0,
        max_tokens=16,
    )


def parse_json_object(content: str) -> dict[str, Any]:
    text = str(content or "").strip()
    try:
        parsed = json.loads(text)
    except json.JSONDecodeError:
        match = JSON_BLOCK_RE.search(text)
        if not match:
            raise ValueError("模型没有返回 JSON 对象。") from None
        try:
            parsed = json.loads(match.group(0))
        except json.JSONDecodeError as exc:
            raise ValueError(f"模型返回的 JSON 无法解析：{exc}") from exc
    if not isinstance(parsed, dict):
        raise ValueError("模型返回的 JSON 顶层不是对象。")
    return parsed


def completion_endpoint(base_url: str) -> str:
    url = str(base_url or "").strip().rstrip("/")
    if url.endswith("/chat/completions"):
        return url
    if url.endswith("/v1"):
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
