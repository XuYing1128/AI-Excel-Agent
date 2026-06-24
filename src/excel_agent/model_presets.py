"""各大模型厂商的 OpenAI 兼容接口预设。

只收录稳定的"基础地址(base_url)"和模型名示例。基础地址多年稳定；模型名会随版本
变化，请以各家控制台为准（示例仅作填写参考）。所有厂商都走 OpenAI 兼容的
``/chat/completions`` 协议，本项目据此调用。
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field


@dataclass(frozen=True)
class ProviderPreset:
    key: str
    name: str
    base_url: str            # 干净的基础地址，不带 /chat/completions
    model_examples: list[str] = field(default_factory=list)
    note: str = ""


# base_url 一律给"干净基础地址"，由 completion_endpoint 统一补 /chat/completions。
PROVIDER_PRESETS: list[ProviderPreset] = [
    ProviderPreset(
        key="deepseek",
        name="DeepSeek",
        base_url="https://api.deepseek.com/v1",
        model_examples=["deepseek-chat", "deepseek-reasoner"],
        note="官方 OpenAI 兼容。deepseek-chat 较快；deepseek-reasoner 会思考，更慢更费 token。",
    ),
    ProviderPreset(
        key="doubao",
        name="豆包(火山方舟 Ark)",
        base_url="https://ark.cn-beijing.volces.com/api/v3",
        model_examples=["在方舟控制台开通后填模型ID或接入点(ep-...)"],
        note="用 /api/v3（OpenAI 兼容）。不要用 /responses（那是另一种 Responses API，格式不同）。",
    ),
    ProviderPreset(
        key="kimi",
        name="Kimi(月之暗面 Moonshot)",
        base_url="https://api.moonshot.cn/v1",
        model_examples=["kimi-k2-0905-preview", "moonshot-v1-128k", "moonshot-v1-32k"],
        note="官方 OpenAI 兼容。",
    ),
    ProviderPreset(
        key="glm",
        name="智谱 GLM",
        base_url="https://open.bigmodel.cn/api/paas/v4",
        model_examples=["glm-4-plus", "glm-4-flash", "glm-4.5"],
        note="官方 OpenAI 兼容（/api/paas/v4）。",
    ),
    ProviderPreset(
        key="qwen",
        name="通义千问(阿里 DashScope)",
        base_url="https://dashscope.aliyuncs.com/compatible-mode/v1",
        model_examples=["qwen-plus", "qwen-max", "qwen-turbo"],
        note="用 compatible-mode（OpenAI 兼容）。qwen3 思考型模型在非流式下可能需要关闭思考，建议用非思考型或 -instruct 版本。",
    ),
    ProviderPreset(
        key="mimo",
        name="小米 MiMo",
        base_url="https://api.xiaomimimo.com/v1",
        model_examples=["mimo-v2.5-pro", "mimo-v2.5"],
        note="OpenAI 兼容。",
    ),
    ProviderPreset(
        key="custom",
        name="其它(自定义/网关)",
        base_url="",
        model_examples=[],
        note="任何 OpenAI 兼容地址（如 OneAPI 网关 http://127.0.0.1:3000/v1）。",
    ),
]

PRESET_BY_KEY = {item.key: item for item in PROVIDER_PRESETS}

# 用主机名片段把已填的 base_url 反查到厂商，便于设置页高亮与提示。
_HOST_HINTS: list[tuple[str, str]] = [
    ("api.deepseek.com", "deepseek"),
    ("volces.com", "doubao"),
    ("ark.cn-", "doubao"),
    ("moonshot.cn", "kimi"),
    ("bigmodel.cn", "glm"),
    ("dashscope", "qwen"),
    ("xiaomimimo.com", "mimo"),
]


def normalize_base_url(base_url: str) -> str:
    """把用户填的地址清洗成"干净基础地址"。

    去掉常见的 endpoint 路径后缀（/chat/completions、/responses、/completions）和
    多余斜杠，这样无论用户从控制台复制了哪段，都能由 completion_endpoint 正确拼接。
    """

    url = str(base_url or "").strip().rstrip("/")
    if not url:
        return url
    lowered = url.lower()
    for suffix in ("/chat/completions", "/responses", "/completions", "/chat"):
        if lowered.endswith(suffix):
            url = url[: -len(suffix)].rstrip("/")
            lowered = url.lower()
    return url


def detect_provider_key(base_url: str) -> str:
    host = str(base_url or "").lower()
    for fragment, key in _HOST_HINTS:
        if fragment in host:
            return key
    return "custom"
