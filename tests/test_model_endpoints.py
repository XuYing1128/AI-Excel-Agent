"""每个厂商的 base_url 都要拼出正确的 OpenAI 兼容 /chat/completions。"""

import pytest

from excel_agent.model_presets import detect_provider_key, normalize_base_url
from excel_agent.model_registry import ProviderConfig
from excel_agent.services.custom_api_service import completion_endpoint


@pytest.mark.parametrize(
    "base,expected",
    [
        ("https://api.deepseek.com/v1", "https://api.deepseek.com/v1/chat/completions"),
        ("https://api.deepseek.com", "https://api.deepseek.com/v1/chat/completions"),
        ("https://ark.cn-beijing.volces.com/api/v3", "https://ark.cn-beijing.volces.com/api/v3/chat/completions"),
        # Responses-API path given by mistake must be corrected.
        ("https://ark.cn-beijing.volces.com/api/v3/responses", "https://ark.cn-beijing.volces.com/api/v3/chat/completions"),
        ("https://open.bigmodel.cn/api/paas/v4", "https://open.bigmodel.cn/api/paas/v4/chat/completions"),
        ("https://open.bigmodel.cn/api/paas/v4/chat/completions", "https://open.bigmodel.cn/api/paas/v4/chat/completions"),
        ("https://dashscope.aliyuncs.com/compatible-mode/v1", "https://dashscope.aliyuncs.com/compatible-mode/v1/chat/completions"),
        ("https://api.moonshot.cn/v1", "https://api.moonshot.cn/v1/chat/completions"),
        ("https://api.xiaomimimo.com/v1", "https://api.xiaomimimo.com/v1/chat/completions"),
        ("http://127.0.0.1:3000/v1", "http://127.0.0.1:3000/v1/chat/completions"),
    ],
)
def test_completion_endpoint_for_every_provider(base, expected):
    assert completion_endpoint(base) == expected


@pytest.mark.parametrize(
    "raw,clean",
    [
        ("https://ark.cn-beijing.volces.com/api/v3/responses", "https://ark.cn-beijing.volces.com/api/v3"),
        ("https://open.bigmodel.cn/api/paas/v4/chat/completions", "https://open.bigmodel.cn/api/paas/v4"),
        ("https://api.deepseek.com/v1/", "https://api.deepseek.com/v1"),
        ("https://api.moonshot.cn/v1", "https://api.moonshot.cn/v1"),
    ],
)
def test_normalize_base_url_strips_endpoint_suffixes(raw, clean):
    assert normalize_base_url(raw) == clean


def test_provider_config_normalizes_doubao_responses_base():
    # The user's saved 豆包 base pointed at /responses; loading must clean it.
    cfg = ProviderConfig(
        id="doubao",
        name="豆包",
        base_url="https://ark.cn-beijing.volces.com/api/v3/responses",
        api_key="k",
        model="doubao-seed-2-1-pro-260628",
    )
    assert cfg.base_url == "https://ark.cn-beijing.volces.com/api/v3"
    assert completion_endpoint(cfg.base_url).endswith("/api/v3/chat/completions")


@pytest.mark.parametrize(
    "base,key",
    [
        ("https://api.deepseek.com/v1", "deepseek"),
        ("https://ark.cn-beijing.volces.com/api/v3", "doubao"),
        ("https://api.moonshot.cn/v1", "kimi"),
        ("https://open.bigmodel.cn/api/paas/v4", "glm"),
        ("https://dashscope.aliyuncs.com/compatible-mode/v1", "qwen"),
        ("https://api.xiaomimimo.com/v1", "mimo"),
        ("http://127.0.0.1:3000/v1", "custom"),
    ],
)
def test_detect_provider_key(base, key):
    assert detect_provider_key(base) == key
