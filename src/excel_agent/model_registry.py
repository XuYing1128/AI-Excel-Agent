"""本地多模型配置与角色分工。

本模块只管理 OpenAI 兼容接口的本机配置，不上传、不同步、不写入仓库。
旧版 ``api_settings.json`` 会在首次读取时自动迁移为一个 provider，旧文件保留。
"""

from __future__ import annotations

import json
import os
import re
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Any, Literal

from .api_settings import ApiSettings, api_settings_path, load_api_settings
from .io_utils import project_root
from .model_presets import normalize_base_url
from .services.custom_api_service import (
    ApiCallResult,
    ToolChatResult,
    chat_completion,
    chat_completion_with_tools,
    test_api_connection,
)


RoleName = Literal["planner", "builder", "reviewer", "fast", "coder"]
ROLE_NAMES: tuple[RoleName, ...] = ("planner", "builder", "reviewer", "fast", "coder")
ROLE_LABELS: dict[str, str] = {
    "planner": "理解需求",
    "builder": "设计并生成表格",
    "reviewer": "审查结果",
    "fast": "分类、摘要、起名",
    "coder": "编写安全脚本",
}


@dataclass
class ProviderConfig:
    id: str = ""
    name: str = "自定义模型"
    base_url: str = ""
    api_key: str = ""
    model: str = ""
    timeout_seconds: int = 120
    enabled: bool = True

    def __post_init__(self) -> None:
        self.id = safe_provider_id(self.id or self.name or self.model or "provider")
        self.name = str(self.name or "自定义模型").strip() or "自定义模型"
        # Clean the base URL so any console-pasted endpoint suffix
        # (/chat/completions, /responses ...) is stripped to a tidy base.
        self.base_url = normalize_base_url(self.base_url)
        self.api_key = str(self.api_key or "").strip()
        self.model = str(self.model or "").strip()
        self.timeout_seconds = min(max(int(self.timeout_seconds or 120), 5), 600)
        self.enabled = bool(self.enabled)

    @property
    def configured(self) -> bool:
        return bool(self.enabled and self.base_url and self.api_key and self.model)

    def to_api_settings(
        self,
        *,
        use_for_intent: bool = True,
        use_for_review: bool = True,
        use_for_generation: bool = True,
    ) -> ApiSettings:
        return ApiSettings(
            enabled=self.enabled,
            base_url=self.base_url,
            api_key=self.api_key,
            model=self.model,
            provider_name=self.name,
            timeout_seconds=self.timeout_seconds,
            use_for_intent=use_for_intent,
            use_for_review=use_for_review,
            use_for_generation=use_for_generation,
        )

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> "ProviderConfig":
        allowed = {
            "id",
            "name",
            "base_url",
            "api_key",
            "model",
            "timeout_seconds",
            "enabled",
        }
        return cls(**{key: data.get(key) for key in allowed if key in data})

    def to_dict(self, include_secret: bool = True) -> dict[str, Any]:
        payload = asdict(self)
        if not include_secret:
            payload["api_key"] = "已保存" if self.api_key else ""
        return payload


@dataclass
class ModelSettings:
    providers: list[ProviderConfig] = field(default_factory=list)
    roles: dict[str, str] = field(default_factory=dict)
    agent_enabled: bool = True
    run_python_enabled: bool = True

    def __post_init__(self) -> None:
        normalized: list[ProviderConfig] = []
        used: set[str] = set()
        for item in self.providers:
            provider = item if isinstance(item, ProviderConfig) else ProviderConfig.from_dict(item)
            base_id = provider.id
            suffix = 2
            while provider.id in used:
                provider.id = f"{base_id}-{suffix}"
                suffix += 1
            used.add(provider.id)
            normalized.append(provider)
        self.providers = normalized
        self.roles = {str(k): str(v) for k, v in (self.roles or {}).items() if str(k) in ROLE_NAMES}
        self.agent_enabled = bool(self.agent_enabled)
        self.run_python_enabled = bool(self.run_python_enabled)

    def to_dict(self, include_secret: bool = True) -> dict[str, Any]:
        return {
            "providers": [item.to_dict(include_secret=include_secret) for item in self.providers],
            "roles": dict(self.roles),
            "agent_enabled": self.agent_enabled,
            "run_python_enabled": self.run_python_enabled,
        }

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> "ModelSettings":
        providers = data.get("providers", [])
        if not isinstance(providers, list):
            providers = []
        roles = data.get("roles", {})
        if not isinstance(roles, dict):
            roles = {}
        return cls(
            providers=[ProviderConfig.from_dict(item) for item in providers if isinstance(item, dict)],
            roles={str(k): str(v) for k, v in roles.items()},
            agent_enabled=bool(data.get("agent_enabled", True)),
            run_python_enabled=bool(data.get("run_python_enabled", True)),
        )


def model_settings_path(path: str | Path | None = None) -> Path:
    override = os.getenv("AI_EXCEL_MODEL_SETTINGS_FILE", "").strip()
    api_override = os.getenv("AI_EXCEL_API_SETTINGS_FILE", "").strip()
    if path is not None:
        target = Path(path)
    elif override:
        target = Path(override)
    elif api_override:
        target = Path(api_override).with_name("model_settings.json")
    else:
        target = project_root() / "data" / "private" / "model_settings.json"
    return target.expanduser().resolve()


def load_model_settings(
    path: str | Path | None = None,
    *,
    legacy_api_path: str | Path | None = None,
) -> ModelSettings:
    target = model_settings_path(path)
    if target.exists():
        try:
            payload = json.loads(target.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError, TypeError):
            return ModelSettings()
        return ModelSettings.from_dict(payload) if isinstance(payload, dict) else ModelSettings()
    legacy = load_api_settings(legacy_api_path)
    if legacy.configured:
        migrated = from_legacy_api_settings(legacy)
        save_model_settings(migrated, target)
        return migrated
    return ModelSettings()


def save_model_settings(settings: ModelSettings, path: str | Path | None = None) -> Path:
    target = model_settings_path(path)
    normalized = settings if isinstance(settings, ModelSettings) else ModelSettings.from_dict(settings)
    target.parent.mkdir(parents=True, exist_ok=True)
    temporary = target.with_suffix(target.suffix + ".tmp")
    temporary.write_text(
        json.dumps(normalized.to_dict(include_secret=True), ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    temporary.replace(target)
    return target


def delete_model_settings(path: str | Path | None = None) -> bool:
    target = model_settings_path(path)
    if not target.exists():
        return False
    target.unlink()
    return True


def from_legacy_api_settings(settings: ApiSettings) -> ModelSettings:
    provider_id = safe_provider_id(settings.provider_name or settings.model or "legacy-model")
    provider = ProviderConfig(
        id=provider_id,
        name=settings.provider_name,
        base_url=settings.base_url,
        api_key=settings.api_key,
        model=settings.model,
        timeout_seconds=settings.timeout_seconds,
        enabled=settings.enabled,
    )
    return ModelSettings(
        providers=[provider],
        roles={role: provider.id for role in ROLE_NAMES},
        agent_enabled=bool(settings.use_for_generation),
        run_python_enabled=True,
    )


def list_providers(settings: ModelSettings | None = None) -> list[ProviderConfig]:
    return list((settings or load_model_settings()).providers)


def get_provider(
    role: str,
    settings: ModelSettings | None = None,
) -> ProviderConfig | None:
    data = settings or load_model_settings()
    by_id = {item.id: item for item in data.providers}
    selected_id = data.roles.get(str(role))
    selected = by_id.get(selected_id or "")
    if selected and selected.configured:
        return selected
    for provider in data.providers:
        if provider.configured:
            return provider
    return None


def get_role_api_settings(
    role: str,
    settings: ModelSettings | None = None,
    *,
    use_for_intent: bool = True,
    use_for_review: bool = True,
    use_for_generation: bool = True,
) -> ApiSettings | None:
    provider = get_provider(role, settings)
    if provider is None:
        return None
    return provider.to_api_settings(
        use_for_intent=use_for_intent,
        use_for_review=use_for_review,
        use_for_generation=use_for_generation,
    )


def chat(
    role: str,
    *,
    system_prompt: str,
    user_prompt: str,
    temperature: float = 0.1,
    max_tokens: int = 800,
    json_mode: bool = False,
    settings: ModelSettings | None = None,
) -> ApiCallResult:
    api_settings = get_role_api_settings(role, settings)
    if api_settings is None:
        return ApiCallResult(False, "", "未配置可用模型。", None, None)
    return chat_completion(
        api_settings,
        system_prompt=system_prompt,
        user_prompt=user_prompt,
        temperature=temperature,
        max_tokens=max_tokens,
        json_mode=json_mode,
    )


def chat_with_tools(
    role: str,
    *,
    messages: list[dict[str, Any]],
    tools: list[dict[str, Any]],
    tool_choice: str | dict[str, Any] = "auto",
    temperature: float = 0.1,
    max_tokens: int = 3000,
    settings: ModelSettings | None = None,
) -> ToolChatResult:
    api_settings = get_role_api_settings(role, settings)
    if api_settings is None:
        return ToolChatResult(False, "", [], "未配置可用模型。", None, None, None)
    return chat_completion_with_tools(
        api_settings,
        messages=messages,
        tools=tools,
        tool_choice=tool_choice,
        temperature=temperature,
        max_tokens=max_tokens,
    )


def test_provider(provider: ProviderConfig) -> ApiCallResult:
    return test_api_connection(provider.to_api_settings())


def legacy_api_settings_file() -> Path:
    return api_settings_path()


def safe_provider_id(value: str) -> str:
    raw = str(value or "provider").strip().lower()
    raw = re.sub(r"[^a-z0-9\u4e00-\u9fff_-]+", "-", raw)
    raw = raw.strip("-_") or "provider"
    return raw[:48]
