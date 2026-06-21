"""Local-only settings for an OpenAI-compatible custom API."""

from __future__ import annotations

import json
import os
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Any

from .io_utils import project_root


DEFAULT_TIMEOUT_SECONDS = 45


@dataclass
class ApiSettings:
    enabled: bool = False
    base_url: str = ""
    api_key: str = ""
    model: str = ""
    provider_name: str = "自定义模型"
    timeout_seconds: int = DEFAULT_TIMEOUT_SECONDS
    use_for_intent: bool = True
    use_for_review: bool = True

    def __post_init__(self) -> None:
        self.base_url = str(self.base_url or "").strip().rstrip("/")
        self.api_key = str(self.api_key or "").strip()
        self.model = str(self.model or "").strip()
        self.provider_name = str(self.provider_name or "自定义模型").strip() or "自定义模型"
        self.timeout_seconds = min(max(int(self.timeout_seconds), 5), 180)

    @property
    def configured(self) -> bool:
        return bool(self.enabled and self.base_url and self.api_key and self.model)

    def to_dict(self, include_secret: bool = True) -> dict[str, Any]:
        data = asdict(self)
        if not include_secret:
            data["api_key"] = mask_api_key(self.api_key)
        return data

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> "ApiSettings":
        allowed = {
            "enabled",
            "base_url",
            "api_key",
            "model",
            "provider_name",
            "timeout_seconds",
            "use_for_intent",
            "use_for_review",
        }
        return cls(**{key: value for key, value in data.items() if key in allowed})


def api_settings_path(path: str | Path | None = None) -> Path:
    override = os.getenv("AI_EXCEL_API_SETTINGS_FILE", "").strip()
    if path is not None:
        target = Path(path)
    elif override:
        target = Path(override)
    else:
        target = project_root() / "data" / "private" / "api_settings.json"
    return target.expanduser().resolve()


def load_api_settings(path: str | Path | None = None) -> ApiSettings:
    target = api_settings_path(path)
    if not target.exists():
        return ApiSettings()
    try:
        payload = json.loads(target.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError, TypeError):
        return ApiSettings()
    return ApiSettings.from_dict(payload) if isinstance(payload, dict) else ApiSettings()


def save_api_settings(settings: ApiSettings, path: str | Path | None = None) -> Path:
    target = api_settings_path(path)
    target.parent.mkdir(parents=True, exist_ok=True)
    temporary = target.with_suffix(target.suffix + ".tmp")
    temporary.write_text(
        json.dumps(settings.to_dict(include_secret=True), ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    temporary.replace(target)
    return target


def delete_api_settings(path: str | Path | None = None) -> bool:
    target = api_settings_path(path)
    if not target.exists():
        return False
    target.unlink()
    return True


def mask_api_key(api_key: str) -> str:
    key = str(api_key or "")
    if not key:
        return "未设置"
    if len(key) <= 8:
        return "•" * len(key)
    return f"{key[:3]}{'•' * min(12, len(key) - 7)}{key[-4:]}"
