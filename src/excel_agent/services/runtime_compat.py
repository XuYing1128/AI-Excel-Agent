"""Runtime compatibility checks for long-running Streamlit processes."""

from __future__ import annotations

import importlib
import inspect
from types import ModuleType


REQUIRED_GENERATION_PARAMETERS = {"api_settings", "progress"}


def load_generation_service() -> ModuleType:
    """Reload a stale cached service module after an in-place project update."""

    from . import generation_service

    module = generation_service
    if not _is_current(module):
        module = importlib.reload(module)
    if not _is_current(module):
        signature = inspect.signature(module.generate_from_task_spec)
        raise RuntimeError(
            "表格生成服务版本不一致。当前参数为："
            + "、".join(signature.parameters)
            + "。请关闭旧页面后重新启动。"
        )
    return module


def _is_current(module: ModuleType) -> bool:
    if getattr(module, "GENERATION_API_VERSION", 0) < 2:
        return False
    signature = inspect.signature(module.generate_from_task_spec)
    return REQUIRED_GENERATION_PARAMETERS.issubset(signature.parameters)
