"""测试环境隔离。

项目运行时会读取 data/private 下的真实接口配置；单元测试必须改用临时路径，
避免测试结果受本机 API key、网络状态或用户私有设置影响。
"""

import pytest


@pytest.fixture(autouse=True)
def isolate_private_settings(tmp_path, monkeypatch):
    private_dir = tmp_path / "private"
    monkeypatch.setenv("AI_EXCEL_API_SETTINGS_FILE", str(private_dir / "api_settings.json"))
    monkeypatch.setenv("AI_EXCEL_MODEL_SETTINGS_FILE", str(private_dir / "model_settings.json"))
    monkeypatch.setenv("AI_EXCEL_MEMORY_DB", str(private_dir / "memory.db"))
