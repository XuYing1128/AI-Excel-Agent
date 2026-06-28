from excel_agent.api_settings import ApiSettings, save_api_settings
from excel_agent.model_registry import (
    ROLE_NAMES,
    ModelSettings,
    ProviderConfig,
    get_provider,
    get_role_api_settings,
    load_model_settings,
    save_model_settings,
)


def test_model_settings_migrates_legacy_api_settings(tmp_path):
    legacy_path = tmp_path / "private" / "api_settings.json"
    model_path = tmp_path / "private" / "model_settings.json"
    save_api_settings(
        ApiSettings(
            enabled=True,
            base_url="https://example.com/v1",
            api_key="secret",
            model="model-a",
            provider_name="旧接口",
        ),
        legacy_path,
    )

    settings = load_model_settings(model_path, legacy_api_path=legacy_path)

    assert model_path.exists()
    assert len(settings.providers) == 1
    assert settings.providers[0].name == "旧接口"
    assert set(settings.roles) == set(ROLE_NAMES)
    assert get_provider("builder", settings).model == "model-a"


def test_model_settings_roundtrip_and_role_fallback(tmp_path):
    path = tmp_path / "model_settings.json"
    first = ProviderConfig(
        id="disabled",
        name="停用",
        base_url="https://disabled.example/v1",
        api_key="secret",
        model="off",
        enabled=False,
    )
    second = ProviderConfig(
        id="qwen",
        name="千问",
        base_url="https://qwen.example/v1",
        api_key="secret2",
        model="qwen-plus",
        timeout_seconds=180,
    )
    save_model_settings(
        ModelSettings(
            providers=[first, second],
            roles={"planner": "missing", "builder": "qwen"},
            agent_enabled=False,
            run_python_enabled=False,
        ),
        path,
    )

    loaded = load_model_settings(path)
    assert loaded.agent_enabled is False
    assert loaded.run_python_enabled is False
    assert get_provider("planner", loaded).id == "qwen"
    assert get_provider("builder", loaded).id == "qwen"

    role_settings = get_role_api_settings("builder", loaded)
    assert role_settings is not None
    assert role_settings.provider_name == "千问"
    assert role_settings.timeout_seconds == 180

