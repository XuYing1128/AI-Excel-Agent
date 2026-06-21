from excel_agent.api_settings import (
    ApiSettings,
    delete_api_settings,
    load_api_settings,
    mask_api_key,
    save_api_settings,
)


def test_api_settings_local_roundtrip(tmp_path):
    path = tmp_path / "private" / "api_settings.json"
    settings = ApiSettings(
        enabled=True,
        base_url="https://example.com/v1/",
        api_key="secret-key-123456",
        model="example-model",
        provider_name="测试接口",
    )
    save_api_settings(settings, path)
    loaded = load_api_settings(path)

    assert loaded.configured is True
    assert loaded.base_url == "https://example.com/v1"
    assert loaded.api_key == settings.api_key
    assert mask_api_key(loaded.api_key).endswith("3456")
    assert delete_api_settings(path) is True
    assert load_api_settings(path).configured is False
