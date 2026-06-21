import tomllib
from pathlib import Path


def test_streamlit_developer_toolbar_and_telemetry_are_hidden():
    path = Path(__file__).resolve().parents[1] / ".streamlit" / "config.toml"
    config = tomllib.loads(path.read_text(encoding="utf-8"))

    assert config["client"]["toolbarMode"] == "minimal"
    assert config["client"]["showErrorDetails"] == "none"
    assert config["ui"]["hideTopBar"] is True
    assert config["browser"]["gatherUsageStats"] is False
    assert config["server"]["address"] == "127.0.0.1"
