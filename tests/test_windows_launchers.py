from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]


def test_start_script_has_no_hidden_powershell_or_download_behavior():
    content = (ROOT / "start.bat").read_text(encoding="utf-8").lower()
    forbidden = [
        "powershell",
        "start-process",
        "windowstyle",
        "encodedcommand",
        "invoke-webrequest",
        "downloadstring",
        "downloadfile",
        "executionpolicy",
    ]
    assert all(token not in content for token in forbidden)
    assert ".venv\\scripts\\python.exe" in content
    # The launcher now references the local python via an %APP_PY% variable, so
    # assert the safe pip-install behaviour without pinning the exact inline path.
    assert "-m pip install --disable-pip-version-check -e ." in content
    assert "-m streamlit run app.py" in content
    assert "--server.address 127.0.0.1" in content
    assert "--server.headless false" in content


def test_install_script_uses_explicit_local_python_without_powershell():
    content = (ROOT / "install.bat").read_text(encoding="utf-8").lower()
    forbidden = [
        "powershell",
        "start-process",
        "windowstyle",
        "encodedcommand",
        "invoke-webrequest",
        "downloadstring",
        "downloadfile",
        "executionpolicy",
        "activate.bat",
    ]
    assert all(token not in content for token in forbidden)
    assert "python -m venv .venv" in content
    assert '".venv\\scripts\\python.exe" -m pip install' in content
