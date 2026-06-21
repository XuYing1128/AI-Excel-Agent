"""Small non-destructive startup check for the local Streamlit app."""

from __future__ import annotations

import socket
import sys
import webbrowser

from excel_agent.services.runtime_compat import load_generation_service


APP_URL = "http://127.0.0.1:8501"


def port_is_open() -> bool:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as client:
        client.settimeout(0.3)
        return client.connect_ex(("127.0.0.1", 8501)) == 0


def main() -> int:
    try:
        load_generation_service()
    except RuntimeError as exc:
        print(f"启动检查失败：{exc}", file=sys.stderr)
        return 1
    if port_is_open():
        print("本地表格助手已经启动，正在打开现有页面。")
        webbrowser.open(APP_URL)
        return 2
    print("启动检查通过。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
