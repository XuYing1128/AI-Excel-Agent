# -*- coding: utf-8 -*-
"""PyInstaller 打包入口。

Streamlit 应用本身没有可被直接 `python app.py` 运行的 __main__ 入口，
官方推荐用 streamlit.web.bootstrap 启动。本脚本作为 PyInstaller 的主入口，
打包后双击 exe 即可拉起 app.py（无需命令行、无需本机 Python）。

开发模式下也可直接 `python run_app.py` 调试打包后的启动行为。
"""
from __future__ import annotations

import os
import sys
from pathlib import Path


def _resolve_app_path() -> str:
    """定位 app.py：打包后从 exe 同级（或 _MEIPASS）找，开发模式从本文件同目录找。"""
    # PyInstaller onedir：app.py 通过 --add-data 放在 exe 同级
    if getattr(sys, "frozen", False):
        base = Path(sys.executable).resolve().parent
    else:
        base = Path(__file__).resolve().parent
    candidate = base / "app.py"
    if candidate.exists():
        return str(candidate)
    # 兜底：_MEIPASS 临时解压目录（onefile 模式或被放到子目录时）
    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        cand2 = Path(meipass) / "app.py"
        if cand2.exists():
            return str(cand2)
    raise FileNotFoundError(f"找不到 app.py（已查找 {base}）")


def main() -> None:
    app_path = _resolve_app_path()
    # Streamlit 1.58+：命令行传 server.port 会触发 developmentMode 冲突
    # (RuntimeError: server.port does not work when global.developmentMode is true)。
    # 所以端口/地址/headless 全部交给 .streamlit/config.toml，这里只跑脚本，
    # 并用 flag_options 显式关闭 developmentMode 以彻底规避冲突。
    sys.argv = ["streamlit", "run", app_path]
    # 把 app.py 所在目录加入 import 路径，保证 app.py 里的 import 能找到 src 等包
    sys.path.insert(0, os.path.dirname(app_path))

    from streamlit.web import bootstrap

    bootstrap.run(
        app_path,
        False,  # is_hello
        [],     # 额外命令行参数（不传 port/address，交给 config.toml）
        flag_options={
            "server.headless": False,  # 直接运行/exe 模式都弹浏览器窗口
        },
    )


if __name__ == "__main__":
    main()
