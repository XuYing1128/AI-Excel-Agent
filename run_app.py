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


def _startup_log_path() -> Path:
    """启动日志写到 exe 同级目录（打包后）或本文件同目录（开发模式）。"""
    if getattr(sys, "frozen", False):
        base = Path(sys.executable).resolve().parent
    else:
        base = Path(__file__).resolve().parent
    return base / "启动日志.txt"


def _write_startup_log(message: str) -> None:
    """把启动信息/报错追加写到启动日志.txt，方便无控制台时排查。"""
    import datetime
    try:
        log = _startup_log_path()
        ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with log.open("a", encoding="utf-8") as f:
            f.write(f"[{ts}] {message}\n")
    except Exception:
        pass  # 日志写入本身不能再抛异常，否则掩盖原始错误


def main() -> None:
    _write_startup_log("==== 程序启动 ====")
    try:
        app_path = _resolve_app_path()
        _write_startup_log(f"定位到 app.py: {app_path}")
    except Exception as exc:
        _write_startup_log(f"定位 app.py 失败: {exc}")
        _crash(f"找不到应用主程序 app.py：{exc}")
        return

    try:
        # Streamlit 1.58+：命令行传 server.port 会触发 developmentMode 冲突
        # (RuntimeError: server.port does not work when global.developmentMode is true)。
        # 所以用 flag_options 显式指定所有 server 参数——既规避 developmentMode 冲突，
        # 又确保 Streamlit 的 "Local URL" 显示和自动弹出浏览器用对端口（否则它默认
        # 显示 localhost:3000，若 3000 被别的程序如 Docker 占用，用户会被误导）。
        sys.argv = ["streamlit", "run", app_path]
        # 把 app.py 所在目录加入 import 路径，保证 app.py 里的 import 能找到 src 等包
        sys.path.insert(0, os.path.dirname(app_path))

        from streamlit.web import bootstrap

        _write_startup_log("正在启动 Streamlit 服务，端口 8501...")
        bootstrap.run(
            app_path,
            False,  # is_hello
            [],     # 额外命令行参数（不传 port/address，全部走 flag_options）
            flag_options={
                "server.headless": False,     # 直接运行/exe 模式都弹浏览器窗口
                "server.port": 8501,          # 固定端口，避免被默认 3000 误导
                "server.address": "127.0.0.1",  # 仅本机回环，不暴露到局域网
                "global.developmentMode": False,  # 规避 server.port 的 developmentMode 冲突
            },
        )
    except Exception as exc:
        import traceback
        tb = traceback.format_exc()
        _write_startup_log(f"启动失败: {exc}\n{tb}")
        _crash(f"启动失败：{exc}\n\n{tb}")


def _crash(message: str) -> None:
    """启动崩溃时：黑窗口打印醒目中文提示 + 等待用户看完再关。

    注意：Windows 控制台默认 GBK 编码，emoji/部分 Unicode 会触发
    UnicodeEncodeError，因此强制 stdout 用 UTF-8 重配，并用纯文本符号。
    """
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
    print("")
    print("=" * 60)
    print("[错误] 程序启动失败")
    print("=" * 60)
    print(message)
    print("=" * 60)
    print("详细错误已写入「启动日志.txt」，可把该文件发给开发者排查。")
    print("按回车键关闭本窗口...")
    try:
        input()
    except EOFError:
        pass


if __name__ == "__main__":
    main()
