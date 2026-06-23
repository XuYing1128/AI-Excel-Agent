"""受限 Python 子进程工具。"""

from __future__ import annotations

import ast
import json
import os
import subprocess
import sys
from pathlib import Path
from typing import Any

from .base import AgentTool, ToolContext, ToolResult


ALLOWED_IMPORTS = {
    "pandas",
    "numpy",
    "openpyxl",
    "datetime",
    "re",
    "math",
    "json",
    "collections",
    "itertools",
    "statistics",
    "decimal",
    "csv",
}


def run_python_tool() -> AgentTool:
    return AgentTool(
        "run_python",
        (
            "在当前任务临时目录中运行受限 Python。只能读取 task/input，"
            "只能写入 task/output 和 task/agent_tmp；禁网络、限时、限制用户代码 import。"
        ),
        {
            "type": "object",
            "properties": {
                "code": {"type": "string", "description": "要执行的 Python 代码。"},
                "timeout_seconds": {"type": "integer", "minimum": 1, "maximum": 120},
            },
            "required": ["code"],
        },
        _handle_run_python,
    )


def _handle_run_python(args: dict[str, Any], ctx: ToolContext) -> ToolResult:
    if not ctx.run_python_enabled:
        return ToolResult(False, "安全脚本工具已关闭。", error="run_python_disabled")
    code = str(args.get("code") or "").strip()
    if not code:
        return ToolResult(False, "没有提供代码。", error="empty_code")
    invalid = _invalid_imports(code)
    if invalid:
        return ToolResult(
            False,
            "代码包含不允许的 import。",
            data={"invalid_imports": invalid, "allowed_imports": sorted(ALLOWED_IMPORTS)},
            error="import_not_allowed",
        )
    timeout = min(max(int(args.get("timeout_seconds") or 60), 1), 120)
    return run_python_code(code, ctx, timeout_seconds=timeout)


def run_python_code(code: str, ctx: ToolContext, *, timeout_seconds: int = 60) -> ToolResult:
    temp_dir = ctx.temp_dir.resolve()
    user_code = temp_dir / "agent_user_code.py"
    runner = temp_dir / "agent_runner.py"
    before = _snapshot_files(ctx.task_paths.output_dir, temp_dir)
    user_code.write_text(code, encoding="utf-8")
    runner.write_text(_runner_source(), encoding="utf-8")
    config = {
        "user_code": str(user_code),
        "allowed_imports": sorted(ALLOWED_IMPORTS),
        "read_roots": [
            str(ctx.task_paths.input_dir.resolve()),
            str(ctx.task_paths.output_dir.resolve()),
            str(temp_dir),
        ],
        "write_roots": [str(ctx.task_paths.output_dir.resolve()), str(temp_dir)],
    }
    env = dict(os.environ)
    env["PYTHONIOENCODING"] = "utf-8"
    env["TMP"] = str(temp_dir)
    env["TEMP"] = str(temp_dir)
    env["TMPDIR"] = str(temp_dir)
    try:
        completed = subprocess.run(
            [sys.executable, str(runner), json.dumps(config, ensure_ascii=False)],
            cwd=temp_dir,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            env=env,
            timeout=timeout_seconds,
        )
    except subprocess.TimeoutExpired as exc:
        return ToolResult(
            False,
            f"脚本超过 {timeout_seconds} 秒，已终止。",
            data={"stdout": exc.stdout or "", "stderr": exc.stderr or ""},
            error="timeout",
        )
    after = _snapshot_files(ctx.task_paths.output_dir, temp_dir)
    new_files = sorted(str(path) for path in after - before)
    ok = completed.returncode == 0
    return ToolResult(
        ok,
        "脚本执行完成。" if ok else "脚本执行失败。",
        data={
            "returncode": completed.returncode,
            "stdout": completed.stdout[-4000:],
            "stderr": completed.stderr[-4000:],
            "new_files": new_files,
        },
        artifacts=new_files,
        error=None if ok else "python_failed",
    )


def _snapshot_files(*roots: Path) -> set[Path]:
    files: set[Path] = set()
    for root in roots:
        if root.exists():
            files.update(path.resolve() for path in root.rglob("*") if path.is_file())
    return files


def _invalid_imports(code: str) -> list[str]:
    try:
        tree = ast.parse(code)
    except SyntaxError as exc:
        return [f"语法错误: {exc.msg}"]
    invalid: list[str] = []
    for node in ast.walk(tree):
        if isinstance(node, ast.Import):
            for alias in node.names:
                top = alias.name.split(".")[0]
                if top not in ALLOWED_IMPORTS and top not in invalid:
                    invalid.append(top)
        elif isinstance(node, ast.ImportFrom):
            top = (node.module or "").split(".")[0]
            if top and top not in ALLOWED_IMPORTS and top not in invalid:
                invalid.append(top)
    return invalid


def _runner_source() -> str:
    return r'''
from __future__ import annotations

import builtins
import inspect
import json
import os
import pathlib
import socket
import sys


cfg = json.loads(sys.argv[1])
USER_CODE = pathlib.Path(cfg["user_code"]).resolve()
ALLOWED_IMPORTS = set(cfg["allowed_imports"])
READ_ROOTS = [pathlib.Path(p).resolve() for p in cfg["read_roots"]]
WRITE_ROOTS = [pathlib.Path(p).resolve() for p in cfg["write_roots"]]


def _inside(path, roots):
    resolved = pathlib.Path(path).resolve()
    for root in roots:
        try:
            if resolved == root or resolved.is_relative_to(root):
                return True
        except AttributeError:
            if str(resolved).startswith(str(root)):
                return True
    return False


def _guard_path(file, mode="r"):
    text_mode = str(mode or "r")
    roots = WRITE_ROOTS if any(ch in text_mode for ch in "wax+") else READ_ROOTS
    if not _inside(file, roots):
        raise PermissionError(f"path_escape: {file}")


_real_open = builtins.open


def _safe_open(file, mode="r", *args, **kwargs):
    _guard_path(file, mode)
    return _real_open(file, mode, *args, **kwargs)


builtins.open = _safe_open
_real_path_open = pathlib.Path.open


def _safe_path_open(self, mode="r", *args, **kwargs):
    _guard_path(self, mode)
    return _real_path_open(self, mode, *args, **kwargs)


pathlib.Path.open = _safe_path_open


class _BlockedSocket:
    def __init__(self, *args, **kwargs):
        raise RuntimeError("run_python 禁止网络访问")


socket.socket = _BlockedSocket

_real_import = builtins.__import__


def _safe_import(name, globals=None, locals=None, fromlist=(), level=0):
    top = str(name).split(".")[0]
    # Only block imports initiated directly by user code; library internals may
    # import their own dependencies after an allowed top-level package is loaded.
    frame = inspect.currentframe()
    caller = frame.f_back if frame else None
    user_in_stack = False
    depth = 0
    while caller is not None and depth < 4:
        if caller.f_code.co_filename == str(USER_CODE):
            user_in_stack = True
            break
        caller = caller.f_back
        depth += 1
    if user_in_stack:
        if top not in ALLOWED_IMPORTS:
            raise ImportError(f"不允许导入模块: {top}")
    return _real_import(name, globals, locals, fromlist, level)


builtins.__import__ = _safe_import

globals_dict = {
    "__name__": "__main__",
    "__file__": str(USER_CODE),
    "INPUT_DIR": str(READ_ROOTS[0]),
    "OUTPUT_DIR": str(WRITE_ROOTS[0]),
    "TEMP_DIR": str(WRITE_ROOTS[1]),
}
source = USER_CODE.read_text(encoding="utf-8")
exec(compile(source, str(USER_CODE), "exec"), globals_dict)
'''
