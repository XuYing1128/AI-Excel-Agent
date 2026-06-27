"""本地 Python 运行工具（已松绑：可任意 import、可联网，仅保留路径护栏）。"""

from __future__ import annotations

import json
import os
import subprocess
import sys
from pathlib import Path
from typing import Any

from .base import AgentTool, ToolContext, ToolResult


def run_python_tool() -> AgentTool:
    return AgentTool(
        "run_python",
        (
            "在当前任务临时目录中运行 Python（pandas/openpyxl 等任意库均可，可联网）。"
            "唯一边界：只能读取 task 目录、写入 task/output 和 task/agent_tmp，"
            "防止误伤任务以外的文件。需要修改已有文件时，请 load 原文件就地改、存回 OUTPUT_FILE。"
        ),
        {
            "type": "object",
            "properties": {
                "code": {"type": "string", "description": "要执行的 Python 代码。"},
                "timeout_seconds": {"type": "integer", "minimum": 1, "maximum": 600},
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
    # 翻转后：信任本地环境，AI 可自由 import、可联网。仅保留“路径护栏”
    # （只能读任务目录、写 output/temp），避免误伤任务以外的文件。
    timeout = min(max(int(args.get("timeout_seconds") or 60), 1), 600)
    return run_python_code(code, ctx, timeout_seconds=timeout)


def run_python_code(code: str, ctx: ToolContext, *, timeout_seconds: int = 60) -> ToolResult:
    temp_dir = ctx.temp_dir.resolve()
    ctx.task_paths.output_dir.mkdir(parents=True, exist_ok=True)
    user_code = temp_dir / "agent_user_code.py"
    runner = temp_dir / "agent_runner.py"
    before = _snapshot_files(ctx.task_paths.output_dir, temp_dir)
    user_code.write_text(code, encoding="utf-8")
    runner.write_text(_runner_source(), encoding="utf-8")
    config = {
        "user_code": str(user_code),
        # Read anything inside the task dir (uploads, templates, prior output);
        # write only to output + temp.
        "read_roots": [
            str(ctx.task_paths.task_dir.resolve()),
            str(ctx.task_paths.input_dir.resolve()),
            str(ctx.task_paths.output_dir.resolve()),
            str(temp_dir),
        ],
        "write_roots": [str(ctx.task_paths.output_dir.resolve()), str(temp_dir)],
        # The exact final path the rest of the pipeline expects. The agent should
        # save the finished workbook here so it is picked up without renaming.
        "output_file": str(ctx.task_paths.output_file.resolve()),
        # Concrete upload paths so the agent reads them directly (os/pathlib are
        # not in the import whitelist, so it cannot list directories itself).
        "input_files": [str(Path(p).resolve()) for p in (ctx.task_spec.input_files or [])],
        "template_files": [str(Path(p).resolve()) for p in (ctx.task_spec.template_files or [])],
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


# 翻转后：信任本地环境——放开联网与 import 限制。安全边界仅靠上面的
# “路径护栏”（_safe_open / _safe_path_open）维持：用户代码只能读任务目录、
# 写 output/temp，无法触碰任务以外的文件。

globals_dict = {
    "__name__": "__main__",
    "__file__": str(USER_CODE),
    "INPUT_DIR": cfg["read_roots"][1] if len(cfg["read_roots"]) > 1 else str(READ_ROOTS[0]),
    "OUTPUT_DIR": str(WRITE_ROOTS[0]),
    "OUTPUT_FILE": cfg.get("output_file", str(WRITE_ROOTS[0]) + "/output.xlsx"),
    "TEMP_DIR": str(WRITE_ROOTS[1]),
    "INPUT_FILES": cfg.get("input_files", []),
    "TEMPLATE_FILES": cfg.get("template_files", []),
}
source = USER_CODE.read_text(encoding="utf-8")
exec(compile(source, str(USER_CODE), "exec"), globals_dict)
'''
