"""多步智能体编排器。

模型只负责选择工具和提交结构化参数；Excel 文件仍由本地确定性函数生成。
"""

from __future__ import annotations

import json
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Any, Callable

from ...model_registry import load_model_settings
from ...rich_workbook_builder import build_rich_workbook, normalize_workbook_blueprint
from ...task_paths import TaskPaths, append_run_log_event
from ...task_spec import TaskSpec
from ...validators import validate_workbook
from ..custom_api_service import parse_json_object
from ... import model_registry


ProgressCallback = Callable[[str, str], None]


@dataclass
class AgentResult:
    success: bool
    output_file: str | None
    message: str
    error: str | None
    steps: int
    tool_calls: int
    mode: str = "agent_orchestrator"
    blueprint_file: str | None = None
    notices: list[str] | None = None

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


def run_agent(
    task_spec: TaskSpec,
    task_paths: TaskPaths,
    *,
    progress: ProgressCallback | None = None,
    max_steps: int = 12,
) -> AgentResult:
    """Run the builder-role model in a small, tool-driven loop."""

    settings = load_model_settings()
    if not settings.agent_enabled:
        return AgentResult(
            success=False,
            output_file=None,
            message="智能体编排未启用。",
            error="智能体编排未启用。",
            steps=0,
            tool_calls=0,
        )
    provider = model_registry.get_provider("builder", settings)
    if provider is None:
        return AgentResult(
            success=False,
            output_file=None,
            message="没有可用的 builder 模型。",
            error="没有可用的 builder 模型。",
            steps=0,
            tool_calls=0,
        )

    _progress(progress, "model", "正在由智能体分析任务并选择本地工具……")
    blueprint_path = task_paths.task_dir / "agent_workbook_blueprint.json"
    messages = [
        {"role": "system", "content": _system_prompt()},
        {
            "role": "user",
            "content": json.dumps(
                {
                    "用户需求": task_spec.user_goal,
                    "已确认任务": _safe_task_spec(task_spec),
                    "文件摘要": task_spec.options.get("input_data_profile", {}),
                    "关键约束": [
                        "不得输出 Markdown 代替工具调用。",
                        "不得要求模型直接写文件。",
                        "必须通过 build_workbook 或 finish_task 推进。",
                        "计算列要写 Excel 公式模板，不写静态结果。",
                    ],
                },
                ensure_ascii=False,
            ),
        },
    ]
    append_run_log_event(
        task_paths,
        event="agent_orchestrator_started",
        status="running",
        details={"provider": provider.name, "model": provider.model},
    )
    built = False
    tool_calls = 0
    last_error = ""
    stalled = 0

    for step in range(1, max_steps + 1):
        _progress(progress, "model", f"智能体第 {step} 步正在判断下一步操作……")
        response = model_registry.chat_with_tools(
            "builder",
            settings=settings,
            messages=messages,
            tools=_tool_schemas(),
            tool_choice="auto",
            temperature=0.05,
            max_tokens=8000,
        )
        if not response.success:
            last_error = response.error or "模型调用失败。"
            break
        assistant_message = dict(response.message or {"role": "assistant"})
        assistant_message.setdefault("role", "assistant")
        messages.append(assistant_message)

        calls = response.tool_calls
        if not calls and response.content:
            fallback_call = _json_instruction_to_call(response.content)
            calls = [fallback_call] if fallback_call else []

        if not calls:
            stalled += 1
            last_error = "模型没有调用工具。"
            messages.append(
                {
                    "role": "user",
                    "content": "请不要只解释。现在必须调用 build_workbook、validate_workbook 或 finish_task。",
                }
            )
            if stalled >= 2:
                break
            continue

        completed = False
        for call in calls:
            tool_calls += 1
            result = _dispatch_tool(
                call.name,
                call.arguments,
                task_spec,
                task_paths,
                blueprint_path,
                progress,
            )
            last_error = "" if result.get("ok") else str(result.get("error") or result.get("summary") or "")
            if call.name == "build_workbook" and result.get("ok"):
                built = True
            if call.name == "finish_task" and result.get("ok"):
                completed = True
            messages.append(
                {
                    "role": "tool",
                    "tool_call_id": call.id,
                    "name": call.name,
                    "content": json.dumps(result, ensure_ascii=False),
                }
            )
            append_run_log_event(
                task_paths,
                event="agent_tool_called",
                status="success" if result.get("ok") else "warning",
                details={
                    "step": step,
                    "tool": call.name,
                    "summary": result.get("summary"),
                    "error": result.get("error"),
                },
            )
        if completed:
            break
        if built:
            messages.append(
                {
                    "role": "user",
                    "content": "文件已生成。请用 validate_workbook 检查；若通过再调用 finish_task。",
                }
            )

    if task_paths.output_file.exists() and (built or tool_calls > 0):
        report = validate_workbook(task_paths.output_file)
        if report.get("status") in {"pass", "warn"}:
            result = AgentResult(
                success=True,
                output_file=str(task_paths.output_file),
                message="智能体已调用本地工具生成工作簿。",
                error=None,
                steps=step,
                tool_calls=tool_calls,
                blueprint_file=str(blueprint_path) if blueprint_path.exists() else None,
                notices=["已由智能体选择本地工具生成；文件仍经过确定性校验。"],
            )
            append_run_log_event(
                task_paths,
                event="agent_orchestrator_completed",
                status="success",
                details=result.to_dict(),
            )
            return result

    result = AgentResult(
        success=False,
        output_file=None,
        message="智能体未能生成可用文件。",
        error=last_error or "智能体未完成 finish_task。",
        steps=locals().get("step", 0),
        tool_calls=tool_calls,
        blueprint_file=str(blueprint_path) if blueprint_path.exists() else None,
        notices=[],
    )
    append_run_log_event(
        task_paths,
        event="agent_orchestrator_failed",
        status="error",
        details=result.to_dict(),
    )
    return result


def _dispatch_tool(
    name: str,
    args: dict[str, Any],
    task_spec: TaskSpec,
    task_paths: TaskPaths,
    blueprint_path: Path,
    progress: ProgressCallback | None,
) -> dict[str, Any]:
    if name == "build_workbook":
        blueprint = args.get("blueprint", args)
        try:
            normalized = normalize_workbook_blueprint(blueprint)
            if task_spec.include_charts and not _workbook_has_charts(normalized):
                normalized["require_charts"] = True
            blueprint_path.write_text(
                json.dumps(normalized, ensure_ascii=False, indent=2),
                encoding="utf-8",
            )
            _progress(progress, "build", "智能体已提交方案，正在写入 Excel 文件……")
            build_rich_workbook(
                normalized,
                task_paths.output_file,
                require_charts=task_spec.include_charts,
            )
            report = validate_workbook(task_paths.output_file)
            return {
                "ok": report.get("status") in {"pass", "warn"},
                "summary": f"已生成文件，校验状态：{report.get('status')}",
                "data": {"validation_status": report.get("status")},
                "artifacts": [str(task_paths.output_file), str(blueprint_path)],
            }
        except Exception as exc:
            return {"ok": False, "summary": "生成失败。", "error": f"{type(exc).__name__}: {exc}"}
    if name == "validate_workbook":
        if not task_paths.output_file.exists():
            return {"ok": False, "summary": "还没有生成工作簿。", "error": "output_missing"}
        report = validate_workbook(task_paths.output_file)
        return {
            "ok": report.get("status") in {"pass", "warn"},
            "summary": f"校验状态：{report.get('status')}",
            "data": {
                "status": report.get("status"),
                "error_count": len(report.get("errors", [])),
                "warning_count": len(report.get("warnings", [])),
            },
            "artifacts": [str(task_paths.output_file)],
        }
    if name == "finish_task":
        ok = task_paths.output_file.exists()
        return {
            "ok": ok,
            "summary": str(args.get("summary") or ("任务已完成。" if ok else "还没有生成文件。")),
            "artifacts": [str(task_paths.output_file)] if ok else [],
        }
    return {"ok": False, "summary": f"未知工具：{name}", "error": "unknown_tool"}


def _json_instruction_to_call(content: str) -> Any | None:
    try:
        payload = parse_json_object(content)
    except ValueError:
        return None
    tool_name = str(payload.get("tool") or payload.get("name") or "")
    args = payload.get("args") or payload.get("arguments") or {}
    if not tool_name or not isinstance(args, dict):
        return None

    class _Call:
        id = "json_instruction"
        name = tool_name
        arguments = args

    return _Call()


def _tool_schemas() -> list[dict[str, Any]]:
    return [
        {
            "type": "function",
            "function": {
                "name": "build_workbook",
                "description": "提交完整工作簿方案，由本地工具生成 xlsx。",
                "parameters": {
                    "type": "object",
                    "properties": {
                        "blueprint": {
                            "type": "object",
                            "description": "工作簿方案。单表或包含 sheets 的多表结构均可。",
                        }
                    },
                    "required": ["blueprint"],
                },
            },
        },
        {
            "type": "function",
            "function": {
                "name": "validate_workbook",
                "description": "检查已生成工作簿能否打开、公式和结构是否有明显问题。",
                "parameters": {"type": "object", "properties": {}},
            },
        },
        {
            "type": "function",
            "function": {
                "name": "finish_task",
                "description": "确认任务完成。只有本地文件已生成并通过校验后才能调用。",
                "parameters": {
                    "type": "object",
                    "properties": {"summary": {"type": "string"}},
                },
            },
        },
    ]


def _system_prompt() -> str:
    return (
        "你是本地 Excel 表格生成智能体。你不能直接写文件，也不能输出 Markdown 当作结果。"
        "你必须选择工具：build_workbook 生成，validate_workbook 检查，finish_task 完成。"
        "所有计算列必须写 Excel 公式模板，并考虑 IFERROR、空值和除零。"
        "用户要求图表时必须在方案里包含 charts；用户给出的列、标题、分组、小计、总计优先。"
        "不要复述全量数据，按已确认任务和文件摘要设计结构。"
    )


def _safe_task_spec(task_spec: TaskSpec) -> dict[str, Any]:
    payload = task_spec.to_dict()
    payload["input_files"] = [Path(item).name for item in task_spec.input_files]
    payload["template_files"] = [Path(item).name for item in task_spec.template_files]
    payload["options"] = {
        key: value
        for key, value in payload.get("options", {}).items()
        if key not in {"pasted_data_text"}
    }
    return payload


def _workbook_has_charts(blueprint: dict[str, Any]) -> bool:
    if blueprint.get("charts"):
        return True
    return any(bool(sheet.get("charts")) for sheet in blueprint.get("sheets", []) if isinstance(sheet, dict))


def _progress(progress: ProgressCallback | None, stage: str, message: str) -> None:
    if progress:
        progress(stage, message)

