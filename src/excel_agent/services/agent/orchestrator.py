"""多步智能体编排器。

模型只负责选择工具和提交结构化参数；Excel 文件仍由本地确定性函数生成。
"""

from __future__ import annotations

import json
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Any, Callable

from ...model_registry import load_model_settings
from ...rich_workbook_builder import normalize_workbook_blueprint
from ...task_paths import TaskPaths, append_run_log_event
from ...task_spec import TaskSpec
from ...validators import validate_workbook
from ..custom_api_service import parse_json_object
from ... import model_registry
from .tools import ToolContext, tool_map, tool_schemas


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
            message="自动多步生成未启用。",
            error="自动多步生成未启用。",
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

    _progress(progress, "model", "正在分析任务并选择合适的本地工具……")
    blueprint_path = task_paths.task_dir / "agent_workbook_blueprint.json"
    tool_context = ToolContext(
        task_spec=task_spec,
        task_paths=task_paths,
        progress=progress,
        run_python_enabled=settings.run_python_enabled,
    )
    matched_skills = _matched_skills(task_spec)
    messages = [
        {"role": "system", "content": _system_prompt(matched_skills)},
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
        details={
            "provider": provider.name,
            "model": provider.model,
            "skills": [item.get("name") for item in matched_skills],
        },
    )
    built = False
    tool_calls = 0
    last_error = ""
    stalled = 0

    for step in range(1, max_steps + 1):
        _progress(progress, "model", f"第 {step} 步：正在判断下一步操作……")
        response = model_registry.chat_with_tools(
            "builder",
            settings=settings,
            messages=messages,
            tools=tool_schemas(tool_context),
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
                tool_context,
                blueprint_path,
            )
            last_error = "" if result.get("ok") else str(result.get("error") or result.get("summary") or "")
            if call.name in {"build_workbook", "build_rich_workbook"} and result.get("ok"):
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
                message="已通过本地工具生成工作簿。",
                error=None,
                steps=step,
                tool_calls=tool_calls,
                blueprint_file=str(blueprint_path) if blueprint_path.exists() else None,
                notices=["已自动选择本地工具生成；文件仍经过确定性校验。"],
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
        message="自动多步生成未能生成可用文件。",
        error=last_error or "自动多步生成未完成。",
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
    ctx: ToolContext,
    blueprint_path: Path,
) -> dict[str, Any]:
    tools = tool_map(ctx)
    tool = tools.get(name)
    if tool is None:
        return {"ok": False, "summary": f"未知工具：{name}", "error": "unknown_tool"}
    if name in {"build_workbook", "build_rich_workbook"}:
        try:
            normalized = normalize_workbook_blueprint(args.get("blueprint", args))
            if ctx.task_spec.include_charts and not _workbook_has_charts(normalized):
                normalized["require_charts"] = True
            blueprint_path.write_text(
                json.dumps(normalized, ensure_ascii=False, indent=2),
                encoding="utf-8",
            )
            args = {"blueprint": normalized}
        except Exception as exc:
            return {"ok": False, "summary": "生成失败。", "error": f"{type(exc).__name__}: {exc}"}
    return tool.handler(args, ctx).to_dict()


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


def _system_prompt(skills: list[dict[str, str]] | None = None) -> str:
    base = (
        "你是本地 Excel 表格生成智能体。你不能直接写文件，也不能输出 Markdown 当作结果。"
        "你必须选择已提供的工具推进任务，常用顺序是生成/处理 → validate_workbook 检查 → finish_task 完成。"
        "所有计算列必须写 Excel 公式模板，并考虑 IFERROR、空值和除零。"
        "用户要求图表时必须在方案里包含 charts；用户给出的列、标题、分组、小计、总计优先。"
        "不要复述全量数据，按已确认任务和文件摘要设计结构。"
    )
    if not skills:
        return base
    blocks = []
    for item in skills[:2]:
        blocks.append(f"\n\n【可用技能：{item['name']}】\n{item['content'][:5000]}")
    return base + "".join(blocks)


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


def _matched_skills(task_spec: TaskSpec) -> list[dict[str, str]]:
    try:
        from skills.registry import match_skills

        return [
            {"name": item.name, "content": item.content}
            for item in match_skills(task_spec)[:2]
        ]
    except Exception:
        return []
