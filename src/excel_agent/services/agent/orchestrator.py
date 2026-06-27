"""多步智能体编排器。

模型只负责选择工具和提交结构化参数；Excel 文件仍由本地确定性函数生成。
"""

from __future__ import annotations

import json
import shutil
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Any, Callable

from ...model_registry import load_model_settings
from ...rich_workbook_builder import normalize_workbook_blueprint
from ...task_paths import TaskPaths, append_run_log_event
from ...task_spec import TaskSpec
from ...validators import validate_workbook
from ..custom_api_service import parse_json_object
from ..recalc import describe_error_cells, recalc_workbook
from ... import model_registry
from .tools import ToolContext, tool_map, tool_schemas


ProgressCallback = Callable[[str, str], None]
# 交卷前真算关卡最多触发几轮“真算→退回修复”，超过则不放行（避免反复真算拖太久）。
MAX_RECALC_FIX = 3


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
    truncation_retries = 0
    recalc_attempts = 0
    last_recalc: dict[str, Any] | None = None

    for step in range(1, max_steps + 1):
        _progress(progress, "model", f"第 {step} 步：正在判断下一步操作……")
        response = model_registry.chat_with_tools(
            "builder",
            settings=settings,
            messages=messages,
            tools=tool_schemas(tool_context),
            tool_choice="auto",
            temperature=0.05,
            max_tokens=16000,
        )
        if not response.success:
            # 截断（推理把输出额度占满）不直接放弃：提示模型少分析、直接动手再试。
            if _is_truncation(response.error) and truncation_retries < 2:
                truncation_retries += 1
                messages.append(
                    {
                        "role": "user",
                        "content": "上一步因思考过长被截断。请不要输出大段分析，"
                        "直接调用 run_python 写最小可用代码完成当前这一步。",
                    }
                )
                continue
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
        finish_requested = False
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
                finish_requested = True
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

        if finish_requested:
            # 交卷前真算关卡：用本机 LibreOffice 真算一遍，发现 #VALUE!/循环引用就退回让模型修，
            # 修不好不放行（静态验证器只看公式长相，抓不到这类“真算才发作”的错误）。
            gate = recalc_workbook(task_paths.output_file)
            last_recalc = gate
            if gate.get("ok"):
                completed = True
            elif recalc_attempts < MAX_RECALC_FIX:
                recalc_attempts += 1
                error_cells = gate.get("error_cells", [])
                _progress(
                    progress,
                    "build",
                    f"真算发现 {len(error_cells)} 处公式报错，正在让模型修正……",
                )
                messages.append(
                    {
                        "role": "user",
                        "content": (
                            "真算（按 Excel 实际计算）发现这些单元格报错："
                            + describe_error_cells(error_cells)
                            + "。这类报错几乎都是公式逻辑问题：① 跨表统计漏了『工作表名!』前缀，"
                            "导致在本表内自引用、绕成循环；② 单元格条件/范围引用错位，最后引用到本格"
                            "自己形成循环。请用 run_python 打开 OUTPUT_FILE 定位并改对这些公式"
                            "（补全跨表前缀、核对引用的行列），改完存回 OUTPUT_FILE，再调用 finish_task。"
                        ),
                    }
                )
                append_run_log_event(
                    task_paths,
                    event="agent_recalc_gate",
                    status="warning",
                    details={"attempt": recalc_attempts, "error_cells": error_cells[:20]},
                )
            else:
                last_error = "真算仍发现公式报错：" + describe_error_cells(gate.get("error_cells", []))
                break
        if completed:
            break
        if built and not finish_requested:
            messages.append(
                {
                    "role": "user",
                    "content": "文件已生成。请用 validate_workbook 检查；若通过再调用 finish_task。",
                }
            )

    produced_file = _claim_produced_workbook(task_paths)
    if produced_file is not None and (built or tool_calls > 0):
        report = validate_workbook(produced_file)
        # 末尾以真算为权威终检；gate 已算过同一文件就复用，省一次 LibreOffice 启动。
        final_recalc = last_recalc if last_recalc is not None else recalc_workbook(produced_file)
        if report.get("status") in {"pass", "warn"} and final_recalc.get("ok", True):
            notices = ["已自动选择本地工具生成；文件仍经过确定性校验。"]
            if final_recalc.get("available"):
                notices.append("已用 LibreOffice 真算复核，公式无 #VALUE!/循环引用等报错。")
            result = AgentResult(
                success=True,
                output_file=str(produced_file),
                message="已通过本地工具生成工作簿。",
                error=None,
                steps=step,
                tool_calls=tool_calls,
                blueprint_file=str(blueprint_path) if blueprint_path.exists() else None,
                notices=notices,
            )
            append_run_log_event(
                task_paths,
                event="agent_orchestrator_completed",
                status="success",
                details=result.to_dict(),
            )
            return result
        if not final_recalc.get("ok", True):
            last_error = last_error or (
                "真算仍发现公式报错：" + describe_error_cells(final_recalc.get("error_cells", []))
            )

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
        "你是本地 Excel 表格智能体。必须用工具完成，不能把 Markdown 当作结果。\n"
        "【首选做法】用 run_python 写 Python（pandas/openpyxl 等任意库，可联网）直接构建或修改工作簿"
        "——想要什么写什么：多工作表、两级表头、公式、条件格式、排序、小计/总计、各类图表都可以。\n"
        "【一段写完】每次 run_python 都是全新独立运行，互不共享变量与 import；需要多步处理时，请在同一段代码里"
        "一次写完（读数据→计算→生成→保存到 OUTPUT_FILE），不要指望沿用上一次 run_python 里定义的变量。\n"
        "【就地编辑铁律】如果用户是要在已有文件上改/填（例如“把A表某列填到B表、其它别动”“给这个表加一列”）："
        "必须 load_workbook(已有文件) 在原工作簿上做最小改动，原样保留未提到的工作表、行列、顺序、样式、公式，"
        "只动用户明确要改的部分，改完存回 OUTPUT_FILE；严禁新建空白工作簿重造——那会丢掉用户原有的排版和数据。\n"
        "硬性约定：\n"
        "1) 上传数据/原文件路径在 INPUT_FILES 列表、模板在 TEMPLATE_FILES 列表（都是绝对路径），直接读取。\n"
        "2) 最终工作簿必须保存到已注入的变量 OUTPUT_FILE（直接 wb.save(OUTPUT_FILE)），不要改用别的文件名。\n"
        "3) 新建的计算列写 Excel 公式（=SUM/=AVERAGE/=IF…），不要把算好的结果写死；注意 IFERROR、空值、除零。\n"
        "4) 图表不要空白：图表引用的单元格要放真实数值；若引用了公式单元格，就 wb.calculation.fullCalcOnLoad=True。\n"
        "5) 写完用 inspect_workbook(OUTPUT_FILE) 自检（工作表、行数、图表数）；有问题就改代码重跑；满足后调用 finish_task。\n"
        "【交付前自检清单】finish 前逐项检查并修好：\n"
        "  · 公式要算得出值、不报错——计算列用 IFERROR 包裹，注意文本/数字类型、除零、空值，避免 #VALUE!/#DIV/0!/#REF!；\n"
        "  · 用户要图表时确实建了图，且图引用的是有真实数值的单元格、坐标轴不为空；\n"
        "  · 数据明细页冻结表头行（ws.freeze_panes）并给表头加自动筛选（ws.auto_filter.ref）；\n"
        "  · 百分比/比率列设百分比格式、金额列设千分位，列宽适配内容、不要过窄。\n"
        "【真算关卡】finish_task 后系统会用 LibreOffice 真算一遍；若有单元格算出 #VALUE!/#DIV/0!/#REF! 或"
        "循环引用，会把具体单元格退回让你改。最常见两类错先避开：① 跨表统计漏写『工作表名!』前缀，在本表里"
        "算成自引用/循环（在统计表写 =AVERAGE(C2:C13) 多半应是 =AVERAGE(成绩表!C2:C13)）；② 单元格条件/范围"
        "引用错位、最后引用到本格自己形成循环。\n"
        "图表(openpyxl)：BarChart(type='col'竖柱/'bar'横条)、LineChart(趋势)、AreaChart(面积)、PieChart(占比)、"
        "DoughnutChart(环形)、RadarChart(多维)、ScatterChart(相关性)、组合用 bar += line(双轴)。"
        "按用户用语选型：占比/构成→饼图，趋势/走势→折线，对比/排名→柱状，多维→雷达，相关性→散点，双指标→组合。"
        "用户如果点名了具体图表类型（例如明确说'雷达图''散点图''环形图'），必须严格用该类型，不得替换成柱状图等其它类型。\n"
        "其它工具(按需)：read_table_summary 先看数据列；fill_template 把数据精确填进上传模板(需要导入原系统时首选)；"
        "build_workbook 提交结构化方案生成常规单表(简单时可用)；validate_workbook 校验。\n"
        "简单需求一步写完即可，别过度设计；复杂需求分步：读数据 → 写代码生成 → 自检 → 修正 → 完成。"
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


def _claim_produced_workbook(task_paths: TaskPaths) -> Path | None:
    """认领智能体实际生成的工作簿。

    优先用流水线期望的 output_file；若智能体把文件存成了别的名字（仍在 output
    目录内），就认领该目录里最新的 .xlsx 并规整为 output_file——避免“明明做对了
    却因为文件名不符被当成失败、再被兜底重造覆盖”。
    """
    expected = task_paths.output_file
    if expected.exists():
        return expected
    output_dir = task_paths.output_dir
    if not output_dir.exists():
        return None
    candidates = [item for item in output_dir.glob("*.xlsx") if item.is_file()]
    if not candidates:
        return None
    newest = max(candidates, key=lambda item: item.stat().st_mtime)
    try:
        shutil.copyfile(newest, expected)
        return expected
    except OSError:
        return newest


def _workbook_has_charts(blueprint: dict[str, Any]) -> bool:
    if blueprint.get("charts"):
        return True
    return any(bool(sheet.get("charts")) for sheet in blueprint.get("sheets", []) if isinstance(sheet, dict))


def _progress(progress: ProgressCallback | None, stage: str, message: str) -> None:
    if progress:
        progress(stage, message)


def _is_truncation(error: str | None) -> bool:
    text = str(error or "")
    return "截断" in text or "length" in text.lower()


def _matched_skills(task_spec: TaskSpec) -> list[dict[str, str]]:
    try:
        from skills.registry import match_skills

        return [
            {"name": item.name, "content": item.content}
            for item in match_skills(task_spec)[:2]
        ]
    except Exception:
        return []
