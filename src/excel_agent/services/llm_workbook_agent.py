"""Model-driven workbook agent using local deterministic spreadsheet tools."""

from __future__ import annotations

import json
import re
from dataclasses import asdict, dataclass
from copy import deepcopy
from pathlib import Path
from typing import Any, Callable

from ..api_settings import ApiSettings
from ..io_utils import read_table
from ..rich_workbook_builder import (
    build_default_chart_spec,
    build_rich_workbook,
    inspect_rich_workbook,
    normalize_blueprint,
    normalize_workbook_blueprint,
)
from ..task_paths import TaskPaths, append_run_log_event
from ..task_spec import TaskSpec
from ..validators import validate_workbook
from .custom_api_service import (
    chat_completion,
    chat_completion_with_tools,
    parse_json_object,
)


ProgressCallback = Callable[[str, str], None]


@dataclass
class AgentGenerationResult:
    success: bool
    output_file: str | None
    message: str
    error: str | None
    tool_calls: int
    rounds: int
    blueprint_file: str | None
    model_summary: str
    validation_status: str | None
    notices: list[str]

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


def generate_with_llm_agent(
    task_spec: TaskSpec,
    task_paths: TaskPaths,
    settings: ApiSettings,
    *,
    progress: ProgressCallback | None = None,
    max_rounds: int = 6,
) -> AgentGenerationResult:
    """Let the model plan and call a local workbook-building tool."""

    agent_settings = ApiSettings.from_dict(settings.to_dict(include_secret=True))
    agent_settings.timeout_seconds = max(agent_settings.timeout_seconds, 240)
    _progress(progress, "model", "正在让大模型逐项拆解表格结构和计算要求……")
    messages = [
        {"role": "system", "content": _system_prompt()},
        {
            "role": "user",
            "content": json.dumps(
                {
                    "original_request": task_spec.user_goal,
                    "confirmed_task_spec": _safe_task_spec(task_spec),
                    "important_rule": (
                        "最终确认项优先于最初文字。include_charts=true 表示用户明确要求图表，"
                        "不得在审查中视为过度设计或建议删除。"
                        "若提供模板：reference 仅参考视觉；flexible 尽量沿用形式但需求优先；"
                        "strict 必须与模板字段和顺序兼容。模板样例数据默认不得作为业务数据。"
                    ),
                },
                ensure_ascii=False,
            ),
        },
    ]
    tools = _tool_definitions()
    tool_calls = 0
    last_error = ""
    last_summary = ""
    built = False
    stalled = 0
    current_blueprint: dict[str, Any] | None = None
    blueprint_path = task_paths.task_dir / "workbook_blueprint.json"

    append_run_log_event(
        task_paths,
        event="llm_agent_started",
        status="running",
        details={
            "provider": settings.provider_name,
            "model": settings.model,
            "include_charts": task_spec.include_charts,
        },
    )

    for round_index in range(1, max_rounds + 1):
        response = chat_completion_with_tools(
            agent_settings,
            messages=messages,
            tools=tools,
            # Some OpenAI-compatible reasoning models reject "required".
            # Keep the portable "auto" mode and correct non-tool replies in-band.
            tool_choice="auto",
            temperature=0.05,
            # Reasoning models may consume a large part of the completion budget
            # before emitting the JSON tool arguments. Keep enough room for a
            # complete blueprint with records, formulas and chart definitions.
            max_tokens=16000,
        )
        if not response.success:
            last_error = response.error or "模型调用失败。"
            break

        assistant_message = dict(response.message or {"role": "assistant"})
        assistant_message.setdefault("role", "assistant")
        messages.append(assistant_message)

        if not response.tool_calls:
            if not built:
                try:
                    payload = parse_json_object(response.content)
                    blueprint = _merge_blueprint_revision(
                        current_blueprint,
                        _normalize_agent_blueprint(
                            payload.get("blueprint", payload)
                        ),
                    )
                    current_blueprint = blueprint
                    tool_result = _execute_build_tool(
                        blueprint,
                        task_spec,
                        task_paths,
                        blueprint_path,
                        progress,
                    )
                    tool_calls += 1
                    built = tool_result["built"]
                    last_error = tool_result.get("error", "")
                    last_summary = response.content
                except (ValueError, TypeError) as exc:
                    last_error = f"模型未调用工具且返回内容无法作为方案：{exc}"
                    stalled += 1
                    if stalled >= 2:
                        # The endpoint/model cannot return a usable plan via the
                        # tool loop. Stop early so the caller can fall back to
                        # JSON-mode or local generation instead of burning rounds.
                        break
                    messages.append(
                        {
                            "role": "user",
                            "content": (
                                "你刚才没有调用工具，因此任务尚未开始。"
                                "不要再输出解释、计划或 Markdown；现在必须调用 "
                                "build_workbook，并在 blueprint 中完整提交原需求的"
                                "列、基础数据、公式、两级表头、小计、总计、排序、"
                                "条件格式和图表。"
                            ),
                        }
                    )
                    continue
                if built:
                    messages.append(
                        {
                            "role": "user",
                            "content": (
                                "本地工具已按你返回的 JSON 方案生成文件。"
                                "请继续逐项核对原需求；如有缺失，调用 build_workbook "
                                "提交完整修订版；全部满足后调用 complete_task。"
                            ),
                        }
                    )
                    continue
                messages.append(
                    {
                        "role": "user",
                        "content": (
                            f"方案未通过本地工具检查：{last_error or '存在需求缺项'}。"
                            "请调用 build_workbook 提交完整修订版。"
                        ),
                    }
                )
                continue
            last_summary = response.content
            messages.append(
                {
                    "role": "user",
                    "content": (
                        "文件已经生成，但你尚未通过 complete_task 明确完成。"
                        "若所有要求已满足，请调用 complete_task；否则调用 "
                        "build_workbook 提交完整修订方案。"
                    ),
                }
            )
            continue

        completed = False
        for call in response.tool_calls:
            tool_calls += 1
            if call.name == "build_workbook":
                _progress(
                    progress,
                    "build",
                    f"大模型已提交第 {round_index} 版方案，正在调用本地 Excel 工具生成……",
                )
                candidate = _normalize_agent_blueprint(
                    call.arguments.get("blueprint", call.arguments)
                )
                candidate = _merge_blueprint_revision(current_blueprint, candidate)
                current_blueprint = candidate
                tool_result = _execute_build_tool(
                    candidate,
                    task_spec,
                    task_paths,
                    blueprint_path,
                    progress,
                )
                built = tool_result["built"]
                last_error = tool_result.get("error", "")
                messages.append(
                    {
                        "role": "tool",
                        "tool_call_id": call.id,
                        "name": call.name,
                        "content": json.dumps(tool_result, ensure_ascii=False),
                    }
                )
            elif call.name == "complete_task":
                last_summary = str(call.arguments.get("summary") or response.content)
                completed = built
                messages.append(
                    {
                        "role": "tool",
                        "tool_call_id": call.id,
                        "name": call.name,
                        "content": json.dumps(
                            {
                                "accepted": completed,
                                "message": (
                                    "任务已完成。"
                                    if completed
                                    else "尚未成功调用 build_workbook。"
                                ),
                            },
                            ensure_ascii=False,
                        ),
                    }
                )
        if completed:
            break

    if not built or not task_paths.output_file.exists():
        result = AgentGenerationResult(
            success=False,
            output_file=None,
            message="大模型表格代理未能生成可用文件。",
            error=last_error or "模型没有成功调用本地工作簿工具。",
            tool_calls=tool_calls,
            rounds=min(max_rounds, max(1, round_index)),
            blueprint_file=str(blueprint_path) if blueprint_path.exists() else None,
            model_summary=last_summary,
            validation_status=None,
            notices=[],
        )
        append_run_log_event(
            task_paths,
            event="llm_agent_failed",
            status="error",
            details=result.to_dict(),
        )
        return result

    final_report = validate_workbook(task_paths.output_file)
    _progress(progress, "done", "大模型工具调用完成，正在整理最终结果……")
    result = AgentGenerationResult(
        success=True,
        output_file=str(task_paths.output_file),
        message="大模型已通过本地 Excel 工具生成工作簿。",
        error=None,
        tool_calls=tool_calls,
        rounds=min(max_rounds, max(1, round_index)),
        blueprint_file=str(blueprint_path),
        model_summary=last_summary,
        validation_status=final_report.get("status"),
        notices=[
            "大模型设计了整体方案，表格里的具体内容、公式和图表由本地稳定生成。",
        ],
    )
    append_run_log_event(
        task_paths,
        event="llm_agent_completed",
        status="success",
        details=result.to_dict(),
    )
    return result


def generate_blueprint_via_json(
    task_spec: TaskSpec,
    task_paths: TaskPaths,
    settings: ApiSettings,
    *,
    progress: ProgressCallback | None = None,
) -> AgentGenerationResult:
    """Fallback path: ask the model for one pure-JSON workbook plan.

    Many OpenAI-compatible endpoints (and reasoning models) return JSON far more
    reliably than function/tool calls. This keeps the model in charge of the
    design while avoiding the tool-calling protocol entirely.
    """

    agent_settings = ApiSettings.from_dict(settings.to_dict(include_secret=True))
    agent_settings.timeout_seconds = max(agent_settings.timeout_seconds, 240)
    _progress(progress, "model", "正在请大模型直接给出表格设计方案……")
    system_prompt = (
        _system_prompt()
        + "\n本次不要使用工具调用。只返回一个 JSON 对象，顶层就是工作簿方案，"
        "单表可包含 title、sheet_name、columns、records；多表任务必须使用 "
        "title、sheets，其中每个 sheet 包含 sheet_name、title、columns、records。"
        "工作表可选 header_groups、"
        "group_subtotals、grand_total、sort、conditional_formats、charts、notes。"
        "不要输出 Markdown、解释或代码块标记。"
    )
    user_payload = json.dumps(
        {
            "original_request": task_spec.user_goal,
            "confirmed_task_spec": _safe_task_spec(task_spec),
        },
        ensure_ascii=False,
    )
    blueprint_path = task_paths.task_dir / "workbook_blueprint.json"
    last_error = ""
    append_run_log_event(
        task_paths,
        event="llm_json_mode_started",
        status="running",
        details={"provider": settings.provider_name, "model": settings.model},
    )
    for attempt in range(2):
        response = chat_completion(
            agent_settings,
            system_prompt=system_prompt,
            user_prompt=(
                user_payload
                if attempt == 0
                else user_payload + "\n请只返回紧凑的 JSON，不要任何额外文字。"
            ),
            temperature=0.05,
            max_tokens=16000,
            json_mode=True,
        )
        if not response.success:
            last_error = response.error or "模型调用失败。"
            continue
        try:
            payload = parse_json_object(response.content)
        except (ValueError, TypeError) as exc:
            last_error = f"模型未返回可解析的 JSON 方案：{exc}"
            continue
        raw_blueprint = payload.get("blueprint", payload)
        result = _execute_build_tool(
            raw_blueprint, task_spec, task_paths, blueprint_path, progress
        )
        if result.get("built") and task_paths.output_file.exists():
            final_report = validate_workbook(task_paths.output_file)
            success_result = AgentGenerationResult(
                success=True,
                output_file=str(task_paths.output_file),
                message="大模型已直接给出表格方案，由本地工具生成。",
                error=None,
                tool_calls=1,
                rounds=attempt + 1,
                blueprint_file=str(blueprint_path),
                model_summary=str(payload.get("summary") or ""),
                validation_status=final_report.get("status"),
                notices=["大模型给出整体方案，具体单元格、公式和图表由本地稳定生成。"],
            )
            append_run_log_event(
                task_paths,
                event="llm_json_mode_completed",
                status="success",
                details=success_result.to_dict(),
            )
            return success_result
        last_error = (
            result.get("error")
            or "；".join(result.get("errors", []))
            or "方案不完整。"
        )
    failure = AgentGenerationResult(
        success=False,
        output_file=None,
        message="大模型直接给方案这一步未能成功。",
        error=last_error or "模型没有返回可用方案。",
        tool_calls=0,
        rounds=2,
        blueprint_file=str(blueprint_path) if blueprint_path.exists() else None,
        model_summary="",
        validation_status=None,
        notices=[],
    )
    append_run_log_event(
        task_paths,
        event="llm_json_mode_failed",
        status="error",
        details=failure.to_dict(),
    )
    return failure


def _execute_build_tool(
    raw_blueprint: Any,
    task_spec: TaskSpec,
    task_paths: TaskPaths,
    blueprint_path: Path,
    progress: ProgressCallback | None,
) -> dict[str, Any]:
    try:
        blueprint = _normalize_agent_blueprint(raw_blueprint)
        blueprint = _attach_input_records(blueprint, task_spec)
        # When the user asked for a chart but the model omitted one, the local
        # builder adds a sensible default instead of bouncing the model. If the
        # user picked a preferred chart type, honour it for that default.
        preferred_types = task_spec.options.get("chart_types") or []
        if (
            task_spec.include_charts
            and preferred_types
            and isinstance(blueprint, dict)
            and blueprint.get("columns")
            and not blueprint.get("charts")
        ):
            spec = build_default_chart_spec(blueprint["columns"], str(preferred_types[0]))
            if spec:
                blueprint["charts"] = [spec]
        blueprint_path.write_text(
            json.dumps(blueprint, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        task_spec.options["agent_blueprint"] = blueprint
        build_rich_workbook(
            blueprint,
            task_paths.output_file,
            require_charts=task_spec.include_charts,
        )
        _progress(progress, "validate", "本地工具已生成文件，正在检查公式、结构、图表和条件格式……")
        validation = validate_workbook(task_paths.output_file)
        inspection = inspect_rich_workbook(task_paths.output_file)
        chart_count = sum(
            int(sheet.get("chart_count", 0))
            for sheet in inspection.get("sheets", [])
        )
        issues = _blueprint_requirement_issues(task_spec, blueprint)
        if task_spec.include_charts and chart_count == 0:
            issues.append("用户要求图表，但生成文件没有图表。")
        return {
            "built": not validation.get("errors") and not issues,
            "validation_status": "fail" if issues else validation.get("status"),
            "errors": [*validation.get("errors", []), *issues],
            "warnings": validation.get("warnings", []),
            "inspection": inspection,
            "confirmed_options": {
                "include_charts": task_spec.include_charts,
                "include_summary": task_spec.include_summary,
                "output_name": task_spec.output_name,
            },
            "instruction": (
                "逐项对照 original_request 和 confirmed_task_spec。"
                "若有缺失，请修订完整 blueprint 后再次调用 build_workbook；"
                "全部满足后调用 complete_task。"
            ),
        }
    except Exception as exc:
        return {
            "built": False,
            "error": f"{type(exc).__name__}: {exc}",
            "instruction": "修正方案后再次调用 build_workbook。",
        }


def _system_prompt() -> str:
    return (
        "你是高级 Excel 工作簿设计代理。你必须通过工具调用完成任务，而不是只返回文字。"
        "先逐条读取 original_request，再读取 confirmed_task_spec；最终确认项优先级最高。"
        "你必须把所有明确要求转成一个完整 blueprint，然后调用 build_workbook。"
        "收到工具返回的结构、公式、图表和校验结果后，逐项核对原需求；发现缺失就提交完整修订版 blueprint "
        "再次调用 build_workbook，全部满足后调用 complete_task。"
        "禁止用无关标准模板代替需求。禁止遗漏基础数据。所有计算列必须是 formula，不能硬编码计算结果。"
        "若需求文字本身包含制表符、Markdown 或 CSV 数据，这些数据就是正式输入，不能声称未提供数据。"
        "若需求要求参数表、明细表或多个关联表，顶层必须使用 sheets 分别描述每个工作表；"
        "跨表公式直接引用对应工作表，不得把参数常量硬编码到公式。"
        "若 confirmed_task_spec.options.input_data_profile 存在，只根据列名和行数设计结构，"
        "records 可以留空；本地工具会读取真实文件并按列名填充，模型不得编造示例数据。"
        "多级表头使用 header_groups；区域小计使用 group_subtotals；总计使用 grand_total；"
        "排序使用 sort；条件格式使用 conditional_formats；图表使用 charts。"
        "group_subtotals 可使用 group_key、label_template、merge_label_keys、sum_keys、"
        "average_keys、average_from、value_map、blank_keys；grand_total 可使用 label、"
        "merge_label_keys、sum_keys、average_keys、average_from、value_map、blank_keys。"
        "average_from 的格式为 {目标列key:源列key}，例如 "
        "{'quarterly_avg':'quarterly_total'} 会在季度平均列写入源列的 ROUND(AVERAGE(...),0)。"
        "value_map 用于小计或总计的固定文本，例如 {'salesperson':'--','grade':'--'}。"
        "区域小计标签必须用 label_template:'{group}小计'，不能只写“小计”。"
        "sort 的区域顺序示例为 {key:'region',order:['华东','华南']}；"
        "组内按多列合计降序示例为 {keys:['jan','feb','mar'],aggregate:'sum',direction:'desc'}。"
        "条件格式 cell_is 使用 column_key、operator、value、font_bold、font_color、fill_color；"
        "公式条件使用 kind:'formula' 和 formula，其中 {cell} 代表该列首个数据单元格。"
        "如需在数值前显示符号，使用列的 number_format，例如 "
        "[>=400000]\"🟢 \"#,##0;[<150000]\"🔴 \"#,##0;#,##0。"
        "若 include_charts=true，charts 至少一个，这是用户明确需求，不是额外设计。"
        "图表 type 可选：column(竖向柱/对比) bar(横向条) line(趋势) area(面积) "
        "pie(占比) doughnut(环形占比) radar(多维度) scatter(相关性) combo(柱+线双轴)；"
        "请根据用户用语选择最合适的类型：占比/构成→pie，趋势/走势→line，对比/排名→column，"
        "多维能力→radar，相关性→scatter，双指标→combo。"
        "不要增加无关工作表。工作簿实际写入由本地工具完成。"
        "请直接给出方案、尽量减少冗长推理；若需要你生成示例数据（没有上传文件时），"
        "示例行数控制在 8 行以内，避免输出过长被截断。"
    )


def _safe_task_spec(task_spec: TaskSpec) -> dict[str, Any]:
    data = task_spec.to_dict()
    data["input_files"] = [Path(item).name for item in task_spec.input_files]
    data["template_files"] = [Path(item).name for item in task_spec.template_files]
    options = dict(data.get("options") or {})
    options.pop("prepared_template_file", None)
    data["options"] = options
    return data


def _attach_input_records(
    blueprint: dict[str, Any],
    task_spec: TaskSpec,
) -> dict[str, Any]:
    if not task_spec.input_files:
        return blueprint
    sheet_plans = (
        list(blueprint.get("sheets", []))
        if isinstance(blueprint.get("sheets"), list)
        else [blueprint]
    )
    records = []
    available_columns: list[str] = []
    target_sheet: dict[str, Any] | None = None
    for input_path in task_spec.input_files:
        frame = read_table(input_path)
        if frame.empty:
            continue
        available_columns.extend(str(item) for item in frame.columns)
        normalized_sources = {
            _normalize_label(str(column)): str(column)
            for column in frame.columns
        }
        mappings: dict[str, str] = {}
        target_sheet = max(
            sheet_plans,
            key=lambda item: _column_overlap_score(item, normalized_sources),
        )
        input_columns = [
            item for item in target_sheet.get("columns", []) if not item.get("formula")
        ]
        for column in input_columns:
            candidates = (
                _normalize_label(str(column.get("label", ""))),
                _normalize_label(str(column.get("key", ""))),
            )
            source = next(
                (
                    normalized_sources[candidate]
                    for candidate in candidates
                    if candidate in normalized_sources
                ),
                None,
            )
            if source:
                mappings[column["key"]] = source
        if not mappings:
            continue
        for source_record in frame.to_dict("records"):
            records.append(
                {
                    target: _json_safe_value(source_record.get(source))
                    for target, source in mappings.items()
                }
            )
    if not records:
        raise ValueError(
            "模型方案中的输入列与上传数据字段无法对应。"
            f"上传字段：{'、'.join(dict.fromkeys(available_columns))}。"
        )
    updated = deepcopy(blueprint)
    if isinstance(updated.get("sheets"), list):
        target_name = str((target_sheet or {}).get("sheet_name") or "")
        for sheet in updated["sheets"]:
            if str(sheet.get("sheet_name") or "") == target_name:
                sheet["records"] = records
                break
    else:
        updated["records"] = records
    return _normalize_agent_blueprint(updated)


def _column_overlap_score(
    sheet: dict[str, Any],
    normalized_sources: dict[str, str],
) -> int:
    score = 0
    for column in sheet.get("columns", []):
        if column.get("formula"):
            continue
        candidates = {
            _normalize_label(str(column.get("label", ""))),
            _normalize_label(str(column.get("key", ""))),
        }
        if candidates & set(normalized_sources):
            score += 1
    return score


def _json_safe_value(value: Any) -> Any:
    if value is None:
        return ""
    if hasattr(value, "to_pydatetime"):
        return value.to_pydatetime().isoformat(sep=" ")
    if hasattr(value, "isoformat") and not isinstance(value, (str, bytes)):
        try:
            return value.isoformat()
        except (TypeError, ValueError):
            pass
    if hasattr(value, "item"):
        try:
            return value.item()
        except (ValueError, TypeError):
            pass
    try:
        if value != value:
            return ""
    except (TypeError, ValueError):
        pass
    return value


def _progress(callback: ProgressCallback | None, stage: str, message: str) -> None:
    if callback:
        callback(stage, message)


_AGGREGATE_LABEL_HINTS = ("合计", "小计", "总计", "平均")


def _blueprint_requirement_issues(
    task_spec: TaskSpec,
    blueprint: dict[str, Any],
) -> list[str]:
    """General, prompt-agnostic coherence checks.

    These deliberately avoid matching any single example. They only flag things
    that are wrong for *any* table: missing columns, duplicate columns, an
    explicit column order the user spelled out, aggregate columns written as
    constants, and an inline-data row count that does not match.
    """

    prompt = str(task_spec.user_goal or "")
    issues: list[str] = []
    sheets = (
        list(blueprint.get("sheets", []))
        if isinstance(blueprint.get("sheets"), list)
        else [blueprint]
    )
    columns_by_sheet = [
        list(sheet.get("columns", []))
        for sheet in sheets
        if isinstance(sheet, dict)
    ]
    columns = [item for group in columns_by_sheet for item in group]
    if not columns:
        return ["方案没有任何数据列，请补全表头后重新生成。"]

    actual_labels = [str(item.get("label", "")).strip() for item in columns]

    # Duplicate columns are only a problem WITHIN one sheet. The same generic key
    # (e.g. col_1) legitimately appears on different sheets of a multi-sheet
    # workbook, so checking globally would falsely reject valid plans.
    duplicates: set[str] = set()
    for group in columns_by_sheet:
        sheet_labels = [_normalize_label(str(item.get("label", ""))) for item in group]
        for label in sheet_labels:
            if label and sheet_labels.count(label) > 1:
                duplicates.add(label)
    if duplicates:
        issues.append(
            "同一张表里有重复的列，请合并或删除：" + "、".join(sorted(duplicates)) + "。"
        )

    requested_labels = _extract_requested_column_order(prompt)
    matching_sheet = any(
        _column_order_matches(
            requested_labels,
            [str(item.get("label", "")).strip() for item in sheet_columns],
        )
        for sheet_columns in columns_by_sheet
    )
    if requested_labels and not matching_sheet:
        issues.append(
            "请严格按用户给出的列顺序生成，不要遗漏、重复或新增列。"
            f"要求：{'、'.join(requested_labels)}；当前：{'、'.join(actual_labels)}。"
        )

    for item in columns:
        label = str(item.get("label", ""))
        if any(hint in label for hint in _AGGREGATE_LABEL_HINTS) and not str(
            item.get("formula", "")
        ).strip():
            issues.append(
                f"“{label}”看起来是汇总/计算列，必须用公式（formula）实现，不能写死数值。"
            )

    expected_rows = _count_inline_data_rows(prompt)
    if expected_rows and not task_spec.input_files:
        actual_rows = max(
            (len(sheet.get("records", [])) for sheet in sheets),
            default=0,
        )
        if actual_rows < expected_rows:
            issues.append(
                f"需求文字里至少包含 {expected_rows} 行基础数据，但方案最多只有 {actual_rows} 行，请补齐。"
            )
    expected_sheets = [
        str(item)
        for item in task_spec.options.get("content_plan", {}).get(
            "expected_sheet_names", []
        )
        if str(item).strip()
    ]
    if len(expected_sheets) > 1 and len(sheets) < len(expected_sheets):
        issues.append(
            f"需求中识别到 {len(expected_sheets)} 个数据表，但方案只有 {len(sheets)} 个工作表。"
        )
    return issues


def _count_inline_data_rows(prompt: str) -> int:
    from ..inline_table_parser import extract_inline_tables, primary_inline_table

    primary = primary_inline_table(extract_inline_tables(prompt))
    return int(primary.get("row_count") or 0) if primary else 0


_COLUMN_CLAUSE_STOP_WORDS = (
    "公式",
    "计算",
    "排序",
    "降序",
    "升序",
    "汇总",
    "小计",
    "总计",
    "生成",
    "图表",
    "柱状",
    "折线",
    "饼图",
    "并",
    "其中",
    "要求",
)


def _extract_requested_column_order(prompt: str) -> list[str]:
    # Capture only up to the first clause terminator. Users routinely write
    # "列为：A、B、C；C 用公式…", so stopping at 、。；; and newlines (not just 。)
    # avoids swallowing the following instructions as if they were columns.
    match = re.search(
        r"(?:列顺序(?:必须)?(?:严格)?(?:固定)?为|列顺序|列为|表头为)"
        r"\s*[：:]\s*([^。；;\r\n]+)",
        str(prompt),
    )
    if not match:
        return []
    values = [item.strip(" `\"“”'") for item in re.split(r"[、，,|]", match.group(1))]
    cleaned: list[str] = []
    for item in values:
        if not item or len(item) > 14:
            break
        if any(word in item for word in _COLUMN_CLAUSE_STOP_WORDS):
            break
        cleaned.append(item)
    return cleaned[:30]


def _column_order_matches(expected: list[str], actual: list[str]) -> bool:
    if len(expected) != len(actual):
        return False
    return all(
        _normalize_label(wanted) == _normalize_label(found)
        or _normalize_label(wanted) in _normalize_label(found)
        or _normalize_label(found) in _normalize_label(wanted)
        for wanted, found in zip(expected, actual)
    )


def _normalize_label(value: str) -> str:
    return re.sub(r"[\s（）()_\-—:：]", "", str(value)).lower()


def _merge_blueprint_revision(
    previous: dict[str, Any] | None,
    revision: dict[str, Any],
) -> dict[str, Any]:
    """Preserve previously satisfied capabilities when a model sends a partial revision."""

    if previous is None:
        return revision
    if revision.get("sheets"):
        return normalize_workbook_blueprint(revision)
    if previous.get("sheets"):
        return previous
    merged = deepcopy(previous)
    for field in ("title", "sheet_name"):
        if revision.get(field):
            merged[field] = revision[field]
    if revision.get("columns"):
        old_columns = list(merged.get("columns", []))
        old_by_key = {item["key"]: item for item in old_columns}
        old_by_label = {
            _normalize_label(str(item.get("label", ""))): item
            for item in old_columns
            if item.get("label")
        }
        combined = []
        key_remap: dict[str, str] = {}
        for item in revision["columns"]:
            old_item = old_by_key.get(item["key"]) or old_by_label.get(
                _normalize_label(str(item.get("label", "")))
            )
            base = deepcopy(old_item or {})
            for key, value in item.items():
                if value not in ("", None, []):
                    base[key] = value
            if old_item and old_item["key"] != base["key"]:
                key_remap[old_item["key"]] = base["key"]
            combined.append(base)
        merged["columns"] = combined
        if key_remap and not revision.get("records"):
            remapped_records = []
            for record in merged.get("records", []):
                remapped = dict(record)
                for old_key, new_key in key_remap.items():
                    if old_key in remapped and new_key not in remapped:
                        remapped[new_key] = remapped.pop(old_key)
                remapped_records.append(remapped)
            merged["records"] = remapped_records
    if revision.get("records"):
        merged["records"] = revision["records"]
    for field in (
        "header_groups",
        "sort",
        "conditional_formats",
        "charts",
        "notes",
    ):
        if revision.get(field):
            merged[field] = revision[field]
    for field in ("group_subtotals", "grand_total"):
        if revision.get(field):
            base = deepcopy(merged.get(field) or {})
            base.update(revision[field])
            for list_field in (
                "merge_label_keys",
                "sum_keys",
                "average_keys",
                "blank_keys",
            ):
                if not revision[field].get(list_field) and base.get(list_field):
                    continue
                if revision[field].get(list_field):
                    base[list_field] = revision[field][list_field]
            merged[field] = base
    return normalize_blueprint(merged)


def _normalize_agent_blueprint(value: Any) -> dict[str, Any]:
    if isinstance(value, dict) and isinstance(value.get("sheets"), list):
        return normalize_workbook_blueprint(value)
    return normalize_blueprint(value)


def _tool_definitions() -> list[dict[str, Any]]:
    return [
        {
            "type": "function",
            "function": {
                "name": "build_workbook",
                "description": (
                    "使用完整工作簿 blueprint 调用本地 Excel 生成器。"
                    "每次修订都必须重新提交完整 blueprint。"
                ),
                "parameters": {
                    "type": "object",
                    "required": ["blueprint"],
                    "properties": {
                        "blueprint": {
                            "type": "object",
                            "required": ["title"],
                            "properties": {
                                "title": {"type": "string"},
                                "sheets": {
                                    "type": "array",
                                    "description": (
                                        "多工作表任务使用。每项沿用单表方案字段："
                                        "sheet_name、title、columns、records、header_groups、"
                                        "sort、group_subtotals、grand_total、conditional_formats、charts。"
                                    ),
                                    "items": {"type": "object"},
                                },
                                "sheet_name": {"type": "string"},
                                "columns": {
                                    "type": "array",
                                    "items": {
                                        "type": "object",
                                        "required": ["key", "label", "type"],
                                        "properties": {
                                            "key": {"type": "string"},
                                            "label": {"type": "string"},
                                            "type": {
                                                "type": "string",
                                                "enum": [
                                                    "text",
                                                    "number",
                                                    "integer",
                                                    "money",
                                                    "percentage",
                                                    "date",
                                                ],
                                            },
                                            "width": {"type": "number"},
                                            "formula": {
                                                "type": "string",
                                                "description": (
                                                    "Excel 公式模板，用 {column_key} 引用同一行列，"
                                                    "例如 =SUM({jan},{feb},{mar})。"
                                                ),
                                            },
                                            "number_format": {"type": "string"},
                                        },
                                    },
                                },
                                "header_groups": {
                                    "type": "array",
                                    "items": {
                                        "type": "object",
                                        "required": ["label", "start_key", "end_key"],
                                        "properties": {
                                            "label": {"type": "string"},
                                            "start_key": {"type": "string"},
                                            "end_key": {"type": "string"},
                                        },
                                    },
                                },
                                "records": {
                                    "type": "array",
                                    "items": {"type": "object"},
                                },
                                "sort": {"type": "array", "items": {"type": "object"}},
                                "group_subtotals": {
                                    "type": "object",
                                    "properties": {
                                        "group_key": {"type": "string"},
                                        "label_template": {"type": "string"},
                                        "merge_label_keys": {
                                            "type": "array",
                                            "items": {"type": "string"},
                                        },
                                        "sum_keys": {
                                            "type": "array",
                                            "items": {"type": "string"},
                                        },
                                        "average_keys": {
                                            "type": "array",
                                            "items": {"type": "string"},
                                        },
                                        "average_from": {
                                            "type": "object",
                                            "additionalProperties": {"type": "string"},
                                        },
                                        "value_map": {
                                            "type": "object",
                                            "additionalProperties": {},
                                        },
                                        "blank_keys": {
                                            "type": "array",
                                            "items": {"type": "string"},
                                        },
                                    },
                                },
                                "grand_total": {
                                    "type": "object",
                                    "properties": {
                                        "label": {"type": "string"},
                                        "merge_label_keys": {
                                            "type": "array",
                                            "items": {"type": "string"},
                                        },
                                        "sum_keys": {
                                            "type": "array",
                                            "items": {"type": "string"},
                                        },
                                        "average_keys": {
                                            "type": "array",
                                            "items": {"type": "string"},
                                        },
                                        "average_from": {
                                            "type": "object",
                                            "additionalProperties": {"type": "string"},
                                        },
                                        "value_map": {
                                            "type": "object",
                                            "additionalProperties": {},
                                        },
                                        "blank_keys": {
                                            "type": "array",
                                            "items": {"type": "string"},
                                        },
                                    },
                                },
                                "conditional_formats": {
                                    "type": "array",
                                    "items": {"type": "object"},
                                },
                                "charts": {
                                    "type": "array",
                                    "items": {
                                        "type": "object",
                                        "required": [
                                            "type",
                                            "title",
                                            "category_key",
                                            "value_keys",
                                        ],
                                        "properties": {
                                            "type": {
                                                "type": "string",
                                                "enum": [
                                                    "column",
                                                    "bar",
                                                    "line",
                                                    "area",
                                                    "pie",
                                                    "doughnut",
                                                    "radar",
                                                    "scatter",
                                                    "combo",
                                                ],
                                            },
                                            "title": {"type": "string"},
                                            "category_key": {"type": "string"},
                                            "value_keys": {
                                                "type": "array",
                                                "items": {"type": "string"},
                                            },
                                            "position": {"type": "string"},
                                        },
                                    },
                                },
                                "notes": {
                                    "type": "array",
                                    "items": {"type": "string"},
                                },
                            },
                        }
                    },
                },
            },
        },
        {
            "type": "function",
            "function": {
                "name": "complete_task",
                "description": "仅在 build_workbook 返回结果已满足全部需求后调用。",
                "parameters": {
                    "type": "object",
                    "required": ["summary"],
                    "properties": {"summary": {"type": "string"}},
                },
            },
        },
    ]
