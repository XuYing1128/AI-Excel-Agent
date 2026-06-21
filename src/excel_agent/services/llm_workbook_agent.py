"""Model-driven workbook agent using local deterministic spreadsheet tools."""

from __future__ import annotations

import json
import re
from dataclasses import asdict, dataclass
from copy import deepcopy
from pathlib import Path
from typing import Any, Callable

from ..api_settings import ApiSettings
from ..rich_workbook_builder import (
    build_rich_workbook,
    inspect_rich_workbook,
    normalize_blueprint,
)
from ..task_paths import TaskPaths, append_run_log_event
from ..task_spec import TaskSpec
from ..validators import validate_workbook
from .custom_api_service import (
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
    agent_settings.timeout_seconds = max(agent_settings.timeout_seconds, 120)
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
                        normalize_blueprint(payload.get("blueprint", payload)),
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
                candidate = normalize_blueprint(
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
            f"大模型共调用本地表格工具 {tool_calls} 次。",
            "具体单元格、公式、合并、样式和图表仍由本地确定性代码写入。",
        ],
    )
    append_run_log_event(
        task_paths,
        event="llm_agent_completed",
        status="success",
        details=result.to_dict(),
    )
    return result


def _execute_build_tool(
    raw_blueprint: Any,
    task_spec: TaskSpec,
    task_paths: TaskPaths,
    blueprint_path: Path,
    progress: ProgressCallback | None,
) -> dict[str, Any]:
    try:
        blueprint = normalize_blueprint(raw_blueprint)
        if task_spec.include_charts and not blueprint.get("charts"):
            return {
                "built": False,
                "error": (
                    "用户最终确认要求生成图表，但方案中 charts 为空。"
                    "请补充与表格数据匹配的图表后再次调用 build_workbook。"
                ),
                "requirement_feedback": ["include_charts=true 是强制需求。"],
            }
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
        "不要增加无关工作表。工作簿实际写入由本地工具完成。"
    )


def _safe_task_spec(task_spec: TaskSpec) -> dict[str, Any]:
    data = task_spec.to_dict()
    data["input_files"] = [Path(item).name for item in task_spec.input_files]
    return data


def _progress(callback: ProgressCallback | None, stage: str, message: str) -> None:
    if callback:
        callback(stage, message)


def _blueprint_requirement_issues(
    task_spec: TaskSpec,
    blueprint: dict[str, Any],
) -> list[str]:
    prompt = task_spec.user_goal
    issues: list[str] = []
    columns = list(blueprint.get("columns", []))
    formulas = {
        str(item.get("label", "")): str(item.get("formula", ""))
        for item in columns
    }
    if "两级表头" in prompt and not blueprint.get("header_groups"):
        issues.append("原需求要求两级表头，但 blueprint 没有 header_groups。")
    if "小计" in prompt and not blueprint.get("group_subtotals"):
        issues.append("原需求要求分组小计，但 blueprint 没有 group_subtotals。")
    if "总计" in prompt and not blueprint.get("grand_total"):
        issues.append("原需求要求总计行，但 blueprint 没有 grand_total。")
    if "排序" in prompt and not blueprint.get("sort"):
        issues.append("原需求要求排序，但 blueprint 没有 sort。")
    if any(word in prompt for word in ("条件格式", "圆点", "警告符号")) and not blueprint.get(
        "conditional_formats"
    ):
        issues.append("原需求要求条件格式或特殊符号，但 blueprint 没有 conditional_formats。")
    if "四舍五入" in prompt:
        average_formulas = [
            formula
            for label, formula in formulas.items()
            if "平均" in label and formula
        ]
        if average_formulas and not all("ROUND(" in formula.upper() for formula in average_formulas):
            issues.append("季度平均要求四舍五入为整数，但平均公式未全部使用 ROUND。")
    if (
        any(phrase in prompt for phrase in ("区域和省份合并", "区域与省份合并"))
        and len((blueprint.get("group_subtotals") or {}).get("merge_label_keys", [])) < 2
    ):
        issues.append("区域小计要求合并区域与省份两列，但 merge_label_keys 不足两列。")
    if (
        "区域、省份、销售员三列合并" in prompt
        and len((blueprint.get("grand_total") or {}).get("merge_label_keys", [])) < 3
    ):
        issues.append("总计行要求合并区域、省份、销售员三列，但 merge_label_keys 不足三列。")
    expected_rows = _count_inline_data_rows(prompt)
    if expected_rows and len(blueprint.get("records", [])) != expected_rows:
        issues.append(
            f"原需求包含 {expected_rows} 条基础数据，但 blueprint records 为 "
            f"{len(blueprint.get('records', []))} 条。"
        )
    for keyword in ("季度总", "季度平均", "业绩等级"):
        matching = [
            formula
            for label, formula in formulas.items()
            if keyword in label
        ]
        if keyword in prompt and (not matching or not all(matching)):
            issues.append(f"原需求要求“{keyword}”使用公式，但对应列未设置 formula。")
    group_config = blueprint.get("group_subtotals") or {}
    total_config = blueprint.get("grand_total") or {}
    if "小计" in prompt and "{group}" not in str(group_config.get("label_template", "")):
        issues.append("区域小计标签必须包含区域名称，请使用 label_template:'{group}小计'。")
    if any(phrase in prompt for phrase in ("销售员为--", "销售员显示“--”")):
        if (group_config.get("value_map") or {}).get("salesperson") != "--":
            issues.append("区域小计的销售员列必须通过 value_map 写入“--”。")
    if any(phrase in prompt for phrase in ("业绩等级为--", "业绩等级显示“--”")):
        if (group_config.get("value_map") or {}).get("grade") != "--":
            issues.append("区域小计的业绩等级列必须通过 value_map 写入“--”。")
    if "区域所有销售员季度总额" in prompt:
        if (group_config.get("average_from") or {}).get("quarterly_avg") != "quarterly_total":
            issues.append(
                "区域小计的季度平均列必须使用 "
                "average_from:{'quarterly_avg':'quarterly_total'}。"
            )
    if "所有销售员季度总额平均值" in prompt:
        if (total_config.get("average_from") or {}).get("quarterly_avg") != "quarterly_total":
            issues.append(
                "总计行的季度平均列必须使用 "
                "average_from:{'quarterly_avg':'quarterly_total'}。"
            )
    quarterly_column = next(
        (item for item in columns if "季度总" in str(item.get("label", ""))),
        {},
    )
    if "绿色圆点" in prompt and "🟢" not in str(quarterly_column.get("number_format", "")):
        issues.append("季度总额的 number_format 缺少绿色圆点显示规则。")
    if "红色圆点" in prompt and "🔴" not in str(quarterly_column.get("number_format", "")):
        issues.append("季度总额的 number_format 缺少红色圆点显示规则。")
    grade_formula = next(
        (
            str(item.get("formula", ""))
            for item in columns
            if "业绩等级" in str(item.get("label", ""))
        ),
        "",
    )
    if "警告符号" in prompt and "⚠" not in grade_formula:
        issues.append("业绩等级公式必须在“待改进”前写入黄色警告符号“⚠”。")
    canonical_sales_labels = [
        "区域",
        "省份",
        "销售员",
        "1月",
        "2月",
        "3月",
        "季度总",
        "季度平均",
        "业绩等级",
    ]
    if all(label in prompt for label in canonical_sales_labels):
        actual_positions = []
        for expected in canonical_sales_labels:
            position = next(
                (
                    index
                    for index, item in enumerate(columns)
                    if expected in str(item.get("label", ""))
                ),
                None,
            )
            actual_positions.append(position)
        if None not in actual_positions and actual_positions != sorted(actual_positions):
            issues.append(
                "列顺序必须为：区域、省份、销售员、1月、2月、3月、"
                "季度总额、季度平均、业绩等级。"
            )
    requested_labels = _extract_requested_column_order(prompt)
    actual_labels = [str(item.get("label", "")).strip() for item in columns]
    if requested_labels and not _column_order_matches(requested_labels, actual_labels):
        issues.append(
            "生成列必须严格对应用户指定的列顺序，不能重复、遗漏或增加无关列。"
            f"要求：{'、'.join(requested_labels)}；当前：{'、'.join(actual_labels)}。"
        )
    normalized_labels = [_normalize_label(item) for item in actual_labels]
    duplicates = sorted(
        {
            label
            for label in normalized_labels
            if label and normalized_labels.count(label) > 1
        }
    )
    if duplicates:
        issues.append("存在重复业务列：" + "、".join(duplicates) + "。")
    return issues


def _count_inline_data_rows(prompt: str) -> int:
    count = 0
    for line in str(prompt).splitlines():
        parts = [item.strip() for item in line.split(",")]
        if len(parts) < 5:
            continue
        numeric = 0
        for value in parts[-3:]:
            try:
                float(value)
                numeric += 1
            except ValueError:
                pass
        if numeric == 3:
            count += 1
    return count


def _extract_requested_column_order(prompt: str) -> list[str]:
    match = re.search(
        r"(?:列顺序(?:必须)?(?:严格)?(?:固定)?为|列顺序|列为|表头为)"
        r"\s*[：:]\s*([^\r\n。]+)",
        str(prompt),
    )
    if not match:
        return []
    values = [
        item.strip(" `\"“”'")
        for item in re.split(r"[、，,；;|]", match.group(1))
    ]
    return [item for item in values if item][:30]


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
                            "required": ["title", "sheet_name", "columns", "records"],
                            "properties": {
                                "title": {"type": "string"},
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
                                                "enum": ["bar", "column", "line", "pie"],
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
