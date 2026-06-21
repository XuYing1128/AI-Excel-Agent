"""Optional model-assisted intent understanding without workbook data."""

from __future__ import annotations

import json
from dataclasses import dataclass

from ..api_settings import ApiSettings
from ..content_plan import merge_model_content_plan, suggest_output_name
from ..intent_classifier import SUPPORTED_TYPES
from ..task_spec import TaskSpecDraft
from .custom_api_service import chat_completion, parse_json_object


@dataclass
class ApiPlanningResult:
    draft: TaskSpecDraft
    used_api: bool
    message: str


def enhance_task_spec_draft(
    draft: TaskSpecDraft,
    *,
    user_prompt: str,
    input_file_names: list[str],
    settings: ApiSettings,
) -> ApiPlanningResult:
    if not settings.configured or not settings.use_for_intent:
        return ApiPlanningResult(draft, False, "使用本地规则完成需求分析。")

    allowed = ", ".join(SUPPORTED_TYPES)
    request = {
        "用户需求": user_prompt,
        "输入文件名": input_file_names,
        "本地规则初判": draft.task_spec.task_type,
        "本地提取方案": draft.task_spec.options.get("content_plan", {}),
        "允许类型": SUPPORTED_TYPES,
    }
    result = chat_completion(
        settings,
        system_prompt=(
            "你是本地表格工具的需求分析助手。你不能读取文件内容、不能指定单元格地址、"
            "不能输出 Excel 公式或代码。你可以把用户明确要求的标题、列名和计算语义整理成结构化方案。"
            "只返回一个 JSON 对象，不要 Markdown。字段必须是："
            "task_type、confidence、goal_summary、clarifying_questions、"
            "include_charts、include_summary、content_plan。"
            "content_plan 字段为对象，可包含 title、layout、columns、formula_rules、summary_rules。"
            "columns 每项只能包含 name、kind、role；kind 只能是 text/number/money/percentage/date/time，"
            "role 只能是 input/formula。formula_rules 只能使用 average、difference、product、ratio、"
            "sum、weather_advice 等计算语义，不得返回公式字符串。summary_rules 只能使用 "
            "averageif、sumif、countif、average、sum、count。"
            "不要增加用户没有要求的图表、汇总页或字段。"
            f"task_type 只能是：{allowed}。clarifying_questions 最多 5 个中文问题。"
        ),
        user_prompt=json.dumps(request, ensure_ascii=False),
        temperature=0.1,
        max_tokens=1800,
    )
    if not result.success:
        return ApiPlanningResult(
            draft,
            False,
            f"自定义模型暂不可用，已自动使用本地规则：{result.error}",
        )
    try:
        payload = parse_json_object(result.content)
        _merge_payload(draft, payload)
    except (ValueError, TypeError) as exc:
        return ApiPlanningResult(
            draft,
            False,
            f"自定义模型返回内容无法采用，已自动使用本地规则：{exc}",
        )
    return ApiPlanningResult(draft, True, "已使用自定义模型辅助理解需求。")


def _merge_payload(draft: TaskSpecDraft, payload: dict) -> None:
    task_type = str(payload.get("task_type", "")).strip()
    if task_type in SUPPORTED_TYPES:
        draft.task_spec.task_type = task_type
    try:
        confidence = float(payload.get("confidence", draft.task_spec.confidence))
    except (TypeError, ValueError):
        confidence = draft.task_spec.confidence
    draft.task_spec.confidence = max(0.0, min(confidence, 1.0))

    summary = str(payload.get("goal_summary", "")).strip()
    if summary:
        draft.task_spec.options["model_goal_summary"] = summary
    if isinstance(payload.get("include_charts"), bool):
        if payload["include_charts"] is False:
            draft.task_spec.include_charts = False
        elif draft.task_spec.options.get("chart_requested_explicitly"):
            draft.task_spec.include_charts = True
    if isinstance(payload.get("include_summary"), bool):
        if payload["include_summary"] is False:
            draft.task_spec.include_summary = False
        elif draft.task_spec.options.get("summary_sheet_requested_explicitly"):
            draft.task_spec.include_summary = True

    content_plan = payload.get("content_plan")
    if isinstance(content_plan, dict):
        local_plan = draft.task_spec.options.get("content_plan", {})
        may_use_custom_structure = bool(local_plan.get("explicit_structure")) or (
            draft.task_spec.task_type == "generic_table"
            and isinstance(content_plan.get("columns"), list)
        )
        if may_use_custom_structure:
            merged_plan = merge_model_content_plan(local_plan, content_plan)
            draft.task_spec.options["content_plan"] = merged_plan
            if merged_plan.get("explicit_structure"):
                draft.task_spec.options["generation_policy"] = "custom_content"
                draft.task_spec.output_name = suggest_output_name(
                    draft.task_spec.user_goal,
                    draft.task_spec.task_type,
                    merged_plan.get("title"),
                )
                if merged_plan.get("layout") == "single_sheet":
                    draft.task_spec.include_instructions_sheet = False
        else:
            model_title = str(content_plan.get("title", "")).strip()
            if model_title:
                draft.task_spec.output_name = suggest_output_name(
                    draft.task_spec.user_goal,
                    draft.task_spec.task_type,
                    model_title,
                )

    questions = payload.get("clarifying_questions", [])
    if isinstance(questions, list) and draft.clarifying_questions:
        model_questions = [str(item).strip() for item in questions if str(item).strip()]
        combined = list(dict.fromkeys([*draft.clarifying_questions, *model_questions]))
        draft.clarifying_questions = combined[:5]
    draft.task_spec.options["intent_source"] = "custom_api"
