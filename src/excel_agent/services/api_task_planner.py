"""Optional model-assisted intent understanding without workbook data."""

from __future__ import annotations

import json
from dataclasses import dataclass

from ..api_settings import ApiSettings
from ..content_plan import merge_model_content_plan, suggest_output_name
from ..intent_classifier import SUPPORTED_TYPES
from ..model_registry import get_role_api_settings
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
    active_settings = (
        settings
        if settings.configured and settings.use_for_intent
        else get_role_api_settings(
            "planner",
            use_for_intent=True,
            use_for_review=False,
            use_for_generation=False,
        )
    )
    if active_settings is None or not active_settings.configured or not active_settings.use_for_intent:
        return ApiPlanningResult(draft, False, "使用本地规则完成需求分析。")

    # Reasoning ("thinking") models spend a large part of the completion budget
    # before emitting the JSON plan, and can be slow. Give the analysis call a
    # longer timeout and enough tokens so it does not silently truncate and fall
    # back to local rules.
    planning_settings = ApiSettings.from_dict(active_settings.to_dict(include_secret=True))
    planning_settings.timeout_seconds = max(active_settings.timeout_seconds, 120)
    allowed = ", ".join(SUPPORTED_TYPES)
    request = {
        "用户需求": user_prompt,
        "输入文件名": input_file_names,
        "本地规则初判": draft.task_spec.task_type,
        "本地提取方案": draft.task_spec.options.get("content_plan", {}),
        "允许类型": SUPPORTED_TYPES,
    }
    result = chat_completion(
        planning_settings,
        system_prompt=(
            "你是资深 Excel 需求分析师。你的任务不是替用户补写一个看似完整但未经确认的需求，"
            "而是先准确提取已明确内容，再找出会实质影响最终工作簿的缺失信息。"
            "你不能读取文件内容、不能指定单元格地址、不能输出 Excel 公式或代码。"
            "你可以把用户明确要求的标题、列名、输入数据、计算语义、汇总、排序和图表整理成结构化方案。"
            "用户需求文字中出现的制表符、Markdown 或 CSV 表格属于正式输入数据，"
            "不能因为没有上传文件就判断为未提供数据。"
            "若本地提取方案包含 inline_tables，必须保留多表关系和各表用途。"
            "只返回一个 JSON 对象，不要 Markdown。字段必须是："
            "task_type、confidence、goal_summary、clarifying_questions、"
            "include_charts、include_summary、content_plan。"
            "content_plan 字段为对象，可包含 title、layout、columns、formula_rules、summary_rules。"
            "columns 每项只能包含 name、kind、role；kind 只能是 text/number/money/percentage/date/time，"
            "role 只能是 input/formula。formula_rules 只能使用 average、difference、product、ratio、"
            "sum、weather_advice 等计算语义，不得返回公式字符串。summary_rules 只能使用 "
            "averageif、sumif、countif、average、sum、count。"
            "请尽量根据表格类型给出一套合理、完整的默认列结构，让用户在确认页能直接看到将生成的列；"
            "但不要凭空添加用户明确表示不需要的图表或汇总页。"
            "clarifying_questions 只询问当前需求中确实缺失、且答案会改变最终表格的问题，"
            "用来请用户补充完善，而不是替用户臆断。"
            "问题必须具体、可直接回答，并在问题中给出简短示例；优先检查："
            "数据来源、明细列及顺序、计算口径、分组排序和小计总计、图表指标、"
            "日期范围、金额税率口径、是否严格保留参考模板。"
            "不得只问“还有什么要求”或“是否需要美化”这类空泛问题。"
            f"task_type 只能是：{allowed}。clarifying_questions 最多 6 个中文问题。"
        ),
        user_prompt=json.dumps(request, ensure_ascii=False),
        temperature=0.1,
        max_tokens=8000,
        json_mode=True,
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
    current_type = draft.task_spec.task_type
    if task_type in SUPPORTED_TYPES and (
        task_type == current_type
        or current_type == "generic_table"
        or draft.task_spec.confidence < 0.65
    ):
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
        model_columns = (
            content_plan.get("columns")
            if isinstance(content_plan.get("columns"), list)
            else []
        )
        has_input_files = bool(draft.task_spec.input_files)
        # Adopt the model's proposed column structure whenever the user did not
        # upload a data file (whose real columns must win). This makes the
        # confirmation preview reflect what will actually be built and gives the
        # generation step concrete columns to follow, instead of inventing its
        # own and drifting away from the request.
        may_use_custom_structure = bool(local_plan.get("explicit_structure")) or (
            bool(model_columns) and not has_input_files
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
    if isinstance(questions, list):
        model_questions = [str(item).strip() for item in questions if str(item).strip()]
        combined = list(dict.fromkeys([*draft.clarifying_questions, *model_questions]))
        draft.clarifying_questions = combined[:6]
    draft.task_spec.options["intent_source"] = "custom_api"
