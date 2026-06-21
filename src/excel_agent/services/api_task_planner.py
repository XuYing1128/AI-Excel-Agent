"""Optional model-assisted intent understanding without workbook data."""

from __future__ import annotations

import json
from dataclasses import dataclass

from ..api_settings import ApiSettings
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
        "允许类型": SUPPORTED_TYPES,
    }
    result = chat_completion(
        settings,
        system_prompt=(
            "你是本地表格工具的需求分析助手。你不能读取文件内容、不能设计具体单元格、"
            "不能生成公式或代码。只根据用户文字和文件名判断需求。"
            "只返回一个 JSON 对象，不要 Markdown。字段必须是："
            "task_type、confidence、goal_summary、clarifying_questions、"
            "include_charts、include_summary。"
            f"task_type 只能是：{allowed}。clarifying_questions 最多 5 个中文问题。"
        ),
        user_prompt=json.dumps(request, ensure_ascii=False),
        temperature=0.1,
        max_tokens=600,
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
        draft.task_spec.include_charts = payload["include_charts"]
    if isinstance(payload.get("include_summary"), bool):
        draft.task_spec.include_summary = payload["include_summary"]

    questions = payload.get("clarifying_questions", [])
    if isinstance(questions, list):
        model_questions = [str(item).strip() for item in questions if str(item).strip()]
        combined = list(dict.fromkeys([*draft.clarifying_questions, *model_questions]))
        draft.clarifying_questions = combined[:5]
    draft.task_spec.options["intent_source"] = "custom_api"
