"""Two fixed structured reviewers: requirement fit and Excel usability."""

from __future__ import annotations

import json
from typing import Any, Callable

from ...api_settings import ApiSettings
from ..custom_api_service import ApiCallResult, chat_completion, parse_json_object


ChatFunc = Callable[..., ApiCallResult]


def run_structured_reviews(
    context: dict[str, Any],
    settings: ApiSettings,
    *,
    chat_func: ChatFunc = chat_completion,
) -> list[dict[str, Any]]:
    return [
        _run_one_review(
            "requirement_review",
            _requirement_prompt(),
            context,
            settings,
            chat_func=chat_func,
        ),
        _run_one_review(
            "excel_usability_review",
            _usability_prompt(),
            context,
            settings,
            chat_func=chat_func,
        ),
    ]


def _run_one_review(
    reviewer: str,
    system_prompt: str,
    context: dict[str, Any],
    settings: ApiSettings,
    *,
    chat_func: ChatFunc,
) -> dict[str, Any]:
    response = chat_func(
        settings,
        system_prompt=system_prompt,
        user_prompt=json.dumps(context, ensure_ascii=False),
        temperature=0.1,
        max_tokens=4000,
        json_mode=True,
    )
    if not response.success:
        return {
            "reviewer": reviewer,
            "model": settings.provider_name,
            "status": "warn",
            "fit_to_user_goal": "warn",
            "over_design_risk": "medium",
            "issues": [f"审查调用失败：{response.error}"],
            "concerns": [f"审查调用失败：{response.error}"],
            "suggestions": ["可直接下载文件；如需要主观审查，请稍后重试。"],
            "requires_user_confirmation": False,
            "error": response.error,
        }
    try:
        payload = parse_json_object(response.content)
    except (ValueError, TypeError) as exc:
        return {
            "reviewer": reviewer,
            "model": settings.provider_name,
            "status": "warn",
            "fit_to_user_goal": "warn",
            "over_design_risk": "medium",
            "issues": [f"审查结果无法解析：{exc}"],
            "concerns": [f"审查结果无法解析：{exc}"],
            "suggestions": ["可直接下载文件；如需要主观审查，请稍后重试。"],
            "requires_user_confirmation": False,
            "error": str(exc),
        }
    return _normalize(reviewer, payload, settings.provider_name)


def _normalize(reviewer: str, payload: dict[str, Any], model_name: str) -> dict[str, Any]:
    status = str(payload.get("status", "warn")).lower()
    if status not in {"pass", "warn", "fail"}:
        status = "warn"
    issues = _string_list(payload.get("issues", payload.get("concerns", [])))
    suggestions = _string_list(payload.get("suggestions", []))
    return {
        "reviewer": reviewer,
        "model": model_name,
        "status": status,
        "fit_to_user_goal": str(payload.get("fit_to_user_goal", status)).lower(),
        "over_design_risk": str(payload.get("over_design_risk", "low")).lower(),
        "issues": issues,
        "concerns": issues,
        "suggestions": suggestions,
        "requires_user_confirmation": bool(payload.get("requires_user_confirmation", status == "fail")),
    }


def _string_list(value: Any) -> list[str]:
    if not isinstance(value, list):
        return []
    return [str(item).strip() for item in value if str(item).strip()]


def _requirement_prompt() -> str:
    return (
        "你是表格交付的需求一致性审查器，只看摘要，不看完整数据。"
        "检查标题、sheet、列、公式语义、排序、小计总计、图表要求和用户确认项是否一致。"
        "不要建议删除用户已确认的图表、说明页或汇总。"
        "只返回 JSON：status(pass/warn/fail)、issues、suggestions、"
        "requires_user_confirmation、fit_to_user_goal、over_design_risk。"
    )


def _usability_prompt() -> str:
    return (
        "你是 Excel 可用性审查器，只看工作簿结构和校验摘要。"
        "检查公式可维护性、冻结窗格、筛选、列宽、图表可读性、是否过度复杂、"
        "是否适合非技术用户继续修改。不要重复确定性校验已通过的客观事实。"
        "只返回 JSON：status(pass/warn/fail)、issues、suggestions、"
        "requires_user_confirmation、fit_to_user_goal、over_design_risk。"
    )

