"""Optional non-blocking subjective review using the configured custom API."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from ..api_settings import ApiSettings
from ..model_registry import get_role_api_settings
from ..task_paths import TaskPaths, append_run_log_event
from ..task_spec import TaskSpec
from .custom_api_service import chat_completion
from .review.structured_review import run_structured_reviews


def run_subjective_review(
    task_spec: TaskSpec,
    validation_summary: dict[str, Any],
    workbook_summary: dict[str, Any],
    generation_summary: dict[str, Any],
    task_paths: TaskPaths,
    api_settings: ApiSettings | None = None,
) -> dict[str, Any]:
    """Review only summaries; never send workbook rows or cell values."""

    reviewer_settings = get_role_api_settings(
        "reviewer",
        use_for_intent=False,
        use_for_review=True,
        use_for_generation=False,
    )
    settings = (
        reviewer_settings
        if reviewer_settings is not None and reviewer_settings.configured and reviewer_settings.use_for_review
        else api_settings
        if api_settings is not None and api_settings.configured and api_settings.use_for_review
        else ApiSettings()
    )
    if not settings.configured or not settings.use_for_review:
        result = _disabled_result(
            "建议审查未启用，不影响文件下载。",
            task_spec,
            validation_summary,
            workbook_summary,
            generation_summary,
        )
        _save_result(result, task_paths)
        append_run_log_event(
            task_paths,
            event="subjective_review_skipped",
            status="skipped",
            details={"reason": result["user_notice"]},
        )
        return result

    safe_context = {
        "任务方案": _safe_task_spec(task_spec),
        "确定性校验摘要": _compact_validation_summary(validation_summary),
        "工作簿结构摘要": _compact_workbook_summary(workbook_summary),
        "生成日志摘要": {
            "mode": generation_summary.get("mode"),
            "message": generation_summary.get("message"),
            "notices": generation_summary.get("notices", []),
        },
    }
    reviews = [
        _filter_structured_review(review, task_spec)
        for review in run_structured_reviews(
            safe_context,
            settings,
            chat_func=chat_completion,
        )
    ]
    has_issue = any(review.get("status") in {"warn", "fail"} for review in reviews)

    result = {
        "enabled": True,
        "reviews": reviews,
        "agreement": "structured_dual_review",
        "user_notice": (
            "建议审查未发现明显问题，但最终以你的确认和确定性校验为准。"
            if not has_issue
            else "建议审查发现需要留意的主观问题，请查看下面的建议。"
        ),
        "input_policy": {
            "allowed": ["任务方案", "确定性校验摘要", "工作簿结构摘要", "生成日志摘要"],
            "full_workbook_data_sent": False,
            "model_may_edit_cells": False,
        },
        "context_summary": {
            "task_type": task_spec.task_type,
            "validation_status": validation_summary.get("status"),
            "sheet_count": workbook_summary.get("sheet_count"),
            "generation_mode": generation_summary.get("mode"),
        },
        "revision_prompt": _revision_prompt_from_reviews(reviews),
    }
    _save_result(result, task_paths)
    append_run_log_event(
        task_paths,
        event="subjective_review_completed",
        status="success",
        details={
            "provider": settings.provider_name,
            "review_statuses": [review.get("status") for review in reviews],
        },
    )
    return result


def _disabled_result(
    reason: str,
    task_spec: TaskSpec,
    validation_summary: dict[str, Any],
    workbook_summary: dict[str, Any],
    generation_summary: dict[str, Any],
) -> dict[str, Any]:
    return {
        "enabled": False,
        "reviews": [],
        "agreement": "not_run",
        "user_notice": reason,
        "input_policy": {
            "allowed": ["任务方案", "确定性校验摘要", "工作簿结构摘要", "生成日志摘要"],
            "full_workbook_data_sent": False,
            "model_may_edit_cells": False,
        },
        "context_summary": {
            "task_type": task_spec.task_type,
            "validation_status": validation_summary.get("status"),
            "sheet_count": workbook_summary.get("sheet_count"),
            "generation_mode": generation_summary.get("mode"),
        },
    }


def _save_result(result: dict[str, Any], task_paths: TaskPaths) -> None:
    Path(task_paths.subjective_review_report).write_text(
        json.dumps(result, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def _compact_validation_summary(summary: dict[str, Any]) -> dict[str, Any]:
    allowed = {
        "status",
        "sheet_count",
        "visible_sheet_count",
        "formula_cell_count",
        "error_count",
        "warning_count",
        "instruction_sheet_present",
        "data_sheet_present",
    }
    return {key: summary.get(key) for key in allowed if key in summary}


def _safe_task_spec(task_spec: TaskSpec) -> dict[str, Any]:
    data = task_spec.to_dict()
    data["input_files"] = [Path(item).name for item in task_spec.input_files]
    data["user_answers"] = {
        key: value
        for key, value in task_spec.user_answers.items()
        if key not in {"api_key", "接口密钥"}
    }
    return data


def _compact_workbook_summary(summary: dict[str, Any]) -> dict[str, Any]:
    sheets = []
    for item in summary.get("sheets", []):
        if isinstance(item, dict):
            sheets.append(
                {
                    "name": item.get("name"),
                    "max_row": item.get("max_row"),
                    "max_column": item.get("max_column"),
                    "hidden": item.get("hidden"),
                    "title": item.get("title"),
                    "headers": item.get("headers", []),
                    "formula_columns": item.get("formula_columns", []),
                    "chart_count": item.get("chart_count", 0),
                    # 带上冻结/筛选状态，reviewer 才不会因为看不到而瞎猜“未冻结/未筛选”。
                    "freeze_panes": item.get("freeze_panes"),
                    "auto_filter": item.get("auto_filter"),
                }
            )
    return {"sheet_count": summary.get("sheet_count", len(sheets)), "sheets": sheets}


def _normalize_review(
    payload: dict[str, Any],
    provider_name: str,
    task_spec: TaskSpec | None = None,
) -> dict[str, Any]:
    status = str(payload.get("status", "warn")).lower()
    if status not in {"pass", "warn"}:
        status = "warn"
    fit = str(payload.get("fit_to_user_goal", status)).lower()
    if fit not in {"pass", "warn"}:
        fit = status
    risk = str(payload.get("over_design_risk", "medium")).lower()
    if risk not in {"low", "medium", "high"}:
        risk = "medium"
    concerns = payload.get("concerns", [])
    suggestions = payload.get("suggestions", [])
    concern_items = [str(item) for item in concerns] if isinstance(concerns, list) else []
    suggestion_items = [str(item) for item in suggestions] if isinstance(suggestions, list) else []
    if task_spec is not None:
        concern_items = [
            item
            for item in concern_items
            if not _contradicts_confirmed_options(item, task_spec)
        ]
        suggestion_items = [
            item
            for item in suggestion_items
            if not _contradicts_confirmed_options(item, task_spec)
        ]
        if not concern_items and status == "warn":
            status = "pass"
            fit = "pass"
    return {
        "model": provider_name,
        "status": status,
        "fit_to_user_goal": fit,
        "over_design_risk": risk,
        "concerns": concern_items,
        "suggestions": suggestion_items,
    }


def _revision_prompt_from_review(review: dict[str, Any]) -> str:
    lines = [
        *[f"修正问题：{item}" for item in review.get("concerns", [])],
        *[f"采用建议：{item}" for item in review.get("suggestions", [])],
    ]
    return "\n".join(lines)


def _revision_prompt_from_reviews(reviews: list[dict[str, Any]]) -> str:
    lines: list[str] = []
    for review in reviews:
        prompt = _revision_prompt_from_review(review)
        if prompt:
            lines.extend(prompt.splitlines())
    return "\n".join(dict.fromkeys(lines))


def _filter_structured_review(
    review: dict[str, Any],
    task_spec: TaskSpec,
) -> dict[str, Any]:
    filtered = dict(review)
    concerns = [
        str(item)
        for item in filtered.get("concerns", filtered.get("issues", []))
        if not _contradicts_confirmed_options(str(item), task_spec)
    ]
    suggestions = [
        str(item)
        for item in filtered.get("suggestions", [])
        if not _contradicts_confirmed_options(str(item), task_spec)
    ]
    filtered["concerns"] = concerns
    filtered["issues"] = concerns
    filtered["suggestions"] = suggestions
    if not concerns and filtered.get("status") == "warn":
        filtered["status"] = "pass"
    return filtered


def _contradicts_confirmed_options(text: str, task_spec: TaskSpec) -> bool:
    lowered = str(text).lower()
    if task_spec.include_charts and any(
        word in lowered
        for word in (
            "图表未要求",
            "未要求图表",
            "未要求的图表",
            "删除图表",
            "移除图表",
            "去掉图表",
        )
    ):
        return True
    if task_spec.include_summary and any(
        word in lowered
        for word in ("汇总页未要求", "删除汇总页", "移除汇总页", "去掉汇总页")
    ):
        return True
    if any(
        word in lowered
        for word in ("删除说明工作表", "移除说明工作表", "说明工作表多余", "说明页多余")
    ):
        return True
    return False
