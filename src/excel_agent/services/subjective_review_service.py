"""Optional non-blocking subjective review using the configured custom API."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from ..api_settings import ApiSettings
from ..task_paths import TaskPaths, append_run_log_event
from ..task_spec import TaskSpec
from .custom_api_service import chat_completion, parse_json_object


def run_subjective_review(
    task_spec: TaskSpec,
    validation_summary: dict[str, Any],
    workbook_summary: dict[str, Any],
    generation_summary: dict[str, Any],
    task_paths: TaskPaths,
    api_settings: ApiSettings | None = None,
) -> dict[str, Any]:
    """Review only summaries; never send workbook rows or cell values."""

    settings = api_settings or ApiSettings()
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
    response = chat_completion(
        settings,
        system_prompt=(
            "你是表格交付的主观审查助手。你只能评价是否符合用户目标、说明是否清楚、"
            "是否有过度设计风险。不能评价客观正确性，不能输出代码，不能要求修改单元格。"
            "只返回 JSON 对象，不要 Markdown。字段为：status、fit_to_user_goal、"
            "over_design_risk、concerns、suggestions。status 只能是 pass 或 warn；"
            "over_design_risk 只能是 low、medium、high。concerns 和 suggestions 必须是中文数组。"
        ),
        user_prompt=json.dumps(safe_context, ensure_ascii=False),
        temperature=0.1,
        max_tokens=700,
    )
    if not response.success:
        result = _disabled_result(
            f"建议审查调用失败，不影响文件下载：{response.error}",
            task_spec,
            validation_summary,
            workbook_summary,
            generation_summary,
        )
        result["error"] = response.error
        _save_result(result, task_paths)
        append_run_log_event(
            task_paths,
            event="subjective_review_failed",
            status="warn",
            details={"error": response.error},
        )
        return result

    try:
        payload = parse_json_object(response.content)
        review = _normalize_review(payload, settings.provider_name)
    except (ValueError, TypeError) as exc:
        result = _disabled_result(
            f"建议审查结果无法解析，不影响文件下载：{exc}",
            task_spec,
            validation_summary,
            workbook_summary,
            generation_summary,
        )
        result["error"] = str(exc)
        _save_result(result, task_paths)
        append_run_log_event(
            task_paths,
            event="subjective_review_failed",
            status="warn",
            details={"error": str(exc)},
        )
        return result

    result = {
        "enabled": True,
        "reviews": [review],
        "agreement": "single_model",
        "user_notice": (
            "建议审查未发现明显问题，但最终以你的确认和确定性校验为准。"
            if review["status"] == "pass"
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
        "revision_prompt": _revision_prompt_from_review(review),
    }
    _save_result(result, task_paths)
    append_run_log_event(
        task_paths,
        event="subjective_review_completed",
        status="success",
        details={"provider": settings.provider_name, "review_status": review["status"]},
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
                }
            )
    return {"sheet_count": summary.get("sheet_count", len(sheets)), "sheets": sheets}


def _normalize_review(payload: dict[str, Any], provider_name: str) -> dict[str, Any]:
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
    return {
        "model": provider_name,
        "status": status,
        "fit_to_user_goal": fit,
        "over_design_risk": risk,
        "concerns": [str(item) for item in concerns] if isinstance(concerns, list) else [],
        "suggestions": [str(item) for item in suggestions] if isinstance(suggestions, list) else [],
    }


def _revision_prompt_from_review(review: dict[str, Any]) -> str:
    lines = [
        *[f"修正问题：{item}" for item in review.get("concerns", [])],
        *[f"采用建议：{item}" for item in review.get("suggestions", [])],
    ]
    return "\n".join(lines)
