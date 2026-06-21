"""Optional non-blocking subjective review placeholder for V1."""

from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Any

from ..task_paths import TaskPaths, append_run_log_event
from ..task_spec import TaskSpec


def run_subjective_review(
    task_spec: TaskSpec,
    validation_summary: dict[str, Any],
    workbook_summary: dict[str, Any],
    generation_summary: dict[str, Any],
    task_paths: TaskPaths,
) -> dict[str, Any]:
    """Write a safe disabled review result unless a future connector is added.

    V1 deliberately sends no workbook rows or cell values to an external model.
    """

    requested = os.getenv("AI_EXCEL_ENABLE_SUBJECTIVE_REVIEW", "").strip().lower() in {
        "1",
        "true",
        "yes",
    }
    if requested:
        reason = (
            "检测到主观审查开关，但 V1 未配置外部模型连接器。"
            "未上传任何工作簿数据，文件下载不受影响。"
        )
    else:
        reason = "主观模型审查未启用，不影响文件下载。"

    result = {
        "enabled": False,
        "reviews": [],
        "agreement": "not_run",
        "user_notice": reason,
        "input_policy": {
            "allowed": ["TaskSpec", "确定性校验摘要", "工作簿结构摘要", "生成日志摘要"],
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
    Path(task_paths.subjective_review_report).write_text(
        json.dumps(result, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    append_run_log_event(
        task_paths,
        event="subjective_review_skipped",
        status="skipped",
        details={"reason": reason},
    )
    return result
