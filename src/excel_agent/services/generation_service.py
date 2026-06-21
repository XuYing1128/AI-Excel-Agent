"""Deterministic bridge from a confirmed TaskSpec to the existing Excel core."""

from __future__ import annotations

from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Any

from openpyxl import load_workbook

from ..custom_workbook_builder import build_custom_workbook, build_dataset_workbook
from ..task_paths import TaskPaths, append_run_log_event, stage_input_files
from ..task_spec import TaskSpec, save_task_spec
from ..workbook_builder import analyze_sales_file, create_workbook


@dataclass
class GenerationResult:
    success: bool
    output_file: str | None
    message: str
    error: str | None
    used_command: str | None
    mode: str
    notices: list[str]

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


def generate_from_task_spec(task_spec: TaskSpec, task_paths: TaskPaths) -> GenerationResult:
    """Generate one workbook without allowing a model to touch cells."""

    append_run_log_event(
        task_paths,
        event="generation_started",
        status="running",
        details={
            "task_type": task_spec.task_type,
            "policy": task_spec.options.get("generation_policy"),
            "model_may_edit_cells": False,
        },
    )
    notices: list[str] = []
    try:
        save_task_spec(task_spec, task_paths.task_spec_file)
        if task_spec.input_files:
            task_spec.input_files = stage_input_files(task_spec.input_files, task_paths)
        save_task_spec(task_spec, task_paths.task_spec_file)

        mode = "standard_template"
        used_command = "excel_agent.workbook_builder.create_workbook"
        policy = str(task_spec.options.get("generation_policy", ""))
        content_plan = task_spec.options.get("content_plan", {})
        if (
            policy == "custom_content"
            and isinstance(content_plan, dict)
            and content_plan.get("explicit_structure")
            and not task_spec.input_files
        ):
            mode = "custom_content"
            used_command = "excel_agent.custom_workbook_builder.build_custom_workbook"
            build_custom_workbook(
                content_plan,
                task_paths.output_file,
                include_charts=task_spec.include_charts,
            )
        elif task_spec.task_type == "sales_report" and task_spec.input_files:
            mode = "sales_input_analysis"
            used_command = "excel_agent.workbook_builder.analyze_sales_file"
            analyze_sales_file(task_spec.input_files[0], task_paths.output_file)
        elif task_spec.input_files:
            mode = "input_dataset"
            used_command = "excel_agent.custom_workbook_builder.build_dataset_workbook"
            build_dataset_workbook(
                task_spec.input_files[0],
                task_paths.output_file,
                title=content_plan.get("title") or task_spec.user_goal[:60],
                include_summary=task_spec.include_summary,
                include_charts=task_spec.include_charts,
            )
            notices.append("已按上传文件的实际字段生成结果，没有套用无关示例模板。")
        else:
            create_workbook(task_spec.task_type, task_paths.output_file)
            if task_spec.task_type == "sales_report":
                notices.append("未提供销售输入文件，已生成销售报表标准模板示例。")

        notices.extend(_apply_task_options(task_paths.output_file, task_spec))
        if not task_paths.output_file.exists():
            raise FileNotFoundError(f"生成内核未产生预期文件: {task_paths.output_file}")

        result = GenerationResult(
            success=True,
            output_file=str(task_paths.output_file),
            message="电子表格已由本地确定性程序生成。",
            error=None,
            used_command=used_command,
            mode=mode,
            notices=notices,
        )
        append_run_log_event(
            task_paths,
            event="generation_completed",
            status="success",
            details=result.to_dict(),
        )
        return result
    except Exception as exc:
        result = GenerationResult(
            success=False,
            output_file=str(task_paths.output_file) if task_paths.output_file.exists() else None,
            message="电子表格生成失败。",
            error=f"{type(exc).__name__}: {exc}",
            used_command=None,
            mode="error",
            notices=notices,
        )
        append_run_log_event(
            task_paths,
            event="generation_failed",
            status="error",
            details=result.to_dict(),
        )
        return result


def _apply_task_options(output_file: Path, task_spec: TaskSpec) -> list[str]:
    """Apply small deterministic presentation choices after core generation."""

    notices: list[str] = []
    wb = load_workbook(output_file)
    changed = False

    if not task_spec.include_charts:
        for ws in wb.worksheets:
            if ws._charts:
                ws._charts = []
                changed = True

    for summary_name in ("Summary", "汇总"):
        if task_spec.include_summary or summary_name not in wb.sheetnames:
            continue
        if task_spec.task_type == "dashboard":
            notices.append("综合仪表盘依赖汇总页，已为避免公式失效保留汇总页。")
        else:
            wb.remove(wb[summary_name])
            changed = True

    for instruction_name in ("Instructions", "说明"):
        if (
            not task_spec.include_instructions_sheet
            and instruction_name in wb.sheetnames
            and len(wb.sheetnames) > 1
        ):
            wb.remove(wb[instruction_name])
            changed = True

    if task_spec.preserve_template_style and task_spec.input_files:
        notices.append(
            "当前版本仅在现有内核支持时保留模板结构；标准模板生成不会复制任意上传工作簿的全部样式。"
        )

    if changed:
        wb.save(output_file)
    return notices
