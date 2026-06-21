"""Deterministic bridge from a confirmed TaskSpec to the existing Excel core."""

from __future__ import annotations

from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Any, Callable

from openpyxl import load_workbook

from ..api_settings import ApiSettings
from ..custom_workbook_builder import build_custom_workbook, build_dataset_workbook
from ..io_utils import read_table
from ..task_paths import TaskPaths, append_run_log_event, stage_input_files
from ..task_spec import TaskSpec, save_task_spec
from ..template_adapter import (
    apply_template_mode,
    inspect_template,
    prepare_template_file,
)
from ..workbook_builder import analyze_sales_file, create_workbook
from .llm_workbook_agent import generate_with_llm_agent


GENERATION_API_VERSION = 2


@dataclass
class GenerationResult:
    success: bool
    output_file: str | None
    message: str
    error: str | None
    used_command: str | None
    mode: str
    notices: list[str]
    agent_tool_calls: int = 0
    agent_rounds: int = 0
    blueprint_file: str | None = None

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


def generate_from_task_spec(
    task_spec: TaskSpec,
    task_paths: TaskPaths,
    api_settings: ApiSettings | None = None,
    progress: Callable[[str, str], None] | None = None,
) -> GenerationResult:
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
        if task_spec.template_files:
            task_spec.template_files = stage_input_files(
                task_spec.template_files,
                task_paths,
            )
            prepared_template = prepare_template_file(
                task_spec.template_files[0],
                task_paths.task_dir,
            )
            task_spec.options["prepared_template_file"] = str(prepared_template)
            task_spec.options["template_summary"] = inspect_template(prepared_template)
            if task_spec.options.get("use_template_data"):
                task_spec.input_files.append(str(prepared_template))
                task_spec.options["template_data_added_to_inputs"] = True
        if task_spec.input_files:
            profiles = []
            total_rows = 0
            all_columns: list[str] = []
            for source_path in task_spec.input_files:
                source_frame = read_table(source_path)
                columns = [str(item) for item in source_frame.columns]
                profiles.append(
                    {
                        "file_name": Path(source_path).name,
                        "row_count": int(len(source_frame)),
                        "columns": columns,
                    }
                )
                total_rows += int(len(source_frame))
                all_columns.extend(columns)
            task_spec.options["input_data_profile"] = {
                "files": profiles,
                "row_count": total_rows,
                "columns": list(dict.fromkeys(all_columns)),
                "values_sent_to_model": False,
            }
        save_task_spec(task_spec, task_paths.task_spec_file)

        settings = api_settings or ApiSettings()
        if settings.configured and settings.use_for_generation:
            agent = generate_with_llm_agent(
                task_spec,
                task_paths,
                settings,
                progress=progress,
            )
            if not agent.success:
                raise RuntimeError(agent.error or agent.message)
            template_notices = _apply_selected_template(task_spec, task_paths)
            save_task_spec(task_spec, task_paths.task_spec_file)
            result = GenerationResult(
                success=True,
                output_file=agent.output_file,
                message=agent.message,
                error=None,
                used_command="excel_agent.services.llm_workbook_agent.generate_with_llm_agent",
                mode="llm_tool_agent",
                notices=[*agent.notices, *template_notices],
                agent_tool_calls=agent.tool_calls,
                agent_rounds=agent.rounds,
                blueprint_file=agent.blueprint_file,
            )
            append_run_log_event(
                task_paths,
                event="generation_completed",
                status="success",
                details=result.to_dict(),
            )
            return result

        if progress:
            progress("build", "未启用大模型生成，正在使用本地确定性生成器……")
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
        notices.extend(_apply_selected_template(task_spec, task_paths))
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


def _apply_selected_template(
    task_spec: TaskSpec,
    task_paths: TaskPaths,
) -> list[str]:
    template_path = str(task_spec.options.get("prepared_template_file") or "")
    if not template_path:
        return []
    mode = str(task_spec.options.get("template_mode") or "reference")
    apply_template_mode(
        task_paths.output_file,
        template_path,
        mode=mode,
        blueprint=task_spec.options.get("agent_blueprint"),
        use_template_data=bool(task_spec.options.get("use_template_data")),
    )
    labels = {
        "reference": "已参考模板的字体、表头样式、列宽和页面设置。",
        "flexible": "已灵活套用模板形式；与本次需求冲突的内容以需求为准。",
        "strict": "已按严格模板模式写入；模板字段顺序和工作表结构保持不变。",
    }
    notice = labels.get(mode, labels["reference"])
    if not task_spec.options.get("use_template_data"):
        notice += " 模板中的示例数据未带入结果。"
    else:
        notice += " 模板中的现有数据已作为额外数据源。"
    return [notice]
