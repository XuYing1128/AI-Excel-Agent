"""Thin wrappers around the existing deterministic Excel core."""

from __future__ import annotations

import datetime
from decimal import Decimal
from pathlib import Path
from typing import Any

from ....custom_workbook_builder import build_dataset_workbook
from ....domain_builders import build_performance_compensation_workbook
from ....io_utils import read_table
from ....preview import workbook_preview
from ....rich_workbook_builder import build_rich_workbook, normalize_workbook_blueprint
from ....schedule_planner import (
    parse_capacity,
    parse_day_count,
    parse_rooms,
    parse_slots,
    schedule_registrations,
)
from ....template_filler import fill_template
from ....validators import inspect_workbook, validate_workbook
from ...recalc import describe_error_cells, recalc_workbook
from .base import AgentTool, ToolContext, ToolResult


def excel_tools() -> list[AgentTool]:
    return [
        AgentTool(
            "read_table_summary",
            "读取任务输入目录中的表格摘要，只返回列名、行数、类型和前 5 行样例。",
            _schema({"path": {"type": "string"}}),
            _read_table_summary,
        ),
        AgentTool(
            "inspect_workbook",
            "检查已生成工作簿结构摘要。",
            _schema({"path": {"type": "string"}}),
            _inspect_workbook,
        ),
        AgentTool(
            "fill_template",
            "按上传模板精确填入任务输入数据，保留模板工作表、表头和样式。",
            _schema({}),
            _fill_template,
        ),
        AgentTool(
            "schedule_exam",
            "根据报名记录、考场、容量和时段生成排考分配摘要。",
            _schema({}),
            _schedule_exam,
        ),
        AgentTool(
            "build_performance_compensation",
            "用本地绩效薪酬业务编译器生成参数表和明细表。",
            _schema({}),
            _build_performance,
        ),
        AgentTool(
            "build_workbook",
            "提交结构化工作簿方案，由本地工具生成复杂 xlsx。",
            _schema({"blueprint": {"type": "object"}}),
            _build_rich,
        ),
        AgentTool(
            "build_rich_workbook",
            "提交结构化工作簿方案，由本地工具生成复杂 xlsx。",
            _schema({"blueprint": {"type": "object"}}),
            _build_rich,
        ),
        AgentTool(
            "build_dataset",
            "根据上传数据文件生成原始数据、清洗数据、汇总和可选图表。",
            _schema({"path": {"type": "string"}, "title": {"type": "string"}}),
            _build_dataset,
        ),
        AgentTool(
            "validate_workbook",
            "运行确定性工作簿校验，返回 pass/warn/fail 摘要。",
            _schema({"path": {"type": "string"}}),
            _validate,
        ),
        AgentTool(
            "recalc_check",
            "用 LibreOffice 真算 OUTPUT_FILE，列出算出 #VALUE!/循环引用等错误的单元格（finish 前自查公式用）。",
            _schema({"path": {"type": "string"}}),
            _recalc_check,
        ),
        AgentTool(
            "render_preview",
            "生成工作簿页面预览摘要，预览不可用时不阻塞主流程。",
            _schema({"path": {"type": "string"}}),
            _render_preview,
        ),
        AgentTool(
            "finish_task",
            "确认任务完成。只有生成文件存在并检查后才能调用。",
            _schema({"summary": {"type": "string"}}),
            _finish,
        ),
    ]


def _schema(properties: dict[str, Any]) -> dict[str, Any]:
    return {"type": "object", "properties": properties}


def _resolve_task_path(raw_path: str | None, ctx: ToolContext, *, default_output: bool = False) -> Path:
    if not raw_path:
        return ctx.task_paths.output_file if default_output else ctx.task_paths.input_dir
    raw = str(raw_path).strip().strip('"').strip("'")
    task_root = ctx.task_paths.task_dir.resolve()
    candidate = Path(raw)
    if candidate.is_absolute():
        resolved = candidate.resolve()
    else:
        # 智能体代码的工作目录在 temp_dir，常用 ../output/ 这类相对路径。
        # 依次按 temp_dir / output_dir / task_dir / input_dir 解析，取落在任务
        # 目录内且确实存在的那个；都不存在时再按默认落点（用于尚未生成的输出）。
        resolved = None
        for base in (
            ctx.temp_dir,
            ctx.task_paths.output_dir,
            ctx.task_paths.task_dir,
            ctx.task_paths.input_dir,
        ):
            trial = (Path(base) / raw).resolve()
            inside = trial == task_root or trial.is_relative_to(task_root)
            if inside and trial.exists():
                resolved = trial
                break
        if resolved is None:
            base = ctx.task_paths.output_file.parent if default_output else ctx.task_paths.task_dir
            resolved = (Path(base) / Path(raw).name).resolve()
    if not (resolved == task_root or resolved.is_relative_to(task_root)):
        raise PermissionError("工具只能访问当前任务目录中的文件。")
    return resolved


def _json_safe(obj: Any) -> Any:
    """把工具返回值里非 JSON 原生类型（datetime/Decimal 等）转成可序列化形式。"""
    if isinstance(obj, (datetime.datetime, datetime.date, datetime.time)):
        return obj.isoformat()
    if isinstance(obj, Decimal):
        return float(obj)
    if isinstance(obj, dict):
        return {str(k): _json_safe(v) for k, v in obj.items()}
    if isinstance(obj, (list, tuple)):
        return [_json_safe(v) for v in obj]
    return obj


def _read_table_summary(args: dict[str, Any], ctx: ToolContext) -> ToolResult:
    try:
        path = _resolve_task_path(str(args.get("path") or ""), ctx)
        frame = read_table(path)
        sample = frame.head(5).astype(object).where(frame.notna(), "").to_dict(orient="records")
        # 日期/时间/Decimal 等非 JSON 原生类型要转成字符串，否则下游 json.dumps 会抛
        # TypeError（datetime is not JSON serializable）导致整轮生成中断。
        sample = [_json_safe(row) for row in sample]
        return ToolResult(
            True,
            f"读取 {path.name}：{len(frame)} 行，{len(frame.columns)} 列。",
            {
                "path": str(path),
                "row_count": int(len(frame)),
                "columns": [str(item) for item in frame.columns],
                "dtypes": {str(k): str(v) for k, v in frame.dtypes.items()},
                "sample_rows": sample,
            },
        )
    except Exception as exc:
        return ToolResult(False, "读取表格摘要失败。", error=f"{type(exc).__name__}: {exc}")


def _inspect_workbook(args: dict[str, Any], ctx: ToolContext) -> ToolResult:
    try:
        path = _resolve_task_path(str(args.get("path") or ""), ctx, default_output=True)
        data = inspect_workbook(path)
        return ToolResult(True, "工作簿结构摘要已生成。", data=data, artifacts=[str(path)])
    except Exception as exc:
        return ToolResult(False, "检查工作簿失败。", error=f"{type(exc).__name__}: {exc}")


def _fill_template(args: dict[str, Any], ctx: ToolContext) -> ToolResult:
    try:
        if not ctx.task_spec.template_files or not ctx.task_spec.input_files:
            return ToolResult(False, "缺少模板或数据文件。", error="missing_template_or_data")
        result = fill_template(
            ctx.task_spec.template_files[0],
            ctx.task_spec.input_files,
            ctx.task_paths.output_file,
            prompt=ctx.task_spec.user_goal,
            task_dir=ctx.task_paths.task_dir,
        )
        return ToolResult(True, "已按模板填充工作簿。", data=result, artifacts=[str(ctx.task_paths.output_file)])
    except Exception as exc:
        return ToolResult(False, "模板填充失败。", error=f"{type(exc).__name__}: {exc}")


def _schedule_exam(args: dict[str, Any], ctx: ToolContext) -> ToolResult:
    try:
        if not ctx.task_spec.input_files:
            return ToolResult(False, "缺少报名数据。", error="missing_input")
        frame = read_table(ctx.task_spec.input_files[0])
        registrations = frame.to_dict(orient="records")
        slots = parse_slots(ctx.task_spec.user_goal, parse_day_count(ctx.task_spec.user_goal))
        assignments = schedule_registrations(
            registrations,
            rooms=parse_rooms(ctx.task_spec.user_goal),
            seats_per_room=parse_capacity(ctx.task_spec.user_goal),
            slots=slots,
        )
        placed = sum(1 for item in assignments if item is not None)
        return ToolResult(
            placed == len(assignments),
            f"已分配 {placed}/{len(assignments)} 条报名记录。",
            {"assigned_count": placed, "total": len(assignments)},
        )
    except Exception as exc:
        return ToolResult(False, "排考失败。", error=f"{type(exc).__name__}: {exc}")


def _build_performance(args: dict[str, Any], ctx: ToolContext) -> ToolResult:
    try:
        plan = ctx.task_spec.options.get("content_plan", {})
        build_performance_compensation_workbook(plan, ctx.task_spec.user_goal, ctx.task_paths.output_file)
        return ToolResult(True, "已生成绩效薪酬工作簿。", artifacts=[str(ctx.task_paths.output_file)])
    except Exception as exc:
        return ToolResult(False, "绩效薪酬生成失败。", error=f"{type(exc).__name__}: {exc}")


def _build_rich(args: dict[str, Any], ctx: ToolContext) -> ToolResult:
    try:
        blueprint = normalize_workbook_blueprint(args.get("blueprint", args))
        build_rich_workbook(blueprint, ctx.task_paths.output_file, require_charts=ctx.task_spec.include_charts)
        return ToolResult(True, "已生成复杂工作簿。", artifacts=[str(ctx.task_paths.output_file)])
    except Exception as exc:
        return ToolResult(False, "复杂工作簿生成失败。", error=f"{type(exc).__name__}: {exc}")


def _build_dataset(args: dict[str, Any], ctx: ToolContext) -> ToolResult:
    try:
        source = _resolve_task_path(str(args.get("path") or ""), ctx)
        title = str(args.get("title") or ctx.task_spec.output_name or ctx.task_spec.user_goal[:40])
        build_dataset_workbook(
            source,
            ctx.task_paths.output_file,
            title=title,
            include_summary=ctx.task_spec.include_summary,
            include_charts=ctx.task_spec.include_charts,
        )
        return ToolResult(True, "已根据数据文件生成工作簿。", artifacts=[str(ctx.task_paths.output_file)])
    except Exception as exc:
        return ToolResult(False, "数据工作簿生成失败。", error=f"{type(exc).__name__}: {exc}")


def _validate(args: dict[str, Any], ctx: ToolContext) -> ToolResult:
    try:
        path = _resolve_task_path(str(args.get("path") or ""), ctx, default_output=True)
        report = validate_workbook(path)
        return ToolResult(
            report.get("status") in {"pass", "warn"},
            f"校验状态：{report.get('status')}",
            {
                "status": report.get("status"),
                "error_count": len(report.get("errors", [])),
                "warning_count": len(report.get("warnings", [])),
            },
            [str(path)],
        )
    except Exception as exc:
        return ToolResult(False, "校验失败。", error=f"{type(exc).__name__}: {exc}")


def _recalc_check(args: dict[str, Any], ctx: ToolContext) -> ToolResult:
    try:
        path = _resolve_task_path(str(args.get("path") or ""), ctx, default_output=True)
        report = recalc_workbook(path)
        if not report.get("available"):
            return ToolResult(True, "本机未安装 LibreOffice，已跳过真算自查。", data=report)
        cells = report.get("error_cells", [])
        summary = (
            "真算通过，无 #VALUE!/循环引用。"
            if not cells
            else f"真算发现 {len(cells)} 处报错：{describe_error_cells(cells)}"
        )
        return ToolResult(
            bool(report.get("ok")),
            summary,
            data={"ok": report.get("ok"), "error_cells": cells[:30]},
            artifacts=[str(path)],
        )
    except Exception as exc:
        return ToolResult(False, "真算自查失败。", error=f"{type(exc).__name__}: {exc}")


def _render_preview(args: dict[str, Any], ctx: ToolContext) -> ToolResult:
    try:
        path = _resolve_task_path(str(args.get("path") or ""), ctx, default_output=True)
        preview = workbook_preview(path, max_rows=20, max_columns=12)
        return ToolResult(True, "已生成工作簿预览摘要。", data=preview, artifacts=[str(path)])
    except Exception as exc:
        return ToolResult(False, "预览失败。", error=f"{type(exc).__name__}: {exc}")


def _finish(args: dict[str, Any], ctx: ToolContext) -> ToolResult:
    output = ctx.task_paths.output_file
    if not output.exists():
        # run_python 自由生成可能把工作簿写成别的文件名（仍在 output 目录）。这里扫最新 .xlsx
        # 认领为标准产物，避免 finish_task 误报 output_missing 害模型无谓重试/丢表（与编排层
        # 末尾的 _claim_produced_workbook 同一思路，统一完成判定基准）。
        try:
            import shutil

            candidates = [p for p in ctx.task_paths.output_dir.glob("*.xlsx") if p.is_file()]
            if candidates:
                newest = max(candidates, key=lambda p: p.stat().st_mtime)
                if newest.resolve() != output.resolve():
                    shutil.copyfile(newest, output)
        except OSError:
            pass
    ok = output.exists()
    return ToolResult(
        ok,
        str(args.get("summary") or ("任务已完成。" if ok else "还没有生成文件。")),
        artifacts=[str(output)] if ok else [],
        error=None if ok else "output_missing",
    )
