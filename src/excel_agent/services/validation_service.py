"""Adapter around the existing deterministic workbook validator."""

from __future__ import annotations

import json
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Any

from openpyxl import load_workbook

from ..task_paths import TaskPaths, append_run_log_event
from ..task_spec import TaskSpec
from ..validators import validate_workbook


@dataclass
class ValidationResult:
    status: str
    issues: list[dict[str, Any]]
    warnings: list[dict[str, Any]]
    suggestions: list[str]
    report_file: str
    summary: dict[str, Any]

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


def validate_generated_workbook(
    output_file: str | Path,
    task_spec: TaskSpec,
    task_paths: TaskPaths,
) -> ValidationResult:
    append_run_log_event(
        task_paths,
        event="validation_started",
        status="running",
        details={"output_file": str(output_file), "validator": "excel_agent.validators.validate_workbook"},
    )
    try:
        report = validate_workbook(output_file, task_paths.validation_report)
        _apply_task_fidelity_checks(report, output_file, task_spec)
        report["summary"]["error_count"] = len(report.get("errors", []))
        report["summary"]["warning_count"] = len(report.get("warnings", []))
        report["status"] = (
            "fail"
            if report["errors"]
            else "warn"
            if report["warnings"]
            else "pass"
        )
        task_paths.validation_report.write_text(
            json.dumps(report, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        result = ValidationResult(
            status=str(report.get("status", "error")),
            issues=list(report.get("errors", [])),
            warnings=list(report.get("warnings", [])),
            suggestions=[str(item) for item in report.get("suggestions", [])],
            report_file=str(task_paths.validation_report),
            summary=dict(report.get("summary", {})),
        )
        append_run_log_event(
            task_paths,
            event="validation_completed",
            status=result.status,
            details={
                "report_file": result.report_file,
                "error_count": len(result.issues),
                "warning_count": len(result.warnings),
                "task_type": task_spec.task_type,
            },
        )
        return result
    except Exception as exc:
        error_item = {
            "check": "validation_service",
            "message": f"{type(exc).__name__}: {exc}",
        }
        fallback_report = {
            "status": "error",
            "file": str(output_file),
            "summary": {},
            "errors": [error_item],
            "warnings": [],
            "suggestions": ["检查生成文件和本地 Python 依赖后重新运行校验。"],
        }
        task_paths.validation_report.write_text(
            json.dumps(fallback_report, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        result = ValidationResult(
            status="error",
            issues=[error_item],
            warnings=[],
            suggestions=fallback_report["suggestions"],
            report_file=str(task_paths.validation_report),
            summary={},
        )
        append_run_log_event(
            task_paths,
            event="validation_failed",
            status="error",
            details=result.to_dict(),
        )
        return result


def _apply_task_fidelity_checks(
    report: dict[str, Any],
    output_file: str | Path,
    task_spec: TaskSpec,
) -> None:
    agent_plan = task_spec.options.get("agent_blueprint")
    content_plan = task_spec.options.get("content_plan", {})
    if isinstance(agent_plan, dict) and agent_plan.get("columns"):
        plan = {
            "title": agent_plan.get("title"),
            "columns": [
                {"name": item.get("label")}
                for item in agent_plan.get("columns", [])
                if isinstance(item, dict)
            ],
            "expected_data_rows": len(agent_plan.get("records", [])),
            "explicit_structure": True,
        }
    else:
        plan = content_plan
    if not isinstance(plan, dict) or not plan.get("explicit_structure"):
        return
    try:
        wb = load_workbook(output_file, data_only=False)
    except Exception:
        return
    expected_columns = [
        str(item.get("name", "")).strip()
        for item in plan.get("columns", [])
        if str(item.get("name", "")).strip()
    ]
    expected_title = str(plan.get("title", "")).strip()
    target = None
    detected_headers: list[str] = []
    header_row = None
    for ws in wb.worksheets:
        for row in range(1, min(ws.max_row, 12) + 1):
            values = [
                str(ws.cell(row, column).value).strip()
                for column in range(1, ws.max_column + 1)
                if ws.cell(row, column).value not in (None, "")
            ]
            if expected_columns and len(set(values) & set(expected_columns)) >= max(
                1, min(2, len(expected_columns))
            ):
                target = ws
                detected_headers = values
                header_row = row
                break
        if target is not None:
            break

    if target is None:
        _fidelity_error(
            report,
            "requirement_columns",
            "生成结果中没有找到用户要求的表头。",
        )
        return

    missing = [name for name in expected_columns if name not in detected_headers]
    if missing:
        _fidelity_error(
            report,
            "requirement_columns",
            f"生成结果缺少用户要求的列：{'、'.join(missing)}。",
        )

    if expected_title:
        visible_titles = [
            str(ws["A1"].value or "").strip() for ws in wb.worksheets if ws.sheet_state == "visible"
        ]
        if not any(expected_title in value or value in expected_title for value in visible_titles if value):
            _fidelity_warning(
                report,
                "requirement_title",
                f"工作簿标题与需求标题“{expected_title}”不一致。",
            )

    expected_rows = int(plan.get("expected_data_rows") or 0)
    if expected_rows and header_row:
        actual_rows = sum(
            1
            for row in range(header_row + 1, target.max_row + 1)
            if any(target.cell(row, col).value not in (None, "") for col in range(1, target.max_column + 1))
        )
        if actual_rows < expected_rows:
            _fidelity_error(
                report,
                "requirement_rows",
                f"需求中识别到 {expected_rows} 条数据，但表格只有 {actual_rows} 条。",
            )

    if task_spec.include_charts and sum(len(ws._charts) for ws in wb.worksheets) == 0:
        _fidelity_warning(report, "requirement_chart", "用户要求图表，但工作簿中未找到图表。")

    report["summary"]["requirement_fidelity_checked"] = True
    report["summary"]["expected_column_count"] = len(expected_columns)


def _fidelity_error(report: dict[str, Any], check: str, message: str) -> None:
    report["errors"].append({"check": check, "message": message})
    report["status"] = "fail"
    if "请根据检查结果修改后重新生成。" not in report["suggestions"]:
        report["suggestions"].append("请根据检查结果修改后重新生成。")


def _fidelity_warning(report: dict[str, Any], check: str, message: str) -> None:
    report["warnings"].append({"check": check, "message": message})
    if report.get("status") == "pass":
        report["status"] = "warn"
