"""Adapter around the existing deterministic workbook validator."""

from __future__ import annotations

import json
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Any

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
