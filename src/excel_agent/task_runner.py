"""Natural-language task file runner."""

from __future__ import annotations

import re
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Any

from .intent_classifier import IntentResult, classify_intent
from .io_utils import project_root, read_table
from .validators import validate_workbook
from .workbook_builder import analyze_sales_file, create_workbook


INPUT_PATH_RE = re.compile(
    r"(?P<path>(?:[A-Za-z]:[^\s，。；;、'\"“”<>]+|[^\s，。；;、'\"“”<>：:]+)\.(?:csv|tsv|xlsx|xlsm|xls))",
    re.IGNORECASE,
)
ANALYZABLE_INPUT_TYPES = {"sales_report"}


@dataclass
class TaskRunResult:
    task_file: str
    output: str
    table_type: str
    confidence: float
    matched_keywords: list[str]
    input_file: str | None
    used_input_file: bool
    mode: str
    validation_status: str
    validation_errors: int
    validation_warnings: int
    message: str


def run_task_file(task_path: str | Path, output_path: str | Path) -> dict[str, Any]:
    task_file = _resolve_existing_path(task_path)
    task_text = _read_task_text(task_file)
    intent = classify_intent(task_text)
    input_file = _extract_existing_input_path(task_text, task_file.parent)

    output = _dispatch_task(intent, input_file, output_path)
    validation = validate_workbook(output)
    result = TaskRunResult(
        task_file=str(task_file),
        output=str(output),
        table_type=intent.table_type,
        confidence=intent.confidence,
        matched_keywords=intent.matched_keywords,
        input_file=str(input_file) if input_file else None,
        used_input_file=input_file is not None and intent.table_type in ANALYZABLE_INPUT_TYPES,
        mode=_mode_for(intent.table_type, input_file),
        validation_status=validation["status"],
        validation_errors=len(validation["errors"]),
        validation_warnings=len(validation["warnings"]),
        message=_message_for(intent, input_file, validation),
    )
    return {"task": asdict(result), "validation": validation}


def _dispatch_task(intent: IntentResult, input_file: Path | None, output_path: str | Path) -> Path:
    if input_file and intent.table_type == "sales_report":
        # Existing analyzer builds a real report from sales-like CSV/XLSX data.
        read_table(input_file)
        return analyze_sales_file(input_file, output_path)

    if input_file:
        # For non-sales MVP templates, verify the file is readable but keep the
        # generated workbook template-driven until a type-specific analyzer exists.
        read_table(input_file)

    return create_workbook(intent.table_type, output_path)


def _read_task_text(path: Path) -> str:
    for encoding in ("utf-8-sig", "utf-8", "gb18030"):
        try:
            return path.read_text(encoding=encoding)
        except UnicodeDecodeError:
            continue
    return path.read_text(encoding="utf-8", errors="replace")


def _extract_existing_input_path(task_text: str, task_dir: Path) -> Path | None:
    candidates = []
    for match in INPUT_PATH_RE.finditer(task_text):
        raw = match.group("path").strip().strip("`'\"“”")
        candidates.extend(_candidate_paths(raw, task_dir))
    for candidate in candidates:
        if candidate.exists() and candidate.is_file():
            return candidate
    return None


def _candidate_paths(raw: str, task_dir: Path) -> list[Path]:
    path = Path(raw)
    if path.is_absolute():
        return [path]
    root = project_root()
    return [
        (Path.cwd() / path).resolve(),
        (root / path).resolve(),
        (task_dir / path).resolve(),
    ]


def _resolve_existing_path(path: str | Path) -> Path:
    p = Path(path)
    if not p.is_absolute():
        p = (project_root() / p).resolve()
    if not p.exists():
        raise FileNotFoundError(f"任务文件不存在: {p}")
    return p


def _mode_for(table_type: str, input_file: Path | None) -> str:
    if input_file and table_type in ANALYZABLE_INPUT_TYPES:
        return "input_analyzer"
    if input_file:
        return "template_with_input_read_check"
    return "template_demo"


def _message_for(intent: IntentResult, input_file: Path | None, validation: dict[str, Any]) -> str:
    if input_file and intent.table_type in ANALYZABLE_INPUT_TYPES:
        source = f"已读取输入文件 {input_file} 并生成 {intent.table_type} 报表。"
    elif input_file:
        source = f"已确认输入文件 {input_file} 可读取；当前类型使用模板 demo 生成。"
    else:
        source = "任务未提供可识别的本地输入文件，已使用示例数据生成 demo。"
    return f"{source} 校验状态：{validation['status']}。"
