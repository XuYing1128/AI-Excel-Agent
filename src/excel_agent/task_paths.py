"""Task-scoped filesystem paths for the local V1 web workflow."""

from __future__ import annotations

import json
import os
import re
import secrets
import shutil
from dataclasses import asdict, dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Iterable

from .intent_classifier import SUPPORTED_TYPES, normalize_table_type
from .io_utils import project_root


INVALID_FILENAME_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')


@dataclass(frozen=True)
class TaskPaths:
    task_id: str
    task_dir: Path
    input_dir: Path
    output_dir: Path
    reports_dir: Path
    task_spec_file: Path
    run_log_file: Path
    output_file: Path
    validation_report: Path
    subjective_review_report: Path

    def to_dict(self) -> dict[str, str]:
        return {key: str(value) for key, value in asdict(self).items()}


def outputs_root() -> Path:
    override = os.getenv("AI_EXCEL_OUTPUTS_DIR", "").strip()
    return Path(override).expanduser().resolve() if override else project_root() / "outputs"


def tasks_root(base_dir: str | Path | None = None) -> Path:
    root = Path(base_dir) if base_dir is not None else outputs_root() / "tasks"
    root.mkdir(parents=True, exist_ok=True)
    return root


def create_task_paths(
    task_type: str,
    base_dir: str | Path | None = None,
    now: datetime | None = None,
) -> TaskPaths:
    resolved_type = normalize_table_type(task_type)
    if resolved_type not in SUPPORTED_TYPES:
        resolved_type = "generic_table"
    timestamp = (now or datetime.now()).strftime("%Y%m%d_%H%M%S")
    base_task_id = f"{timestamp}_{resolved_type}"
    root = tasks_root(base_dir)
    task_id = base_task_id
    task_dir = root / task_id
    while task_dir.exists():
        task_id = f"{base_task_id}_{secrets.token_hex(2)}"
        task_dir = root / task_id

    input_dir = task_dir / "input"
    output_dir = task_dir / "output"
    reports_dir = task_dir / "reports"
    for directory in (input_dir, output_dir, reports_dir):
        directory.mkdir(parents=True, exist_ok=False)

    return TaskPaths(
        task_id=task_id,
        task_dir=task_dir,
        input_dir=input_dir,
        output_dir=output_dir,
        reports_dir=reports_dir,
        task_spec_file=task_dir / "task_spec.json",
        run_log_file=task_dir / "run_log.json",
        output_file=output_dir / "result.xlsx",
        validation_report=reports_dir / "validation.json",
        subjective_review_report=reports_dir / "subjective_review.json",
    )


def sanitize_filename(filename: str, fallback: str = "input") -> str:
    name = Path(str(filename)).name.strip()
    name = INVALID_FILENAME_CHARS.sub("_", name).rstrip(". ")
    return name or fallback


def unique_destination(directory: Path, filename: str) -> Path:
    safe_name = sanitize_filename(filename)
    candidate = directory / safe_name
    counter = 1
    while candidate.exists():
        candidate = directory / f"{Path(safe_name).stem}_{counter}{Path(safe_name).suffix}"
        counter += 1
    return candidate


def stage_input_files(paths: Iterable[str | Path], task_paths: TaskPaths) -> list[str]:
    staged: list[str] = []
    input_root = task_paths.input_dir.resolve()
    for raw_path in paths:
        source = Path(raw_path).expanduser()
        if not source.exists() or not source.is_file():
            raise FileNotFoundError(f"输入文件不存在: {source}")
        resolved = source.resolve()
        if resolved.parent == input_root:
            staged.append(str(resolved))
            continue
        destination = unique_destination(task_paths.input_dir, source.name)
        shutil.copy2(resolved, destination)
        staged.append(str(destination.resolve()))
    return staged


def save_uploaded_bytes(filename: str, data: bytes, task_paths: TaskPaths) -> Path:
    destination = unique_destination(task_paths.input_dir, filename)
    destination.write_bytes(data)
    return destination.resolve()


def append_run_log_event(
    task_paths: TaskPaths,
    event: str,
    status: str,
    details: dict[str, Any] | None = None,
) -> dict[str, Any]:
    payload: dict[str, Any] = {
        "task_id": task_paths.task_id,
        "events": [],
    }
    if task_paths.run_log_file.exists():
        try:
            loaded = json.loads(task_paths.run_log_file.read_text(encoding="utf-8"))
            if isinstance(loaded, dict):
                payload.update(loaded)
            if not isinstance(payload.get("events"), list):
                payload["events"] = []
        except (json.JSONDecodeError, OSError):
            payload["events"] = []
    payload["events"].append(
        {
            "timestamp": datetime.now().astimezone().isoformat(timespec="seconds"),
            "event": event,
            "status": status,
            "details": details or {},
        }
    )
    payload["latest_event_status"] = status
    if status not in {"running", "skipped"}:
        payload["latest_status"] = status
    task_paths.run_log_file.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    return payload
