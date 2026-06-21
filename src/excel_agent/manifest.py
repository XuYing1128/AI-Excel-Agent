"""Local JSON manifest for completed and failed V1 tasks."""

from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path
from typing import Any

from .task_paths import outputs_root


MANIFEST_VERSION = 1


def manifest_path(path: str | Path | None = None) -> Path:
    target = Path(path) if path is not None else outputs_root() / "manifest.json"
    target.parent.mkdir(parents=True, exist_ok=True)
    return target


def empty_manifest() -> dict[str, Any]:
    return {"version": MANIFEST_VERSION, "tasks": []}


def load_manifest(path: str | Path | None = None) -> dict[str, Any]:
    target = manifest_path(path)
    if not target.exists():
        return empty_manifest()
    try:
        data = json.loads(target.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return empty_manifest()
    if not isinstance(data, dict) or not isinstance(data.get("tasks"), list):
        return empty_manifest()
    data.setdefault("version", MANIFEST_VERSION)
    return data


def save_manifest(data: dict[str, Any], path: str | Path | None = None) -> Path:
    target = manifest_path(path)
    temp_path = target.with_suffix(target.suffix + ".tmp")
    temp_path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    temp_path.replace(target)
    return target


def append_manifest_record(record: dict[str, Any], path: str | Path | None = None) -> Path:
    required = {
        "task_id",
        "created_at",
        "task_type",
        "user_prompt",
        "input_files",
        "output_file",
        "validation_report",
        "status",
        "error",
    }
    missing = sorted(required - set(record))
    if missing:
        raise ValueError(f"manifest 记录缺少字段: {missing}")
    data = load_manifest(path)
    tasks = [item for item in data["tasks"] if item.get("task_id") != record["task_id"]]
    tasks.append(record)
    data["tasks"] = tasks
    return save_manifest(data, path)


def update_manifest_record(
    task_id: str,
    updates: dict[str, Any],
    path: str | Path | None = None,
) -> Path:
    data = load_manifest(path)
    for item in data["tasks"]:
        if item.get("task_id") == task_id:
            item.update(updates)
            item["updated_at"] = datetime.now().astimezone().isoformat(timespec="seconds")
            return save_manifest(data, path)
    raise KeyError(f"manifest 中不存在 task_id: {task_id}")


def recent_tasks(limit: int = 10, path: str | Path | None = None) -> list[dict[str, Any]]:
    tasks = load_manifest(path)["tasks"]
    ordered = sorted(tasks, key=lambda item: str(item.get("created_at", "")), reverse=True)
    return ordered[: max(0, limit)]


def build_manifest_record(
    *,
    task_id: str,
    task_type: str,
    user_prompt: str,
    input_files: list[str],
    output_file: str | None,
    validation_report: str | None,
    status: str,
    error: str | None,
) -> dict[str, Any]:
    return {
        "task_id": task_id,
        "created_at": datetime.now().astimezone().isoformat(timespec="seconds"),
        "task_type": task_type,
        "user_prompt": user_prompt,
        "input_files": list(input_files),
        "output_file": output_file,
        "validation_report": validation_report,
        "status": status,
        "error": error,
    }
