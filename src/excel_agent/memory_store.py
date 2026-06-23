"""本地 SQLite 记忆层。

只保存低风险偏好、历史任务索引和 skill 版本；不保存 API key，不上传。
"""

from __future__ import annotations

import json
import os
import sqlite3
from datetime import datetime
from pathlib import Path
from typing import Any

from .io_utils import project_root
from .task_spec import TaskSpec


def memory_db_path(path: str | Path | None = None) -> Path:
    override = os.getenv("AI_EXCEL_MEMORY_DB", "").strip()
    if path is not None:
        target = Path(path)
    elif override:
        target = Path(override)
    else:
        target = project_root() / "data" / "private" / "memory.db"
    return target.expanduser().resolve()


def connect_memory(path: str | Path | None = None) -> sqlite3.Connection:
    target = memory_db_path(path)
    target.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(target)
    conn.row_factory = sqlite3.Row
    init_memory_store(conn)
    return conn


def init_memory_store(conn: sqlite3.Connection) -> None:
    conn.executescript(
        """
        CREATE TABLE IF NOT EXISTS preferences (
            key TEXT PRIMARY KEY,
            value TEXT NOT NULL,
            updated_at TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS task_history (
            task_id TEXT PRIMARY KEY,
            prompt TEXT NOT NULL,
            task_type TEXT NOT NULL,
            output_file TEXT,
            status TEXT NOT NULL,
            created_at TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS skills (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            version INTEGER NOT NULL,
            content TEXT NOT NULL,
            enabled INTEGER NOT NULL DEFAULT 1,
            created_at TEXT NOT NULL,
            UNIQUE(name, version)
        );
        """
    )
    conn.commit()


def set_preference(key: str, value: Any, path: str | Path | None = None) -> None:
    now = _now()
    with connect_memory(path) as conn:
        conn.execute(
            """
            INSERT INTO preferences(key, value, updated_at)
            VALUES (?, ?, ?)
            ON CONFLICT(key) DO UPDATE SET value=excluded.value, updated_at=excluded.updated_at
            """,
            (key, json.dumps(value, ensure_ascii=False), now),
        )
        conn.commit()


def get_preference(key: str, default: Any = None, path: str | Path | None = None) -> Any:
    with connect_memory(path) as conn:
        row = conn.execute("SELECT value FROM preferences WHERE key=?", (key,)).fetchone()
    if row is None:
        return default
    try:
        return json.loads(row["value"])
    except json.JSONDecodeError:
        return row["value"]


def list_preferences(path: str | Path | None = None) -> dict[str, Any]:
    with connect_memory(path) as conn:
        rows = conn.execute("SELECT key, value FROM preferences ORDER BY key").fetchall()
    result: dict[str, Any] = {}
    for row in rows:
        try:
            result[row["key"]] = json.loads(row["value"])
        except json.JSONDecodeError:
            result[row["key"]] = row["value"]
    return result


def clear_preferences(path: str | Path | None = None) -> None:
    with connect_memory(path) as conn:
        conn.execute("DELETE FROM preferences")
        conn.commit()


def record_task_history(
    *,
    task_id: str,
    prompt: str,
    task_type: str,
    output_file: str | None,
    status: str,
    path: str | Path | None = None,
) -> None:
    with connect_memory(path) as conn:
        conn.execute(
            """
            INSERT INTO task_history(task_id, prompt, task_type, output_file, status, created_at)
            VALUES (?, ?, ?, ?, ?, ?)
            ON CONFLICT(task_id) DO UPDATE SET
                prompt=excluded.prompt,
                task_type=excluded.task_type,
                output_file=excluded.output_file,
                status=excluded.status
            """,
            (task_id, prompt, task_type, output_file, status, _now()),
        )
        conn.commit()


def list_task_history(limit: int = 30, path: str | Path | None = None) -> list[dict[str, Any]]:
    with connect_memory(path) as conn:
        rows = conn.execute(
            """
            SELECT task_id, prompt, task_type, output_file, status, created_at
            FROM task_history
            ORDER BY created_at DESC
            LIMIT ?
            """,
            (limit,),
        ).fetchall()
    return [dict(row) for row in rows]


def save_skill_version(
    name: str,
    content: str,
    *,
    enabled: bool = True,
    path: str | Path | None = None,
) -> int:
    with connect_memory(path) as conn:
        row = conn.execute("SELECT MAX(version) AS version FROM skills WHERE name=?", (name,)).fetchone()
        version = int(row["version"] or 0) + 1
        conn.execute(
            "INSERT INTO skills(name, version, content, enabled, created_at) VALUES (?, ?, ?, ?, ?)",
            (name, version, content, 1 if enabled else 0, _now()),
        )
        conn.commit()
    return version


def list_skill_versions(name: str | None = None, path: str | Path | None = None) -> list[dict[str, Any]]:
    sql = "SELECT name, version, content, enabled, created_at FROM skills"
    params: tuple[Any, ...] = ()
    if name:
        sql += " WHERE name=?"
        params = (name,)
    sql += " ORDER BY name, version DESC"
    with connect_memory(path) as conn:
        rows = conn.execute(sql, params).fetchall()
    return [dict(row) for row in rows]


def rollback_skill_version(name: str, version: int, path: str | Path | None = None) -> str:
    with connect_memory(path) as conn:
        row = conn.execute(
            "SELECT content FROM skills WHERE name=? AND version=?",
            (name, int(version)),
        ).fetchone()
        if row is None:
            raise ValueError(f"skill 版本不存在: {name} v{version}")
        conn.execute("UPDATE skills SET enabled=0 WHERE name=?", (name,))
        conn.execute(
            "UPDATE skills SET enabled=1 WHERE name=? AND version=?",
            (name, int(version)),
        )
        conn.commit()
        return str(row["content"])


def learn_preferences_from_task(
    task_spec: TaskSpec,
    workbook_summary: dict[str, Any] | None = None,
    *,
    path: str | Path | None = None,
) -> dict[str, Any]:
    """Conservatively learn only low-risk UI/style preferences."""

    learned: dict[str, Any] = {}
    if task_spec.include_charts:
        learned["prefer_charts_when_requested"] = True
    if task_spec.include_instructions_sheet is False:
        learned["prefer_compact_workbook"] = True
    if task_spec.preserve_template_style:
        learned["prefer_template_style_when_uploaded"] = True
    sheet_names = [
        str(item.get("name", ""))
        for item in (workbook_summary or {}).get("sheets", [])
        if isinstance(item, dict)
    ]
    if sheet_names and all(_has_cjk(name) for name in sheet_names if name):
        learned["sheet_name_language"] = "中文"
    for key, value in learned.items():
        set_preference(key, value, path)
    return learned


def _has_cjk(value: str) -> bool:
    return any("\u4e00" <= char <= "\u9fff" for char in value)


def _now() -> str:
    return datetime.now().astimezone().isoformat(timespec="seconds")

