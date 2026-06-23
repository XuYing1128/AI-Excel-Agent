"""Agent tool abstractions."""

from __future__ import annotations

from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Any, Callable

from ....task_paths import TaskPaths
from ....task_spec import TaskSpec


@dataclass
class ToolResult:
    ok: bool
    summary: str
    data: dict[str, Any] = field(default_factory=dict)
    artifacts: list[str] = field(default_factory=list)
    error: str | None = None

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass
class ToolContext:
    task_spec: TaskSpec
    task_paths: TaskPaths
    progress: Callable[[str, str], None] | None = None
    run_python_enabled: bool = True

    @property
    def temp_dir(self) -> Path:
        target = self.task_paths.task_dir / "agent_tmp"
        target.mkdir(parents=True, exist_ok=True)
        return target


@dataclass
class AgentTool:
    name: str
    description: str
    schema: dict[str, Any]
    handler: Callable[[dict[str, Any], ToolContext], ToolResult]

    def openai_schema(self) -> dict[str, Any]:
        return {
            "type": "function",
            "function": {
                "name": self.name,
                "description": self.description,
                "parameters": self.schema,
            },
        }

