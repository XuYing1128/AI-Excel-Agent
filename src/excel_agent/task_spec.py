"""Structured task contract confirmed before workbook generation."""

from __future__ import annotations

import json
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Any

from .intent_classifier import SUPPORTED_TYPES


@dataclass
class TaskSpec:
    task_type: str
    user_goal: str
    input_files: list[str] = field(default_factory=list)
    output_name: str = "生成结果.xlsx"
    preserve_template_style: bool = False
    include_charts: bool = False
    include_summary: bool = True
    include_instructions_sheet: bool = True
    confidence: float = 0.0
    assumptions: list[str] = field(default_factory=list)
    user_answers: dict[str, Any] = field(default_factory=dict)
    options: dict[str, Any] = field(default_factory=dict)

    def __post_init__(self) -> None:
        if self.task_type not in SUPPORTED_TYPES:
            raise ValueError(f"不支持的 task_type: {self.task_type}")
        self.user_goal = str(self.user_goal).strip()
        self.input_files = [str(path) for path in self.input_files]
        self.output_name = Path(str(self.output_name or "生成结果.xlsx")).name
        if not self.output_name.lower().endswith(".xlsx"):
            self.output_name = f"{self.output_name}.xlsx"
        self.confidence = max(0.0, min(1.0, float(self.confidence)))
        self.assumptions = [str(item).strip() for item in self.assumptions if str(item).strip()]
        self.user_answers = dict(self.user_answers or {})
        self.options = dict(self.options or {})

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> "TaskSpec":
        return cls(**data)


@dataclass
class TaskSpecDraft:
    task_spec: TaskSpec
    clarifying_questions: list[str] = field(default_factory=list)
    classification_alternatives: list[str] = field(default_factory=list)

    @property
    def needs_clarification(self) -> bool:
        return bool(self.clarifying_questions)


def save_task_spec(task_spec: TaskSpec, path: str | Path) -> Path:
    target = Path(path)
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_text(
        json.dumps(task_spec.to_dict(), ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    return target


def load_task_spec(path: str | Path) -> TaskSpec:
    data = json.loads(Path(path).read_text(encoding="utf-8"))
    if not isinstance(data, dict):
        raise ValueError("task_spec.json 顶层必须是对象。")
    return TaskSpec.from_dict(data)
