"""Project skill matcher for the local Excel agent."""

from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any


SKILLS_ROOT = Path(__file__).resolve().parent


@dataclass
class Skill:
    name: str
    path: Path
    content: str
    triggers: list[str]
    tools: list[str]


SKILL_RULES: dict[str, list[str]] = {
    "template_fill": ["模板", "套用", "严格", "导入", "参考格式", "按模板"],
    "schedule": ["排考", "考试", "考场", "教室", "座位", "时间冲突", "场次"],
    "perf_review": ["绩效", "薪酬", "工资", "奖金", "调薪", "考核", "缺勤扣分"],
    "data_clean": ["清洗", "去重", "缺失", "异常", "合并", "标准化", "脏数据"],
    "report_chart": ["图表", "仪表盘", "dashboard", "趋势", "占比", "top", "排名", "报表"],
    "xlsx_pro": ["excel", "xlsx", "csv", "表格", "工作簿"],
}


def match_skills(task: Any) -> list[Skill]:
    text = _task_text(task)
    scored: list[tuple[int, str]] = []
    for name, keywords in SKILL_RULES.items():
        score = sum(_keyword_score(text, item) for item in keywords)
        if score:
            scored.append((score, name))
    scored.sort(key=lambda item: (-item[0], item[1]))
    return [skill for _, name in scored if (skill := load_skill(name)) is not None]


def load_skill(name: str) -> Skill | None:
    path = SKILLS_ROOT / name / "SKILL.md"
    if not path.exists():
        return None
    content = path.read_text(encoding="utf-8")
    meta = _frontmatter(content)
    triggers = _csvish(meta.get("triggers", ""))
    tools = _csvish(meta.get("tools", ""))
    return Skill(name=str(meta.get("name") or name), path=path, content=content, triggers=triggers, tools=tools)


def _task_text(task: Any) -> str:
    if isinstance(task, str):
        return task.lower()
    pieces = [
        str(getattr(task, "task_type", "")),
        str(getattr(task, "user_goal", "")),
        str(getattr(task, "output_name", "")),
        str(getattr(task, "options", "")),
    ]
    return "\n".join(pieces).lower()


def _keyword_score(text: str, keyword: str) -> int:
    word = keyword.lower()
    if re.fullmatch(r"[a-z0-9_]+", word):
        return 2 if word in text else 0
    return text.count(word)


def _frontmatter(content: str) -> dict[str, str]:
    if not content.startswith("---"):
        return {}
    end = content.find("\n---", 3)
    if end < 0:
        return {}
    raw = content[3:end].strip().splitlines()
    meta: dict[str, str] = {}
    for line in raw:
        if ":" not in line:
            continue
        key, value = line.split(":", 1)
        meta[key.strip()] = value.strip().strip('"')
    return meta


def _csvish(value: str) -> list[str]:
    return [item.strip() for item in re.split(r"[,，]", value or "") if item.strip()]

