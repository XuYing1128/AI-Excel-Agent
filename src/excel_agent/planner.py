"""Simple task planner used by the CLI and future Codex runs."""

from __future__ import annotations

from dataclasses import dataclass

from .intent_classifier import classify_intent, normalize_table_type


@dataclass
class WorkPlan:
    table_type: str
    steps: list[str]
    risk_note: str | None = None


HIGH_RISK_TYPES = {"finance_model", "quotation"}


def plan_task(prompt: str | None = None, table_type: str | None = None) -> WorkPlan:
    resolved = normalize_table_type(table_type)
    if resolved == "generic_table" and prompt:
        resolved = classify_intent(prompt).table_type

    steps = [
        f"识别表格类型：{resolved}",
        "准备说明页、数据输入区、计算区和汇总区",
        "把派生值写入 Excel 公式并应用样式、筛选和冻结窗格",
        "保存到 outputs/ 或指定输出路径",
        "运行工作簿静态校验并输出 JSON 结果",
    ]
    note = None
    if resolved in HIGH_RISK_TYPES:
        note = "该类型可能涉及财务、合同或报价决策，结果必须人工复核。"
    return WorkPlan(resolved, steps, note)

