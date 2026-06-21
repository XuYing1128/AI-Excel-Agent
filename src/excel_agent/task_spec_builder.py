"""Conservative one-round TaskSpec drafting and clarification rules."""

from __future__ import annotations

from copy import deepcopy
from pathlib import Path
from typing import Any

from .intent_classifier import KEYWORDS, SUPPORTED_TYPES, classify_intent, normalize_table_type
from .task_spec import TaskSpec, TaskSpecDraft


VAGUE_PROMPTS = {
    "帮我整理一下",
    "整理一下",
    "做个表格",
    "帮我做个表格",
    "处理一下",
    "帮我处理一下",
    "看看这个文件",
}
DATA_REQUEST_WORDS = ["根据", "基于", "分析", "清洗", "这个文件", "这份数据", "订单数据", "原始数据"]
CHART_TYPES = {"sales_report", "ecommerce_analysis", "dashboard"}
TYPE_LABELS = {
    "personal_budget": "个人月度预算",
    "family_budget": "家庭年度预算",
    "quotation": "报价单",
    "invoice_draft": "发票草稿/收款明细",
    "inventory": "库存进销存",
    "sales_report": "销售报表",
    "ecommerce_analysis": "电商订单分析",
    "project_plan": "项目计划",
    "schedule": "课程/排班表",
    "attendance": "考勤统计",
    "finance_model": "财务测算",
    "dashboard": "综合仪表盘",
    "generic_table": "通用表格",
}


def build_task_spec_draft(user_prompt: str, input_files: list[str]) -> TaskSpecDraft:
    prompt = str(user_prompt or "").strip()
    files = [str(path) for path in input_files]
    classification = classify_intent(prompt)
    alternatives = _classification_alternatives(prompt)
    assumptions = _default_assumptions(classification.table_type, prompt, files)
    task_spec = TaskSpec(
        task_type=classification.table_type,
        user_goal=prompt or "根据上传文件生成合适的本地电子表格",
        input_files=files,
        output_name="生成结果.xlsx",
        preserve_template_style=bool(
            files
            and any(Path(path).suffix.lower() in {".xlsx", ".xlsm"} for path in files)
            and any(word in prompt for word in ["保留", "参考模板", "原样", "样式"])
        ),
        include_charts=classification.table_type in CHART_TYPES or any(
            word in prompt.lower() for word in ["图表", "dashboard", "看板"]
        ),
        include_summary=True,
        include_instructions_sheet=True,
        confidence=classification.confidence,
        assumptions=assumptions,
        user_answers={},
        options={
            "classification_keywords": classification.matched_keywords,
            "generation_policy": (
                "sales_input_analysis"
                if classification.table_type == "sales_report" and files
                else "standard_template"
            ),
            "model_may_edit_cells": False,
            "deterministic_validation_required": True,
        },
    )
    questions = _build_questions(task_spec, prompt, files, alternatives)
    return TaskSpecDraft(
        task_spec=task_spec,
        clarifying_questions=questions[:5],
        classification_alternatives=alternatives,
    )


def get_clarifying_questions(task_spec_or_context: TaskSpecDraft | TaskSpec | dict[str, Any]) -> list[str]:
    if isinstance(task_spec_or_context, TaskSpecDraft):
        return list(task_spec_or_context.clarifying_questions)
    if isinstance(task_spec_or_context, TaskSpec):
        return _build_questions(
            task_spec_or_context,
            task_spec_or_context.user_goal,
            task_spec_or_context.input_files,
            _classification_alternatives(task_spec_or_context.user_goal),
        )[:5]
    prompt = str(task_spec_or_context.get("user_prompt", ""))
    files = [str(path) for path in task_spec_or_context.get("input_files", [])]
    return build_task_spec_draft(prompt, files).clarifying_questions


def merge_user_answers_into_task_spec(task_spec: TaskSpec, answers: dict[str, Any]) -> TaskSpec:
    merged = TaskSpec.from_dict(deepcopy(task_spec.to_dict()))
    clean_answers = {str(key): value for key, value in answers.items()}

    selected_type = normalize_table_type(str(clean_answers.get("task_type", "") or ""))
    if selected_type in SUPPORTED_TYPES and clean_answers.get("task_type"):
        merged.task_type = selected_type
        merged.confidence = max(merged.confidence, 0.9)

    goal_detail = str(
        clean_answers.get("goal_detail")
        or clean_answers.get("extra_requirements")
        or ""
    ).strip()
    if goal_detail:
        merged.user_goal = f"{merged.user_goal}\n补充说明：{goal_detail}".strip()

    output_name = str(clean_answers.get("output_name", "")).strip()
    if output_name:
        merged.output_name = Path(output_name).name
        if not merged.output_name.lower().endswith(".xlsx"):
            merged.output_name += ".xlsx"

    for key in (
        "preserve_template_style",
        "include_charts",
        "include_summary",
        "include_instructions_sheet",
    ):
        if key in clean_answers:
            setattr(merged, key, bool(clean_answers[key]))

    data_mode = str(clean_answers.get("data_mode", "")).strip()
    if data_mode == "template" and not merged.input_files:
        _append_unique(merged.assumptions, "未提供原始数据，当前版本将生成标准模板示例。")
    elif data_mode == "upload" and not merged.input_files:
        _append_unique(merged.assumptions, "用户计划使用原始数据，但确认时尚未提供可用文件。")

    freeform_answers = clean_answers.get("clarifications")
    if isinstance(freeform_answers, dict):
        useful = [str(value).strip() for value in freeform_answers.values() if str(value).strip()]
        if useful:
            merged.user_goal = f"{merged.user_goal}\n补充回答：" + "；".join(useful)

    merged.user_answers.update(clean_answers)
    merged.options["clarification_rounds"] = 1
    merged.options["generation_policy"] = (
        "sales_input_analysis"
        if merged.task_type == "sales_report" and merged.input_files
        else "standard_template"
    )
    return merged


def _build_questions(
    task_spec: TaskSpec,
    prompt: str,
    input_files: list[str],
    alternatives: list[str],
) -> list[str]:
    questions: list[str] = []
    normalized_prompt = prompt.strip()
    if task_spec.confidence < 0.6:
        questions.append("系统暂时不能确定表格类型，请确认最接近的表格类型。")
    if not normalized_prompt and input_files:
        questions.append("你希望对上传文件做什么：清洗、分析、生成报表，还是仅套用标准模板？")
    if normalized_prompt in VAGUE_PROMPTS or len(normalized_prompt) < 5:
        questions.append("请补充这张表要解决的具体问题，以及你最希望看到的汇总结果。")
    if _needs_input_file(normalized_prompt) and not input_files:
        questions.append("这个需求看起来依赖原始数据；请上传 CSV/XLSX/XLSM，或确认先生成标准模板。")
    if {"sales_report", "ecommerce_analysis"}.issubset(set(alternatives[:3])):
        questions.append("需求同时像销售报表和电商订单分析，请确认更偏向哪一种。")
    return list(dict.fromkeys(questions))


def _classification_alternatives(prompt: str) -> list[str]:
    lowered = prompt.lower()
    scored: list[tuple[int, int, str]] = []
    for table_type, words in KEYWORDS.items():
        matched = [word for word in words if word.lower() in lowered]
        if matched:
            scored.append((len(matched), len("".join(matched)), table_type))
    scored.sort(reverse=True)
    return [table_type for _, _, table_type in scored]


def _needs_input_file(prompt: str) -> bool:
    lowered = prompt.lower()
    return any(word.lower() in lowered for word in DATA_REQUEST_WORDS)


def _default_assumptions(task_type: str, prompt: str, input_files: list[str]) -> list[str]:
    assumptions = [
        "所有具体单元格、公式和格式均由确定性 Python 内核处理，大模型不直接修改单元格。",
        "生成后必须运行确定性校验器；主观模型审查不会替代客观校验。",
    ]
    if not input_files:
        assumptions.append("当前没有输入文件，将使用标准模板中的示例数据生成示例表格。")
    if task_type != "sales_report":
        assumptions.append(
            "当前版本对该类型主要生成标准模板，不保证自动完成复杂真实数据分析。"
        )
    if task_type in {"quotation", "invoice_draft", "finance_model", "attendance"}:
        assumptions.append("该任务包含高风险业务结果，必须由人工复核。")
    if "不需要图表" in prompt:
        assumptions.append("按用户描述默认不添加图表。")
    return assumptions


def _append_unique(values: list[str], value: str) -> None:
    if value not in values:
        values.append(value)
