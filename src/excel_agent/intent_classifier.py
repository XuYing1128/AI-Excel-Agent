"""Rule-based table intent classifier for Chinese and English prompts."""

from __future__ import annotations

from dataclasses import dataclass


SUPPORTED_TYPES = [
    "personal_budget",
    "family_budget",
    "quotation",
    "invoice_draft",
    "inventory",
    "sales_report",
    "ecommerce_analysis",
    "project_plan",
    "schedule",
    "attendance",
    "finance_model",
    "dashboard",
    "generic_table",
]


KEYWORDS: dict[str, list[str]] = {
    "personal_budget": ["预算", "收支", "家庭", "个人", "生活费", "budget", "expense"],
    "family_budget": ["家庭年度", "年度预算", "家庭预算", "家庭", "年度", "family budget"],
    "quotation": ["报价", "采购报价", "采购单", "合同", "quote", "quotation", "po"],
    "invoice_draft": ["发票", "发票草稿", "收款", "收款明细", "开票", "invoice"],
    "inventory": ["库存", "进销存", "入库", "出库", "补货", "sku", "inventory", "stock"],
    "sales_report": ["销售", "月报", "日报", "销售额", "sales", "revenue"],
    "ecommerce_analysis": ["电商", "订单", "gmv", "sku", "退款", "店铺", "转化", "ecommerce"],
    "project_plan": ["项目", "计划", "里程碑", "甘特", "gantt", "project", "milestone"],
    "schedule": ["课程表", "排班", "值班", "课表", "schedule", "shift", "timetable"],
    "attendance": ["考勤", "签到", "出勤", "迟到", "attendance", "checkin"],
    "finance_model": ["财务", "利润", "测算", "毛利", "成本", "roi", "pnl", "finance", "model"],
    "dashboard": ["仪表盘", "看板", "dashboard", "bi", "kpi", "经营分析"],
}


ALIASES = {
    "budget": "personal_budget",
    "personal": "personal_budget",
    "family": "family_budget",
    "quote": "quotation",
    "invoice": "invoice_draft",
    "stock": "inventory",
    "sales": "sales_report",
    "ecommerce": "ecommerce_analysis",
    "project": "project_plan",
    "shift": "schedule",
    "finance": "finance_model",
    "model": "finance_model",
}


@dataclass(frozen=True)
class IntentResult:
    table_type: str
    confidence: float
    matched_keywords: list[str]


def normalize_table_type(value: str | None) -> str:
    if not value:
        return "generic_table"
    raw = value.strip().lower()
    if raw in SUPPORTED_TYPES:
        return raw
    return ALIASES.get(raw, raw if raw in SUPPORTED_TYPES else "generic_table")


def classify_intent(text: str) -> IntentResult:
    lowered = text.lower()
    scores: dict[str, list[str]] = {}
    for table_type, words in KEYWORDS.items():
        matched = [word for word in words if word.lower() in lowered]
        if matched:
            scores[table_type] = matched
    if not scores:
        return IntentResult("generic_table", 0.2, [])
    table_type, matched = max(scores.items(), key=lambda item: (len(item[1]), len("".join(item[1]))))
    confidence = min(0.95, 0.35 + len(matched) * 0.18)
    return IntentResult(table_type, confidence, matched)


def classify_intent_with_llm(text: str, llm_client: object | None = None) -> IntentResult:
    """Reserved extension point for a future LLM-backed classifier.

    The MVP stays fully local and deterministic. A future caller can pass an
    LLM client here, but this function deliberately falls back to rules today.
    """

    return classify_intent(text)
