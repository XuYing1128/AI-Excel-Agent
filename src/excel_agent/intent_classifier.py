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
    "attendance": ["考勤统计", "考勤表", "签到", "出勤率", "迟到", "缺勤", "attendance", "checkin"],
    "finance_model": [
        "财务",
        "利润",
        "测算",
        "毛利",
        "成本",
        "薪酬调整",
        "薪资调整",
        "工资",
        "奖金",
        "绩效评估",
        "加权总分",
        "roi",
        "pnl",
        "finance",
        "model",
    ],
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
    scores: dict[str, tuple[float, list[str]]] = {}
    for table_type, words in KEYWORDS.items():
        matched = [word for word in words if word.lower() in lowered]
        if matched:
            score = sum(_keyword_weight(word) for word in matched)
            scores[table_type] = (score, matched)
    if not scores:
        return IntentResult("generic_table", 0.2, [])
    table_type, (score, matched) = max(
        scores.items(),
        key=lambda item: (item[1][0], len(item[1][1]), len("".join(item[1][1]))),
    )
    confidence = min(0.95, 0.35 + score * 0.08)
    return IntentResult(table_type, confidence, matched)


def ranked_intents(text: str) -> list[str]:
    """Return table types ordered by weighted semantic evidence."""

    lowered = str(text or "").lower()
    scored: list[tuple[float, int, str]] = []
    for table_type, words in KEYWORDS.items():
        matched = [word for word in words if word.lower() in lowered]
        if matched:
            scored.append(
                (
                    sum(_keyword_weight(word) for word in matched),
                    len("".join(matched)),
                    table_type,
                )
            )
    scored.sort(reverse=True)
    return [table_type for _, _, table_type in scored]


def _keyword_weight(word: str) -> float:
    normalized = str(word).strip()
    if len(normalized) >= 4:
        return 3.0
    if len(normalized) >= 2:
        return 1.5
    return 1.0


def classify_intent_with_llm(text: str, llm_client: object | None = None) -> IntentResult:
    """Reserved extension point for a future LLM-backed classifier.

    The MVP stays fully local and deterministic. A future caller can pass an
    LLM client here, but this function deliberately falls back to rules today.
    """

    return classify_intent(text)
