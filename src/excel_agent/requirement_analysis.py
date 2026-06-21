"""Assess whether a spreadsheet request is detailed enough to generate reliably."""

from __future__ import annotations

from dataclasses import asdict, dataclass
from typing import Any


CALCULATION_TYPES = {
    "personal_budget",
    "family_budget",
    "quotation",
    "invoice_draft",
    "inventory",
    "sales_report",
    "ecommerce_analysis",
    "attendance",
    "finance_model",
    "dashboard",
}
REPORT_TYPES = {
    "inventory",
    "sales_report",
    "ecommerce_analysis",
    "project_plan",
    "attendance",
    "dashboard",
}


@dataclass(frozen=True)
class RequirementGap:
    key: str
    title: str
    question: str
    example: str
    required: bool = True

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


def analyze_requirement_gaps(
    prompt: str,
    task_type: str,
    input_files: list[str],
    content_plan: dict[str, Any],
    *,
    confidence: float,
    alternatives: list[str] | None = None,
) -> list[RequirementGap]:
    """Return only omissions that could materially change the workbook."""

    text = str(prompt or "").strip()
    lowered = text.lower()
    gaps: list[RequirementGap] = []
    columns = list(content_plan.get("columns") or [])
    records = list(content_plan.get("records") or [])
    formulas = list(content_plan.get("formula_rules") or [])
    alternatives = alternatives or []

    explicit_structure = bool(content_plan.get("explicit_structure"))
    if confidence < 0.6 and task_type == "generic_table" and not explicit_structure:
        gaps.append(
            RequirementGap(
                key="task_type",
                title="用途不明确",
                question="这张表主要用于什么业务场景？",
                example="例如：销售月报、库存进销存、报价单、考勤统计或项目计划。",
            )
        )
    elif {"sales_report", "ecommerce_analysis"}.issubset(set(alternatives[:3])):
        gaps.append(
            RequirementGap(
                key="task_type",
                title="类型需要确认",
                question="需求同时像销售报表和电商订单分析，请确认更偏向哪一种？",
                example="说明主要关注销售员/区域，还是订单、SKU、平台、退款和GMV。",
            )
        )

    depends_on_data = any(
        word in lowered
        for word in (
            "根据",
            "基于",
            "分析",
            "清洗",
            "这个文件",
            "这份数据",
            "原始数据",
            "订单数据",
            "名单",
        )
    )
    if (
        depends_on_data
        and not input_files
        and not records
    ) or (
        task_type in {"sales_report", "ecommerce_analysis", "dashboard"}
        and not input_files
        and not records
        and not explicit_structure
    ):
        gaps.append(
            RequirementGap(
                key="data_source",
                title="缺少数据来源",
                question="表格要使用哪些原始数据？",
                example="请上传 CSV/XLSX，粘贴完整数据，或明确说明先生成空白可填写模板。",
            )
        )

    if task_type == "generic_table" and not columns and not input_files:
        gaps.append(
            RequirementGap(
                key="columns",
                title="缺少表格字段",
                question="明细区必须包含哪些列？请按最终顺序填写。",
                example="例如：日期、订单号、客户、产品、数量、单价、销售额、毛利率、备注。",
            )
        )

    if (
        task_type in CALCULATION_TYPES
        and explicit_structure
        and not formulas
        and not _mentions_calculation(text)
    ):
        gaps.append(
            RequirementGap(
                key="calculations",
                title="缺少计算规则",
                question="哪些结果需要自动计算？请说明计算口径。",
                example="例如：销售额=数量×单价；毛利=销售额-成本；毛利率=毛利/销售额，并处理除零。",
            )
        )

    if (
        task_type in REPORT_TYPES
        and explicit_structure
        and not _mentions_summary_or_sort(text)
    ):
        gaps.append(
            RequirementGap(
                key="summary",
                title="缺少汇总方式",
                question="需要按什么维度汇总、排序或添加小计/总计？",
                example="例如：按区域分组、组内按销售额降序，每个区域后加小计，最后加总计。",
                required=False,
            )
        )

    if task_type in {"quotation", "invoice_draft", "finance_model"} and not _mentions_finance_basis(
        text
    ):
        gaps.append(
            RequirementGap(
                key="finance_basis",
                title="金额口径未说明",
                question="请补充币种、含税/未税、税率或测算周期等金额口径。",
                example="例如：人民币、含13%增值税、按月测算，金额保留2位小数。",
            )
        )

    if task_type in {"schedule", "attendance", "project_plan"} and not _mentions_time_basis(text):
        gaps.append(
            RequirementGap(
                key="time_basis",
                title="时间范围未说明",
                question="请补充日期范围、时间段或计划周期。",
                example="例如：2026年7月1日至7月31日；每天分早班、中班、晚班。",
            )
        )

    if "参考" in text or "完全参照" in text or "保持格式" in text:
        if not input_files:
            gaps.append(
                RequirementGap(
                    key="reference_file",
                    title="缺少参考文件",
                    question="请上传需要参照或保留格式的原始 Excel 文件。",
                    example="只有拿到原文件后才能核对合并单元格、列宽、字体、打印区域和公式位置。",
                )
            )

    return _deduplicate(gaps)[:6]


def questions_from_gaps(gaps: list[RequirementGap]) -> list[str]:
    return [item.question for item in gaps]


def _mentions_calculation(text: str) -> bool:
    return any(
        word in text
        for word in (
            "公式",
            "计算",
            "合计",
            "平均",
            "占比",
            "增长率",
            "毛利",
            "利润",
            "金额",
            "销售额",
            "出勤率",
            "完成率",
            "期末库存",
        )
    )


def _mentions_summary_or_sort(text: str) -> bool:
    return any(word in text for word in ("汇总", "分组", "排序", "小计", "总计", "排名", "top"))


def _mentions_chart_choice(text: str) -> bool:
    return any(
        word in text.lower()
        for word in ("图表", "柱状图", "折线图", "饼图", "看板", "dashboard", "不要图表", "不需要图表")
    )


def _mentions_finance_basis(text: str) -> bool:
    return any(word in text for word in ("人民币", "美元", "币种", "含税", "未税", "税率", "按月", "按年"))


def _mentions_time_basis(text: str) -> bool:
    return any(
        word in text
        for word in ("年", "月", "日", "周", "日期", "时间", "开始", "结束", "周期", "班次")
    )


def _deduplicate(gaps: list[RequirementGap]) -> list[RequirementGap]:
    result: list[RequirementGap] = []
    seen: set[str] = set()
    for item in gaps:
        if item.key in seen:
            continue
        seen.add(item.key)
        result.append(item)
    return result
