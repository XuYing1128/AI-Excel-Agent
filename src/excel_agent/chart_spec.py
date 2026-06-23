"""Map plain-language chart descriptions to concrete chart types.

A non-programmer should be able to say "做个占比饼图" or "看趋势" and get the
right chart, so the UI and the model both share this small vocabulary instead of
each guessing. Types line up with ``rich_workbook_builder.ALLOWED_CHART_TYPES``.
"""

from __future__ import annotations

# Friendly label + one-line hint for every supported chart type (used by the UI
# picker so the user never has to know the English type name).
CHART_TYPE_LABELS: dict[str, str] = {
    "column": "柱状图（竖条，适合分类对比）",
    "bar": "条形图（横条，适合名称较长的对比）",
    "line": "折线图（适合看趋势、随时间变化）",
    "area": "面积图（趋势 + 累积量感）",
    "pie": "饼图（适合看占比、构成）",
    "doughnut": "环形图（占比，中心可留白）",
    "radar": "雷达图（适合多维度能力对比）",
    "scatter": "散点图（适合看两个数值的相关性）",
    "combo": "组合图（柱+线，双指标对比）",
}

# Plain-language keywords -> chart type. Ordered by specificity.
_CHART_KEYWORDS: list[tuple[str, tuple[str, ...]]] = [
    ("combo", ("组合图", "柱线", "柱+线", "双轴", "双坐标", "主次坐标")),
    ("scatter", ("散点", "相关性", "相关图", "分布图", "xy")),
    ("radar", ("雷达", "蜘蛛图", "能力图", "多维对比")),
    ("doughnut", ("环形", "圆环", "甜甜圈", "doughnut", "donut")),
    ("pie", ("饼图", "饼状", "占比", "构成", "比例图", "份额", "pie")),
    ("area", ("面积图", "面积")),
    ("line", ("折线", "趋势", "走势", "曲线", "随时间", "变化趋势", "line")),
    ("bar", ("条形图", "横向柱", "横条", "横向条")),
    ("column", ("柱状", "柱形", "竖向柱", "对比图", "对比", "比较", "排名图", "column", "bar chart")),
]


def chart_type_from_text(text: str, default: str = "column") -> str:
    """Return the best chart type for a free-text description."""

    lowered = str(text or "").lower()
    for chart_type, keywords in _CHART_KEYWORDS:
        if any(keyword.lower() in lowered for keyword in keywords):
            return chart_type
    return default


def wants_chart(text: str) -> bool:
    """True when the text clearly asks for any chart/visual."""

    lowered = str(text or "").lower()
    if any(word in lowered for word in ("不要图", "不需要图", "无需图", "不用图")):
        return False
    triggers = ("图表", "图形", "可视化", "看板", "dashboard", "chart", "graph")
    if any(word in lowered for word in triggers):
        return True
    return chart_type_from_text(text, default="") != ""
