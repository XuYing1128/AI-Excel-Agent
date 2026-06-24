"""Map plain-language chart descriptions to concrete chart types.

A non-programmer should be able to say "做个占比饼图" or "看趋势" and get the
right chart, so the UI and the model both share this small vocabulary instead of
each guessing. Types line up with ``rich_workbook_builder.ALLOWED_CHART_TYPES``.
"""

from __future__ import annotations

from typing import Any

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
    ("line", ("折线", "趋势", "趋势图", "走势", "曲线", "随时间", "变化趋势", "line")),
    ("bar", ("条形图", "横向柱", "横条", "横向条")),
    (
        "column",
        (
            "柱状",
            "柱形",
            "竖向柱",
            "柱状对比",
            "柱状对比图",
            "对比图",
            "对比",
            "比较",
            "排名图",
            "排行榜",
            "top n",
            "topn",
            "column",
            "bar chart",
        ),
    ),
]

_NEGATIVE_CHART_WORDS = (
    "不要图",
    "不需要图",
    "无需图",
    "不用图",
    "删除图表",
    "去掉图表",
    "移除图表",
    "纯表格不带图",
    "纯表格不要图",
)

_GENERIC_CHART_TRIGGERS = (
    "图表",
    "图形",
    "图示",
    "可视化",
    "看板",
    "dashboard",
    "chart",
    "graph",
    "数据透视图",
    "透视图",
)


def chart_type_from_text(text: str, default: str | None = "column") -> str | None:
    """Return the best chart type for a free-text description."""

    lowered = str(text or "").lower()
    for chart_type, keywords in _CHART_KEYWORDS:
        if any(keyword.lower() in lowered for keyword in keywords):
            return chart_type
    return default


def chart_types_from_text(text: str, default: str | None = "column") -> list[str]:
    """Return all explicitly mentioned chart types in priority order."""

    lowered = str(text or "").lower()
    result: list[str] = []
    for chart_type, keywords in _CHART_KEYWORDS:
        if any(keyword.lower() in lowered for keyword in keywords):
            result.append(chart_type)
    if not result and default:
        result.append(default)
    return list(dict.fromkeys(result))


def normalize_chart_types(value: Any) -> list[str]:
    """Normalize UI/model chart type values to supported internal names."""

    if value is None:
        return []
    raw_values = value if isinstance(value, list) else [value]
    normalized: list[str] = []
    for item in raw_values:
        text = str(item or "").strip()
        if not text:
            continue
        key = text.lower()
        if key in CHART_TYPE_LABELS:
            normalized.append(key)
            continue
        guessed = chart_type_from_text(text, default="")
        if guessed:
            normalized.append(guessed)
    return list(dict.fromkeys(normalized))


def wants_chart(text: str) -> bool:
    """True when the text clearly asks for any chart/visual."""

    lowered = str(text or "").lower()
    if any(word in lowered for word in _NEGATIVE_CHART_WORDS):
        return False
    if any(word in lowered for word in _GENERIC_CHART_TRIGGERS):
        return True
    return chart_type_from_text(text, default="") != ""


def analyze_chart_requirements(
    text: str,
    *,
    force_include: bool = False,
    requested_types: Any = None,
) -> dict[str, Any]:
    """Extract whether charts are required and which types should be generated.

    ``数据透视表`` is a table requirement, not a chart. ``数据透视图`` or
    ``数据透视图表`` is a chart requirement.
    """

    raw = str(text or "")
    lowered = raw.lower()
    negative = any(word in lowered for word in _NEGATIVE_CHART_WORDS)
    selected_types = normalize_chart_types(requested_types)
    mentioned_types = chart_types_from_text(raw, default=None)
    generic_trigger = any(word in lowered for word in _GENERIC_CHART_TRIGGERS)
    chart_word_trigger = wants_chart(raw)
    required = bool(force_include or selected_types or chart_word_trigger)
    if negative and not force_include:
        required = False
        selected_types = []
        mentioned_types = []

    chart_types = selected_types or mentioned_types
    if required and not chart_types:
        chart_types = ["column"]

    reason = ""
    if negative:
        reason = "用户明确表示不需要图表"
    elif selected_types:
        reason = "用户在界面选择了图表类型"
    elif mentioned_types:
        reason = "需求文字中指定了图表类型"
    elif generic_trigger or force_include:
        reason = "需求文字或确认项要求生成图表"
    else:
        reason = "未识别到明确图表要求"

    return {
        "required": required,
        "types": chart_types,
        "explicit": required and not negative,
        "negative": negative,
        "reason": reason,
    }
