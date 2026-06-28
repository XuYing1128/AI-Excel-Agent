"""Turn a natural-language request into a deterministic workbook content plan."""

from __future__ import annotations

import re
from pathlib import Path
from typing import Any

from .inline_table_parser import extract_inline_tables, primary_inline_table
from .task_paths import sanitize_filename


TITLE_RE = re.compile(
    r"(?:主题|标题|表名)\s*[：:]\s*(?:[「“\"]([^」”\"\r\n]+)[」”\"]|([^\r\n。]+))"
)
TITLE_ACTION_RE = re.compile(
    r"(?:显示|命名为|名称为)\s*[「“\"]([^」”\"\r\n]+)[」”\"]"
)
COLUMNS_RE = re.compile(
    r"(?:(?:包含|包括|需要|表格需包含)(?:以下)?(?:这些)?列|列为|表头为|字段为)"
    r"\s*[：:]\s*([^。；;\r\n]+)"
)
COLUMN_CLAUSE_STOP_WORDS = (
    "公式",
    "计算",
    "排序",
    "降序",
    "升序",
    "汇总",
    "小计",
    "总计",
    "生成",
    "图表",
    "柱状",
    "折线",
    "饼图",
)
DATE_LINE_RE = re.compile(r"^\s*(\d{1,2}月\d{1,2}日)\s*$")
WEATHER_LINE_RE = re.compile(
    r"^\s*([^：:，,]+)\s*[：:]\s*([^，,]+?)\s*[，,]\s*"
    r"(-?\d+(?:\.\d+)?)\s*[～~—\-至]\s*(-?\d+(?:\.\d+)?)\s*℃?\s*[，,]\s*"
    r"(?:降水概率)?\s*(\d+(?:\.\d+)?)\s*%\s*$"
)

FORMULA_KINDS = {
    "average",
    "difference",
    "product",
    "ratio",
    "sum",
    "weather_advice",
}
SUMMARY_KINDS = {"averageif", "sumif", "countif", "average", "sum", "count"}


def suggest_output_name(
    prompt: str,
    task_type: str,
    title: str | None = None,
) -> str:
    """Return a concise, content-based Chinese workbook filename."""

    candidate = str(title or "").strip() or extract_title(prompt)
    if not candidate:
        labels = {
            "personal_budget": "个人月度收支预算表",
            "family_budget": "家庭年度预算表",
            "quotation": "产品报价单",
            "invoice_draft": "发票草稿与收款明细",
            "inventory": "库存进销存表",
            "sales_report": "销售分析报告",
            "ecommerce_analysis": "电商订单分析报告",
            "project_plan": "项目计划与进度表",
            "schedule": "课程排班表",
            "attendance": "考勤统计表",
            "finance_model": "收入成本利润测算表",
            "dashboard": "综合经营仪表盘",
            "generic_table": "自定义表格",
        }
        candidate = labels.get(task_type, "自定义表格")
    candidate = re.sub(
        r"^(请|帮我|麻烦|根据|基于|制作|生成|做一个|做一份|创建)+",
        "",
        candidate,
    ).strip(" ：:，,。")
    candidate = re.split(r"[，,。；;\n]", candidate, maxsplit=1)[0].strip()
    candidate = sanitize_filename(candidate or "自定义表格", fallback="自定义表格")
    if len(candidate) > 42:
        candidate = candidate[:42].rstrip()
    return f"{candidate}.xlsx"


def extract_title(prompt: str) -> str:
    text = str(prompt or "").strip()
    match = TITLE_RE.search(text)
    if match:
        candidate = next(group for group in match.groups() if group).strip()
        quoted = re.search(r"[「“\"]([^」”\"]+)[」”\"]", candidate)
        return quoted.group(1).strip() if quoted else candidate
    action_match = TITLE_ACTION_RE.search(text)
    if action_match:
        return action_match.group(1).strip()
    for line in text.splitlines():
        clean = line.strip(" \t-*#")
        if not clean:
            continue
        if len(clean) <= 40 and not any(
            marker in clean for marker in ("具体要求", "数据如下", "表格需包含", "请根据")
        ):
            return clean
    return ""


def build_local_content_plan(
    prompt: str,
    task_type: str,
    input_files: list[str] | None = None,
) -> dict[str, Any]:
    """Build a conservative plan from explicit structure and inline records."""

    text = str(prompt or "").strip()
    data_text = _strip_revision_sections(text)
    columns = extract_requested_columns(data_text)
    inline_tables = extract_inline_tables(data_text)
    primary_table = primary_inline_table(inline_tables)
    if primary_table and not columns:
        columns = list(primary_table.get("columns") or [])
    weather_records = extract_weather_records(data_text)
    if weather_records and not columns:
        columns = [
            "日期",
            "城市",
            "天气状况",
            "最高气温（℃）",
            "最低气温（℃）",
            "降水概率",
            "日均气温（℃）",
            "出行建议",
        ]

    records: list[dict[str, Any]] = []
    if weather_records:
        records = [_weather_record_for_columns(item, columns) for item in weather_records]
    elif primary_table:
        records = [dict(item) for item in primary_table.get("records", [])]
    elif columns:
        records = extract_delimited_records(data_text, columns)

    explicit_title_match = TITLE_RE.search(text)
    explicit_title_match = explicit_title_match or TITLE_ACTION_RE.search(text)
    title = extract_title(text)
    if not title:
        title = Path(suggest_output_name(text, task_type)).stem

    formula_rules = infer_formula_rules(text, columns)
    summary_rules = infer_summary_rules(text, columns, records)
    explicit_structure = bool(
        columns and (records or formula_rules or inline_tables or "列" in text)
    )
    single_sheet = not any(
        word in text.lower()
        for word in ("汇总页", "summary sheet", "dashboard", "仪表盘", "数据源页")
    ) and len(inline_tables) <= 1
    sheet_name = _safe_sheet_name(title)
    return {
        "title": title,
        "title_explicit": bool(explicit_title_match),
        "sheet_name": sheet_name,
        "layout": "single_sheet" if single_sheet else "multi_sheet",
        "columns": [
            {
                "name": name,
                "kind": infer_column_kind(name),
                "role": (
                    "formula"
                    if any(rule["target"] == name for rule in formula_rules)
                    else "input"
                ),
            }
            for name in columns
        ],
        "records": records,
        "formula_rules": formula_rules,
        "summary_rules": summary_rules,
        "inline_tables": inline_tables,
        "primary_table_name": (
            str(primary_table.get("name") or "") if primary_table else ""
        ),
        "expected_sheet_names": [
            str(item.get("name") or "") for item in inline_tables
        ],
        "notes": _extract_plain_notes(text),
        "expected_data_rows": len(records),
        "explicit_structure": explicit_structure,
        "source": "local_rules",
        "input_file_names": [Path(item).name for item in (input_files or [])],
    }


def _strip_revision_sections(text: str) -> str:
    """Keep revision/review prose from being parsed as user data tables."""

    markers = (
        "\n本次修改要求：",
        "\n本次修改要求:",
        "\n修正问题：",
        "\n修正问题:",
        "\n采用建议：",
        "\n采用建议:",
    )
    end = len(text)
    for marker in markers:
        index = text.find(marker)
        if index != -1:
            end = min(end, index)
    return text[:end].strip()


def merge_model_content_plan(
    local_plan: dict[str, Any],
    model_plan: dict[str, Any],
) -> dict[str, Any]:
    """Accept only a bounded semantic plan; retain locally extracted source rows."""

    merged = dict(local_plan)
    title = str(model_plan.get("title", "")).strip()
    if title and not merged.get("title_explicit"):
        merged["title"] = title[:80]
        merged["sheet_name"] = _safe_sheet_name(title)

    layout = str(model_plan.get("layout", "")).strip()
    if layout in {"single_sheet", "multi_sheet"}:
        merged["layout"] = layout

    columns = _normalize_columns(model_plan.get("columns"))
    if columns and not merged.get("columns"):
        merged["columns"] = columns
        merged["explicit_structure"] = True

    formula_rules = _normalize_formula_rules(model_plan.get("formula_rules"))
    if formula_rules:
        existing = {
            str(item.get("target")): item
            for item in merged.get("formula_rules", [])
            if item.get("target")
        }
        for item in formula_rules:
            existing.setdefault(item["target"], item)
        merged["formula_rules"] = list(existing.values())
        formula_targets = {item["target"] for item in merged["formula_rules"]}
        for column in merged.get("columns", []):
            if column.get("name") in formula_targets:
                column["role"] = "formula"

    summary_rules = _normalize_summary_rules(model_plan.get("summary_rules"))
    if summary_rules and not merged.get("summary_rules"):
        merged["summary_rules"] = summary_rules

    merged["source"] = "custom_api"
    return merged


def extract_requested_columns(prompt: str) -> list[str]:
    match = COLUMNS_RE.search(str(prompt or ""))
    if not match:
        return []
    raw = match.group(1)
    raw = re.sub(r"^(?:为|是)\s*", "", raw)
    values = [
        re.sub(r"^\d+[.、]\s*", "", item).strip(" `\"“”'")
        for item in re.split(r"[、，,|]", raw)
    ]
    cleaned: list[str] = []
    for item in values:
        if not item:
            continue
        if len(item) > 40 or any(word in item for word in COLUMN_CLAUSE_STOP_WORDS):
            break
        cleaned.append(item)
    return list(dict.fromkeys(cleaned))


def extract_weather_records(prompt: str) -> list[dict[str, Any]]:
    records: list[dict[str, Any]] = []
    current_date = ""
    for raw_line in str(prompt or "").splitlines():
        line = raw_line.strip()
        if not line:
            continue
        date_match = DATE_LINE_RE.match(line)
        if date_match:
            current_date = date_match.group(1)
            continue
        match = WEATHER_LINE_RE.match(line)
        if not match or not current_date:
            continue
        low = float(match.group(3))
        high = float(match.group(4))
        if low > high:
            low, high = high, low
        records.append(
            {
                "日期": current_date,
                "城市": match.group(1).strip(),
                "天气状况": match.group(2).strip(),
                "最高气温（℃）": high,
                "最低气温（℃）": low,
                "降水概率": float(match.group(5)) / 100,
            }
        )
    return records


def extract_delimited_records(prompt: str, columns: list[str]) -> list[dict[str, Any]]:
    marker = re.search(r"数据如下\s*[：:]?", str(prompt or ""))
    if not marker or len(columns) < 2:
        return []
    records: list[dict[str, Any]] = []
    for raw_line in str(prompt)[marker.end() :].splitlines():
        line = raw_line.strip()
        if not line or len(line) > 500:
            continue
        parts = [item.strip() for item in re.split(r"[，,\t|]", line)]
        if len(parts) != len(columns):
            continue
        records.append(dict(zip(columns, parts)))
    return records


def infer_formula_rules(prompt: str, columns: list[str]) -> list[dict[str, Any]]:
    text = str(prompt or "")
    rules: list[dict[str, Any]] = []
    average_target = _find_column(columns, ("日均气温", "平均气温", "平均值"))
    high = _find_column(columns, ("最高气温", "最高"))
    low = _find_column(columns, ("最低气温", "最低"))
    if average_target and high and low and any(word in text for word in ("平均", "日均")):
        rules.append(
            {
                "target": average_target,
                "kind": "average",
                "sources": [high, low],
            }
        )

    advice = _find_column(columns, ("出行建议", "建议"))
    precipitation = _find_column(columns, ("降水概率", "降雨概率"))
    if advice and precipitation and high and any(word in text for word in ("出行建议", "雨具", "防暑")):
        rules.append(
            {
                "target": advice,
                "kind": "weather_advice",
                "sources": [precipitation, high],
                "options": {
                    "rain_threshold": 0.5,
                    "heat_threshold": 30,
                    "rain_text": "建议携带雨具",
                    "heat_text": "注意防暑",
                    "default_text": "适宜出行",
                },
            }
        )
    return rules


def infer_summary_rules(
    prompt: str,
    columns: list[str],
    records: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    text = str(prompt or "")
    if "周平均" not in text and "分组平均" not in text:
        return []
    group_col = _find_column(columns, ("城市", "地区", "区域", "类别"))
    value_col = _find_column(columns, ("日均气温", "平均气温", "金额", "销售额"))
    if not group_col or not value_col:
        return []
    groups = list(
        dict.fromkeys(
            str(record.get(group_col, "")).strip()
            for record in records
            if str(record.get(group_col, "")).strip()
        )
    )
    return [
        {
            "label": f"{group}周平均{value_col}",
            "kind": "averageif",
            "group_col": group_col,
            "group_value": group,
            "value_col": value_col,
        }
        for group in groups
    ]


def infer_column_kind(name: str) -> str:
    lowered = str(name).lower()
    if any(word in lowered for word in ("日期", "date", "月份")):
        return "date"
    if any(word in lowered for word in ("概率", "率", "占比", "roi", "完成度")):
        return "percentage"
    if any(word in lowered for word in ("金额", "单价", "成本", "收入", "利润", "gmv")):
        return "money"
    if any(word in lowered for word in ("数量", "库存", "气温", "小时", "分钟", "天数")):
        return "number"
    return "text"


def _weather_record_for_columns(
    source: dict[str, Any],
    columns: list[str],
) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for column in columns:
        if column in source:
            result[column] = source[column]
            continue
        alias = _weather_alias(column)
        result[column] = source.get(alias, "")
    return result


def _weather_alias(column: str) -> str:
    if "最高" in column:
        return "最高气温（℃）"
    if "最低" in column:
        return "最低气温（℃）"
    if "降水" in column or "降雨" in column:
        return "降水概率"
    if "天气" in column:
        return "天气状况"
    return column


def _normalize_columns(value: Any) -> list[dict[str, str]]:
    if not isinstance(value, list):
        return []
    result: list[dict[str, str]] = []
    for item in value[:30]:
        if isinstance(item, str):
            name = item.strip()
            kind = infer_column_kind(name)
            role = "input"
        elif isinstance(item, dict):
            name = str(item.get("name", "")).strip()
            kind = str(item.get("kind", "")).strip()
            role = str(item.get("role", "input")).strip()
        else:
            continue
        if not name or len(name) > 40:
            continue
        if kind not in {"text", "number", "money", "percentage", "date", "time"}:
            kind = infer_column_kind(name)
        if role not in {"input", "formula"}:
            role = "input"
        result.append({"name": name, "kind": kind, "role": role})
    names: set[str] = set()
    unique: list[dict[str, str]] = []
    for item in result:
        if item["name"] not in names:
            names.add(item["name"])
            unique.append(item)
    return unique


def _normalize_formula_rules(value: Any) -> list[dict[str, Any]]:
    if not isinstance(value, list):
        return []
    result = []
    for item in value[:20]:
        if not isinstance(item, dict):
            continue
        target = str(item.get("target", "")).strip()
        kind = str(item.get("kind", "")).strip()
        sources = item.get("sources", [])
        if not target or kind not in FORMULA_KINDS or not isinstance(sources, list):
            continue
        result.append(
            {
                "target": target,
                "kind": kind,
                "sources": [str(source).strip() for source in sources if str(source).strip()],
                "options": dict(item.get("options") or {}),
            }
        )
    return result


def _normalize_summary_rules(value: Any) -> list[dict[str, Any]]:
    if not isinstance(value, list):
        return []
    result = []
    for item in value[:30]:
        if not isinstance(item, dict):
            continue
        kind = str(item.get("kind", "")).strip()
        label = str(item.get("label", "")).strip()
        if kind not in SUMMARY_KINDS or not label:
            continue
        result.append(
            {
                key: value
                for key, value in item.items()
                if key
                in {
                    "label",
                    "kind",
                    "group_col",
                    "group_value",
                    "value_col",
                    "source_col",
                }
            }
        )
    return result


def _find_column(columns: list[str], keywords: tuple[str, ...]) -> str | None:
    for column in columns:
        if any(keyword.lower() in column.lower() for keyword in keywords):
            return column
    return None


def _safe_sheet_name(title: str) -> str:
    name = re.sub(r"[\[\]:*?/\\]", "_", str(title)).strip() or "数据"
    return name[:31]


def _extract_plain_notes(prompt: str) -> list[str]:
    notes: list[str] = []
    if "人工复核" in prompt:
        notes.append("该结果需要人工复核。")
    return notes
