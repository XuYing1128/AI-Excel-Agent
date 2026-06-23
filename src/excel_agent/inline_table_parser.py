"""Extract one or more structured tables embedded in a natural-language prompt."""

from __future__ import annotations

import re
from typing import Any


SECTION_PREFIX_RE = re.compile(r"^\s*(?:表\s*\d+\s*[：:]?\s*)?(.*)$")
NUMBERED_SECTION_RE = re.compile(r"^\s*\d+[.、]\s*(.+)$")
TABLE_NAME_HINTS = (
    "参数表",
    "明细表",
    "基础数据",
    "指标权重",
    "判定标准",
    "调整系数",
    "对照表",
    "数据",
    "名单",
)


def extract_inline_tables(text: str) -> list[dict[str, Any]]:
    """Return stable tabular blocks from tabs, Markdown tables, or CSV-like rows.

    The parser intentionally requires at least a header and one data row. It
    does not guess prose into cells, and it keeps every detected block so a
    multi-table request can survive model failure without becoming a generic
    template.
    """

    lines = str(text or "").splitlines()
    tables: list[dict[str, Any]] = []
    index = 0
    while index < len(lines):
        parsed = _parse_row(lines[index])
        if parsed is None or len(parsed) < 2:
            index += 1
            continue

        block = [parsed]
        delimiter = _delimiter_kind(lines[index])
        cursor = index + 1
        while cursor < len(lines):
            current = _parse_row(lines[cursor], preferred=delimiter)
            if current is None:
                break
            if _is_markdown_separator(current):
                cursor += 1
                continue
            if len(current) != len(block[0]):
                break
            block.append(current)
            cursor += 1

        if len(block) >= 2 and _looks_like_header(block[0], block[1:]):
            headers = _unique_headers(block[0])
            records = [
                {
                    header: _coerce_scalar(value)
                    for header, value in zip(headers, row)
                }
                for row in block[1:]
                if any(str(value).strip() for value in row)
            ]
            if records:
                name = _infer_table_name(lines, index, len(tables) + 1)
                tables.append(
                    {
                        "name": name,
                        "columns": headers,
                        "records": records,
                        "row_count": len(records),
                        "source_line_start": index + 1,
                        "source_line_end": cursor,
                    }
                )
                index = cursor
                continue
        index += 1
    return _deduplicate_tables(tables)


def primary_inline_table(tables: list[dict[str, Any]]) -> dict[str, Any] | None:
    """Choose the main business table, preferring larger and wider blocks."""

    if not tables:
        return None
    return max(
        tables,
        key=lambda item: (
            len(item.get("columns") or []),
            int(item.get("row_count") or 0),
        ),
    )


def _parse_row(line: str, preferred: str | None = None) -> list[str] | None:
    raw = str(line).strip()
    if not raw:
        return None
    kind = preferred or _delimiter_kind(raw)
    if kind == "tab" and "\t" in raw:
        parts = raw.split("\t")
    elif kind == "markdown" and "|" in raw:
        parts = raw.strip("|").split("|")
    elif kind == "csv" and ("," in raw or "，" in raw):
        parts = re.split(r"[,，]", raw)
    else:
        return None
    cleaned = [item.strip() for item in parts]
    return cleaned if len(cleaned) >= 2 else None


def _delimiter_kind(line: str) -> str:
    if "\t" in line:
        return "tab"
    if "|" in line:
        return "markdown"
    if "," in line or "，" in line:
        return "csv"
    return ""


def _is_markdown_separator(row: list[str]) -> bool:
    return bool(row) and all(re.fullmatch(r":?-{2,}:?", item.strip()) for item in row)


def _looks_like_header(header: list[str], rows: list[list[str]]) -> bool:
    if not rows or any(not str(item).strip() for item in header):
        return False
    if len(set(_normalize_header(item) for item in header)) != len(header):
        return False
    text_cells = sum(not _looks_numeric(item) for item in header)
    return text_cells >= max(1, len(header) // 2)


def _infer_table_name(lines: list[str], start: int, sequence: int) -> str:
    for offset in range(1, 5):
        candidate_index = start - offset
        if candidate_index < 0:
            break
        candidate = lines[candidate_index].strip(" \t-*#")
        if not candidate:
            continue
        numbered = NUMBERED_SECTION_RE.match(candidate)
        if numbered:
            candidate = numbered.group(1).strip()
        match = SECTION_PREFIX_RE.match(candidate)
        candidate = match.group(1).strip() if match else candidate
        candidate = candidate.strip("：:")
        quoted = re.search(r"[「“\"]([^」”\"]*(?:表|数据))[^」”\"]*[」”\"]", candidate)
        if quoted:
            return _safe_name(quoted.group(1))
        if len(candidate) <= 40 and (
            any(hint in candidate for hint in TABLE_NAME_HINTS)
            or candidate.endswith("表")
        ):
            return _safe_name(candidate)
        if any(marker in candidate for marker in ("要求", "规则", "说明", "创建", "请在")):
            continue
    return f"内嵌数据表{sequence}"


def _safe_name(value: str) -> str:
    name = re.sub(r"[\[\]:*?/\\]", "_", value).strip()
    name = re.sub(r"^(?:表\s*\d+\s*[：:]?\s*)", "", name)
    name = re.sub(r"^[一二三四五六七八九十]+[、.．]\s*", "", name)
    return (name or "内嵌数据")[:31]


def _unique_headers(headers: list[str]) -> list[str]:
    counts: dict[str, int] = {}
    result: list[str] = []
    for index, raw in enumerate(headers, start=1):
        name = str(raw).strip() or f"第{index}列"
        normalized = _normalize_header(name)
        counts[normalized] = counts.get(normalized, 0) + 1
        result.append(name if counts[normalized] == 1 else f"{name}_{counts[normalized]}")
    return result


def _normalize_header(value: str) -> str:
    return re.sub(r"[\s（）()_\-—:：]", "", str(value)).lower()


def _looks_numeric(value: Any) -> bool:
    text = str(value).strip().replace(",", "").replace("，", "")
    text = text.removesuffix("%").lstrip("+")
    try:
        float(text)
        return True
    except ValueError:
        return False


def _coerce_scalar(value: Any) -> Any:
    text = str(value).strip()
    if not text:
        return ""
    normalized = text.replace(",", "").replace("，", "")
    if normalized.endswith("%"):
        try:
            return float(normalized[:-1].lstrip("+")) / 100
        except ValueError:
            return text
    try:
        number = float(normalized)
        return int(number) if number.is_integer() else number
    except ValueError:
        return text


def _deduplicate_tables(tables: list[dict[str, Any]]) -> list[dict[str, Any]]:
    seen: set[tuple[tuple[str, ...], int, int]] = set()
    result: list[dict[str, Any]] = []
    for table in tables:
        signature = (
            tuple(str(item) for item in table.get("columns", [])),
            int(table.get("source_line_start") or 0),
            int(table.get("source_line_end") or 0),
        )
        if signature in seen:
            continue
        seen.add(signature)
        result.append(table)
    return result
