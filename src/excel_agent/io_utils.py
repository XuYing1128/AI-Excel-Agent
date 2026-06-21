"""File and dataframe IO helpers."""

from __future__ import annotations

import json
import re
import shutil
from datetime import datetime
from pathlib import Path
from typing import Any

import pandas as pd


INVALID_SHEET_CHARS = r"[\[\]\:\*\?\/\\]"


def project_root() -> Path:
    return Path(__file__).resolve().parents[2]


def ensure_output_path(path: str | Path | None, default_name: str = "workbook.xlsx") -> Path:
    out = Path(path) if path else project_root() / "outputs" / default_name
    if not out.is_absolute():
        out = project_root() / out
    out.parent.mkdir(parents=True, exist_ok=True)
    return out


def backup_file(path: str | Path, suffix: str = ".bak") -> Path:
    source = Path(path)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup = source.with_name(f"{source.stem}.{timestamp}{suffix}{source.suffix}")
    shutil.copy2(source, backup)
    return backup


def safe_sheet_name(name: str, fallback: str = "Sheet") -> str:
    clean = re.sub(INVALID_SHEET_CHARS, "_", str(name)).strip("' ")
    clean = clean or fallback
    return clean[:31]


def read_table(path: str | Path, sheet_name: str | int | None = 0) -> pd.DataFrame:
    source = Path(path)
    if not source.exists():
        raise FileNotFoundError(source)
    suffix = source.suffix.lower()
    if suffix in {".csv", ".txt"}:
        return pd.read_csv(source)
    if suffix in {".tsv"}:
        return pd.read_csv(source, sep="\t")
    if suffix in {".xlsx", ".xlsm", ".xls"}:
        return pd.read_excel(source, sheet_name=sheet_name)
    raise ValueError(f"Unsupported table file: {source}")


def save_json(data: dict[str, Any], path: str | Path) -> Path:
    output = ensure_output_path(path, "report.json")
    output.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    return output


def normalize_output_extension(path: str | Path, suffix: str = ".xlsx") -> Path:
    p = Path(path)
    return p if p.suffix else p.with_suffix(suffix)

