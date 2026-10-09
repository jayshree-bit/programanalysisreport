from __future__ import annotations

import json
import pandas as pd
import re
from pathlib import Path
from typing import Any


def esc(value: Any) -> str:
    from html import escape
    return escape(str(value))

def slugify(value: str) -> str:
    value = re.sub(r"[^A-Za-z0-9]+", "_", str(value).strip())
    return value.strip("_") or "Program_Analysis_Report"

def num(value: Any, default: int | None = None) -> int | None:
    if value is None or str(value).strip() == "":
        return default
    try:
        return int(float(str(value).replace(",", "").strip()))
    except Exception:
        return default

def pct(a: int | float | None, b: int | float | None) -> float | None:
    if a is None or b in (None, 0):
        return None
    return (float(a) / float(b)) * 100

def pct_text(value: float | None) -> str:
    return "N/A" if value is None else f"{value:.2f}%"

def fmt(value: int | float | None) -> str:
    return "N/A" if value is None else f"{int(round(value)):,}"

def series_clean(df: pd.DataFrame, column: str | None) -> pd.Series:
    if not column:
        return pd.Series(dtype=str)
    return df[column].fillna("").astype(str).str.strip()

def find_column(df: pd.DataFrame, candidates: list[str]) -> str | None:
    wanted = {re.sub(r"[^a-z0-9]", "", x.lower()) for x in candidates}
    for column in df.columns:
        key = re.sub(r"[^a-z0-9]", "", str(column).lower())
        if key in wanted:
            return column
    return None

def top_counts(df: pd.DataFrame, column: str | None, n: int = 10) -> list[tuple[str, int]]:
    s = series_clean(df, column)
    if s.empty:
        return []
    s = s[s != ""]
    return [(str(k), int(v)) for k, v in s.value_counts().head(n).items()]

def read_raw_leads(excel_path: Path) -> pd.DataFrame:
    if excel_path.suffix.lower() == ".csv":
        return pd.read_csv(excel_path)
    if excel_path.suffix.lower() == ".ods":
        try:
            with pd.ExcelFile(excel_path, engine="odf") as book:
                sheet = "Sheet1" if "Sheet1" in book.sheet_names else book.sheet_names[0]
                return book.parse(sheet)
        except Exception:
            pass

    with pd.ExcelFile(excel_path) as book:
        sheet = "Sheet1" if "Sheet1" in book.sheet_names else book.sheet_names[0]
        return book.parse(sheet)

def ordered_job_levels(pairs: list[tuple[str, int]]) -> list[tuple[str, int]]:
    order = {
        "c-level": 0,
        "c level": 0,
        "c-suite": 0,
        "vice president": 1,
        "vp": 1,
        "director": 2,
        "manager": 3,
    }
    return sorted(pairs, key=lambda x: (order.get(x[0].strip().lower(), 50), x[0]))

def safe_json(value: Any) -> str:
    return json.dumps(value, ensure_ascii=False)

def business_days(start_text: str, end_text: str) -> int | None:
    try:
        start = pd.to_datetime(start_text, dayfirst=True)
        end = pd.to_datetime(end_text, dayfirst=True)
        return int(pd.bdate_range(start, end).size)
    except Exception:
        return None

# ============================================================
# ASSET METRICS
# ============================================================

def numeric_column(df: pd.DataFrame, candidates: list[str]) -> str | None:
    column = find_column(df, candidates)
    if not column:
        return None
    test = pd.to_numeric(df[column], errors="coerce")
    return column if test.notna().any() else None