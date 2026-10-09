from __future__ import annotations

import re
from pathlib import Path


# ============================================================
# LOCAL PROJECT PATHS
# ============================================================

BASE_DIR = Path(__file__).resolve().parent.parent  # project folder (contains pra/)

DEFAULT_TEMPLATE = str(BASE_DIR / "PRA_test_campagin.html")

DEFAULT_EXCEL = str(next(
    (p for p in (BASE_DIR / "PRA_Test.ods", BASE_DIR / "PRA_Test.xlsx") if p.exists()),
    BASE_DIR / "PRA_Test.ods",
))

DATE_FORMATS = {
    "month_year": "%B %Y",        # September 2026
    "mon_year": "%b %Y",          # Sep 2026
    "day_month_year": "%d-%b-%Y",  # 01-Sep-2026
}

HEX_COLOR = re.compile(r"^#[0-9a-fA-F]{6}$")