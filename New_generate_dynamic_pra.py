from __future__ import annotations

"""
Dynamic Program Analysis Report generator.

IMPORTANT:
- The visual HTML template is NOT redesigned.
- This script uses the supplied HTML as the template and injects a dynamic
  data/calculation layer at the end of the existing file.
- The same layout, CSS, slides, charts and branding are retained.
- Campaign inputs are requested one-by-one.
- Lead-level audience / industry / location / asset data comes from the raw Excel/ODS.
- Sent / Delivered / Opens / Clicks / Conversion are entered by the user.
- Bounce is calculated automatically as Sent - Delivered.
- Open / Click / Conversion rates are calculated automatically.
- The last slide's observations/recommendations are generated from the data.

Project folder:
./

Files:
1. PRA_test_campagin.html                                     (existing page style to preserve)
2. PRA_Test.ods                                               (preferred raw leads source)
3. PRA_Test.xlsx                                              (fallback raw leads source)
4. New_generate_dynamic_pra.py                                 (this script)
"""

import argparse
import json
import math
import os
import re
from pathlib import Path
from typing import Any
from urllib.parse import quote

import pandas as pd

try:
    from countryinfo import CountryInfo
except ImportError:
    CountryInfo = None

try:
    import plotly  # noqa: F401
except ImportError:
    plotly = None


# ============================================================
# LOCAL PROJECT PATHS
# ============================================================

BASE_DIR = Path(__file__).resolve().parent
DEFAULT_TEMPLATE = str(BASE_DIR / "PRA_test_campagin.html")
DEFAULT_EXCEL = str(next(
    (p for p in (BASE_DIR / "PRA_Test.ods", BASE_DIR / "PRA_Test.xlsx") if p.exists()),
    BASE_DIR / "PRA_Test.ods",
))


# ============================================================
# HELPERS
# ============================================================

US_STATE_COORDS = {
    "Alabama": (32.8067, -86.7911),
    "Alaska": (61.3707, -152.4044),
    "Arizona": (33.7298, -111.4312),
    "Arkansas": (34.9697, -92.3731),
    "California": (36.1162, -119.6816),
    "Colorado": (39.0598, -105.3111),
    "Connecticut": (41.5978, -72.7554),
    "Delaware": (39.3185, -75.5071),
    "Florida": (27.7663, -81.6868),
    "Georgia": (33.0406, -83.6431),
    "Hawaii": (21.0943, -157.4983),
    "Idaho": (44.2405, -114.4788),
    "Illinois": (40.3495, -88.9861),
    "Indiana": (39.8494, -86.2583),
    "Iowa": (42.0115, -93.2105),
    "Kansas": (38.5266, -96.7265),
    "Kentucky": (37.6681, -84.6701),
    "Louisiana": (31.1695, -91.8678),
    "Maine": (44.6939, -69.3819),
    "Maryland": (39.0639, -76.8021),
    "Massachusetts": (42.2302, -71.5301),
    "Michigan": (43.3266, -84.5361),
    "Minnesota": (45.6945, -93.9002),
    "Mississippi": (32.7416, -89.6787),
    "Missouri": (38.4561, -92.2884),
    "Montana": (46.9219, -110.4544),
    "Nebraska": (41.1254, -98.2681),
    "Nevada": (38.3135, -117.0554),
    "New Hampshire": (43.4525, -71.5639),
    "New Jersey": (40.2989, -74.5210),
    "New Mexico": (34.8405, -106.2485),
    "New York": (42.1657, -74.9481),
    "North Carolina": (35.6301, -79.8064),
    "North Dakota": (47.5289, -99.7840),
    "Ohio": (40.3888, -82.7649),
    "Oklahoma": (35.5653, -96.9289),
    "Oregon": (44.5720, -122.0709),
    "Pennsylvania": (40.5908, -77.2098),
    "Rhode Island": (41.6809, -71.5118),
    "South Carolina": (33.8569, -80.9450),
    "South Dakota": (44.2998, -99.4388),
    "Tennessee": (35.7478, -86.6923),
    "Texas": (31.0545, -97.5635),
    "Utah": (40.1500, -111.8624),
    "Vermont": (44.0459, -72.7107),
    "Virginia": (37.7693, -78.1700),
    "Washington": (47.4009, -121.4905),
    "West Virginia": (38.4912, -80.9545),
    "Wisconsin": (44.2685, -89.6165),
    "Wyoming": (42.7560, -107.3025),
    "District of Columbia": (38.9060, -77.0334),
}

COUNTRY_ALIAS = {
    "USA": "United States",
    "US": "United States",
    "United States of America": "United States",
    "UK": "United Kingdom",
    "U.K.": "United Kingdom",
    "UAE": "United Arab Emirates",
}

COUNTRY_COORDS = {
  "Canada": (56.1304, -106.3468),
  "France": (46.2276, 2.2137),
  "Germany": (51.1657, 10.4515),
  "Netherlands": (52.1326, 5.2913),
  "United Kingdom": (55.3781, -3.4360),
  "United States": (39.8283, -98.5795),
}


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


def parse_date_input(prompt: str, optional: bool = False) -> str | None:
    while True:
        value = input(prompt).strip()
        if optional and not value:
            return None
        if not value:
            print("Please enter a value.")
            continue
        dt = pd.to_datetime(value, errors="coerce")
        if pd.notna(dt):
            return dt.strftime("%d-%b-%Y")
        print("Invalid date. Example: 21-Jul-2026")


def input_int(prompt: str, minimum: int = 0) -> int:
    while True:
        raw = input(prompt).strip().replace(",", "")
        try:
            value = int(raw)
            if value < minimum:
                raise ValueError
            return value
        except Exception:
            print(f"Please enter a whole number >= {minimum}.")


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


def asset_event_metrics(
    df: pd.DataFrame,
    asset_col: str | None,
    total_opens: int,
    total_clicks: int,
    total_conversions: int,
) -> tuple[list[tuple[str, int]], list[tuple[str, int]], list[tuple[str, int]], str]:

    assets = top_counts(df, asset_col, 8)
    if not assets:
        return [], [], [], "No asset data available."

    open_col = numeric_column(df, ["Opens", "Open", "Open Count", "Unique Opens"])
    click_col = numeric_column(df, ["Clicks", "Click", "Click Count", "Unique Clicks"])
    conv_col = numeric_column(df, ["Conversion", "Conversions", "Conversion Count"])

    asset_series = series_clean(df, asset_col)
    actual_available = bool(open_col or click_col or conv_col)

    def group_numeric(column: str | None, total: int) -> list[tuple[str, int]]:
        if not column:
            return []
        tmp = pd.DataFrame({
            "asset": asset_series,
            "metric": pd.to_numeric(df[column], errors="coerce").fillna(0)
        })
        grouped = tmp[tmp["asset"] != ""].groupby("asset")["metric"].sum().sort_values(ascending=False)
        return [(str(k), int(round(v))) for k, v in grouped.items()]

    if actual_available:
        opens = group_numeric(open_col, total_opens)
        clicks = group_numeric(click_col, total_clicks)
        convs = group_numeric(conv_col, total_conversions)

        def fallback_total(groups: list[tuple[str, int]], total: int) -> list[tuple[str, int]]:
            if groups:
                return groups
            return allocate_by_asset_share(assets, total)

        return (
            fallback_total(opens, total_opens),
            fallback_total(clicks, total_clicks),
            fallback_total(convs, total_conversions),
            "Asset-level engagement values were read from numeric event columns in the raw Excel."
        )

    return (
        allocate_by_asset_share(assets, total_opens),
        allocate_by_asset_share(assets, total_clicks),
        allocate_by_asset_share(assets, total_conversions),
        "Asset-level Open/Click/Conversion columns were not found. The report allocates total campaign events in proportion to asset lead share; this is a calculated allocation, not measured asset engagement."
    )


def allocate_by_asset_share(assets: list[tuple[str, int]], total: int) -> list[tuple[str, int]]:
    if not assets:
        return []
    asset_total = sum(v for _, v in assets)
    raw = [(name, (value / asset_total) * total) for name, value in assets]
    rounded = [(name, int(math.floor(x))) for name, x in raw]
    remainder = total - sum(v for _, v in rounded)

    # Distribute remainder by largest fractional component.
    order = sorted(
        range(len(raw)),
        key=lambda i: (raw[i][1] - math.floor(raw[i][1])),
        reverse=True
    )
    for i in order[:remainder]:
        rounded[i] = (rounded[i][0], rounded[i][1] + 1)
    return rounded


# ============================================================
# LOCATION MAP
# ============================================================

def country_coord(country: str) -> tuple[float, float] | None:
    country = COUNTRY_ALIAS.get(country, country)
    if country in COUNTRY_COORDS:
        return COUNTRY_COORDS[country]
    if CountryInfo is None:
        return None
    try:
        value = CountryInfo(country).info().get("latlng")
        if value and len(value) == 2:
            return float(value[0]), float(value[1])
    except Exception:
        pass
    return None


def location_points(df: pd.DataFrame) -> list[tuple[str, float, float, int]]:
    country_col = find_column(df, ["Country", "Country Name"])
    result: list[tuple[str, float, float, int]] = []
    country_series = series_clean(df, country_col)
    counts = country_series[country_series != ""].value_counts()
    for location, count in counts.items():
        coord = country_coord(location)
        if coord:
            result.append((location, coord[0], coord[1], int(count)))
    return result


def build_plotly_geo_spec(df: pd.DataFrame) -> dict[str, Any]:
    points = location_points(df)

    if not points:
        return {
            "available": False,
            "points": [],
            "message": "No valid geographic data was found in the raw lead file."
        }

    max_size = max(x[3] for x in points) if points else 1

    return {
        "available": True,
        "points": points,
        "trace": {
        "type": "scatter3d",
        "mode": "markers+text",
        "x": [x[2] for x in points],
        "y": [x[1] for x in points],
        "z": [x[3] for x in points],
            "text": [x[0] for x in points],
        "customdata": [x[3] for x in points],
        "textposition": "top center",
        "hovertemplate": "<b>%{text}</b><br>Leads: %{customdata}<extra></extra>",
            "marker": {
          "size": [8 + (x[3] / max_size) * 18 for x in points],
                "color": [x[3] for x in points],
                "colorscale": [[0.0, "#ffe0c2"], [0.5, "#f47b20"], [1.0, "#a63e00"]],
                "showscale": True,
                "colorbar": {
                    "title": "Leads",
                    "thickness": 10,
                    "len": 0.65,
                    "tickfont": {"size": 9},
                },
                "opacity": 0.80,
                "line": {"width": 1, "color": "#ffffff"},
            },
        },
        "layout": {
        "height": 220,
        "margin": {"l": 0, "r": 0, "t": 4, "b": 0},
            "paper_bgcolor": "rgba(0,0,0,0)",
            "plot_bgcolor": "rgba(0,0,0,0)",
        "scene": {
          "xaxis": {"title": "Longitude", "gridcolor": "#d7dde5", "zerolinecolor": "#d7dde5"},
          "yaxis": {"title": "Latitude", "gridcolor": "#d7dde5", "zerolinecolor": "#d7dde5"},
          "zaxis": {"title": "Leads", "gridcolor": "#d7dde5", "rangemode": "tozero"},
          "bgcolor": "rgba(0,0,0,0)",
          "camera": {"eye": {"x": 1.45, "y": 1.45, "z": 0.95}},
        },
        }
    }



# ============================================================
# ALL-COLUMN SUMMARY (every populated column in the raw file)
# ============================================================

COLUMN_COLORS = ["#3f5bd8", "#f47b20", "#12a8b8", "#18a878", "#7c5ce5", "#e66aa4"]


def column_summaries(df: pd.DataFrame, top_n: int = 6) -> list[dict[str, Any]]:
    """Count + percentage for every populated, categorical column in the file."""
    result: list[dict[str, Any]] = []
    total = len(df)
    seen: list[pd.Series] = []
    for column in df.columns:
        s = series_clean(df, column)
        s = s[~s.isin(["", "nan", "NaN", "None"])]
        if s.empty:
            continue  # completely empty column (e.g. City, Asset2) - skipped
        unique = int(s.nunique())
        if unique <= 1:
            continue  # same value on every row (e.g. Lead Type = BANT) - nothing to chart
        dup_idx = next((i for i, prev in enumerate(seen) if s.equals(prev)), None)
        if dup_idx is not None:
            # exact duplicate of an earlier column (e.g. State = Country): keep the clearer "Country" name
            if "country" in str(column).lower():
                result[dup_idx]["column"] = str(column)
            continue
        seen.append(s)
        if total > 20 and unique > total * 0.9:
            continue  # free-text / id style column
        counts = s.value_counts()
        pairs = [(str(k), int(v)) for k, v in counts.head(top_n).items()]
        rest = int(counts.iloc[top_n:].sum()) if unique > top_n else 0
        result.append({
            "column": str(column),
            "filled": int(len(s)),
            "unique": unique,
            "pairs": pairs,
            "other": rest,
            "total": total,
        })
    return result


def build_summary_slide(summaries: list[dict[str, Any]], total: int) -> str:
    panels = []
    for idx, item in enumerate(summaries):
        rows = []
        top = max((c for _, c in item["pairs"]), default=1)
        for i, (name, count) in enumerate(item["pairs"]):
            color = COLUMN_COLORS[i % len(COLUMN_COLORS)]
            width = max(3, round(count / top * 100))
            rows.append(
                f'<div class="pra-row"><div class="pra-rl"><span title="{esc(name)}">{esc(name)}</span>'
                f'<b>{count:,} &middot; {pct_text(pct(count, total))}</b></div>'
                f'<div class="pra-bar"><i style="width:{width}%;background:{color}"></i></div></div>'
            )
        if item["other"]:
            rows.append(
                f'<div class="pra-row"><div class="pra-rl"><span>Other ({item["unique"] - len(item["pairs"])} more)</span>'
                f'<b>{item["other"]:,} &middot; {pct_text(pct(item["other"], total))}</b></div></div>'
            )
        panels.append(
            f'<div class="panel pra-panel"><div class="section">{esc(item["column"])}</div>'
            f'<div class="pra-sub">{item["unique"]} distinct &middot; {item["filled"]:,} of {total:,} filled</div>'
            + "".join(rows) + "</div>"
        )
    css = (
        "<style>"
        "#s13 .pra-cols{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:10px;"
        "max-height:calc(100% - 120px);overflow:auto;padding-right:4px}"
        "#s13 .pra-panel{padding:10px 12px}"
        "#s13 .pra-sub{font-size:9px;color:#64748b;margin:-2px 0 6px}"
        "#s13 .pra-row{margin-bottom:5px}"
        "#s13 .pra-rl{display:flex;justify-content:space-between;gap:6px;font-size:10px;color:#334155}"
        "#s13 .pra-rl span{overflow:hidden;text-overflow:ellipsis;white-space:nowrap}"
        "#s13 .pra-rl b{color:#f47b20;white-space:nowrap}"
        "#s13 .pra-bar{height:6px;background:#eef2f8;border-radius:4px;overflow:hidden;margin-top:2px}"
        "#s13 .pra-bar i{display:block;height:100%;border-radius:4px}"
        "@media(max-width:1100px){#s13 .pra-cols{grid-template-columns:repeat(2,minmax(0,1fr))}}"
        "</style>"
    )
    return (
        css + '<section class="slide hidden" id="s13"><div class="kicker">Lead Data Summary</div>'
        '<h2 class="title">Every Lead Field &middot; Count &amp; Percentage</h2>'
        '<div class="pra-cols">' + "".join(panels) + "</div>"
        '<div class="footer">VALASYS MEDIA</div><div class="slideNo">13</div></section>'
    )


def add_summary_slide(html: str, summaries: list[dict[str, Any]], total: int) -> str:
    """Adds slide 13 + nav entry + contents row. Existing slides are untouched."""
    html = re.sub(r"<style>\s*#s13 .*?</style>\s*<section class=\"slide hidden\" id=\"s13\">.*?</section>\s*", "", html, flags=re.DOTALL)
    html = re.sub(r"<button class=\"slide-nav-item\" onclick=\"goToSlide\(12\)\">.*?</button>\s*", "", html, flags=re.DOTALL)
    html = re.sub(r"<div class=\"toc-row\"><span>08</span><b>Lead Data Summary</b><span>13</span></div>", "", html)
    slide = build_summary_slide(summaries, total)
    m = re.search(r'<section class="slide hidden" id="s12">.*?</section>', html, flags=re.DOTALL)
    if m:
        html = html[:m.end()] + "\n" + slide + html[m.end():]
    nav = ('<button class="slide-nav-item" onclick="goToSlide(12)"><span class="thumb">13</span>'
           '<span><b>Lead Data Summary</b><small>All fields</small></span></button>')
    m = re.search(r'<button class="slide-nav-item"[^>]*goToSlide\(11\)[^>]*>.*?</button>', html, flags=re.DOTALL)
    if m:
        html = html[:m.end()] + nav + html[m.end():]
    toc = '<div class="toc-row"><span>08</span><b>Lead Data Summary</b><span>13</span></div>'
    html = re.sub(r'(<div class="toc-row"><span>07</span><b>Observations &amp; Recommendations</b><span>12</span></div>)', r"\1" + toc, html)
    html = re.sub(r'(<div class="toc-row"><span>07</span><b>Observations & Recommendations</b><span>12</span></div>)', r"\1" + toc, html)
    return html


# ============================================================
# DATA PREPARATION
# ============================================================

def prepare_data(df: pd.DataFrame, inputs: dict[str, Any]) -> dict[str, Any]:
    total_leads = len(df)

    job_level_col = find_column(df, ["Job Level", "Seniority", "Seniority Level"])
    job_function_col = find_column(df, ["Job Function", "Function"])
    industry_col = find_column(df, ["Industry", "Industries"])
    company_size_col = find_column(df, ["Company Size", "Employee Size", "Company Size Range"])
    asset_col = find_column(df, ["Asset", "Asset Name", "Content Asset", "Promoted Asset"])
    lead_type_col = find_column(df, ["Lead Type", "LeadType"])
    decision_col = find_column(df, ["Decision Makers", "Decision Maker", "Persona"])
    device_col = find_column(df, ["Device", "Devices", "Device Type", "Device Category"])

    job_levels = ordered_job_levels(top_counts(df, job_level_col, 10))
    job_functions = top_counts(df, job_function_col, 6)
    industries = top_counts(df, industry_col, 8)
    company_sizes = top_counts(df, company_size_col, 8)
    assets = top_counts(df, asset_col, 8)
    unique_asset_count = int(series_clean(df, asset_col).replace("", pd.NA).nunique())
    decisions = top_counts(df, decision_col, 6)
    devices = top_counts(df, device_col, 6)
    device_mix_is_estimate = False
    # No device column in the file -> show real Decision Maker / Influencer counts instead of made-up numbers.
    use_decision_mix = not bool(devices)
    device_title = "Decision Maker Mix" if use_decision_mix else "Devices"
    if use_decision_mix:
        devices = list(decisions)
    lead_types = top_counts(df, lead_type_col, 6)

    sent = inputs["sent"]
    delivered = inputs["delivered"]
    opens = inputs["opens"]
    clicks = inputs["clicks"]
    conversions = inputs["conversion"]
    bounced = max(sent - delivered, 0)

    open_rate = pct(opens, delivered)
    click_rate = pct(clicks, delivered)
    conversion_rate = pct(conversions, delivered)
    delivery_rate = pct(delivered, sent)
    bounce_rate = pct(bounced, sent)

    asset_opens, asset_clicks, asset_conversions, asset_event_note = asset_event_metrics(
        df, asset_col, opens, clicks, conversions
    )

    country_col = find_column(df, ["Country", "Country Name"])
    country_counts = top_counts(df, country_col, len(df))
    geo = location_points(df)
    geo_total = sum(count for _, count in country_counts)

    return {
        "lead_count": total_leads,
        "campaign_name": inputs["campaign_name"],
        "report_date": inputs["report_date"],
        "start_date": inputs["start_date"],
        "end_date": inputs["end_date"],
        "prepared_by": inputs["prepared_by"],
        "sent": sent,
        "delivered": delivered,
        "opens": opens,
        "clicks": clicks,
        "conversion": conversions,
        "bounced": bounced,
        "delivery_rate": delivery_rate,
        "open_rate": open_rate,
        "click_rate": click_rate,
        "conversion_rate": conversion_rate,
        "bounce_rate": bounce_rate,
        "business_days": business_days(inputs["start_date"], inputs["end_date"]),
        "job_levels": job_levels,
        "job_functions": job_functions,
        "industries": industries,
        "company_sizes": company_sizes,
        "assets": assets,
        "unique_asset_count": unique_asset_count,
        "decisions": decisions,
        "devices": devices,
        "device_mix_is_estimate": device_mix_is_estimate,
        "device_title": device_title,
        "lead_types": lead_types,
        "asset_opens": asset_opens,
        "asset_clicks": asset_clicks,
        "asset_conversions": asset_conversions,
        "asset_event_note": asset_event_note,
        "geo": geo,
        "country_counts": country_counts,
        "geo_total": geo_total,
        "maps_key": inputs.get("maps_key", ""),
        "industry_col": industry_col,
        "asset_col": asset_col,
    }


# ============================================================
# HTML DYNAMIC LAYER
# ============================================================

def chart_payload(pairs: list[tuple[str, int]]) -> tuple[str, str]:
    return safe_json([x[0] for x in pairs]), safe_json([x[1] for x in pairs])


def top_two_pairs(pairs: list[tuple[str, int]]) -> list[tuple[str, int]]:
    return pairs[:2]


def build_matrix(df: pd.DataFrame, row_col: str | None, col_col: str | None, rows: list[str], cols: list[str]) -> list[list[int]]:
    if not row_col or not col_col:
        return [[0 for _ in cols] for _ in rows]
    temp = pd.DataFrame({
        "row": series_clean(df, row_col),
        "col": series_clean(df, col_col),
    })
    pivot = pd.crosstab(temp["row"], temp["col"])
    result = []
    for r in rows:
        result.append([int(pivot.loc[r, c]) if r in pivot.index and c in pivot.columns else 0 for c in cols])
    return result


PRA_EXTRA_JS = r'''// PRA-EXTRA: count + percentage labels, Google Maps location slide, decision-maker mix
var PRA_PAL=['#4e7fd8','#f28a3d','#12a8b8','#18a878','#7c5ce5','#e66aa4'];
if(typeof PRA!=='undefined'&&PRA.palette&&PRA.palette.length>=2)PRA_PAL=PRA.palette.slice();
function praC(i){return PRA_PAL[i%PRA_PAL.length];}
function praFmtN(n){return Number(n||0).toLocaleString('en-US');}
function praEsc(t){return String(t).replace(/[&<>"']/g,function(c){return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c];});}

/* ---------------- Chart style: 2D / 3D ---------------- */
var PRA_MODE=(function(){try{return localStorage.getItem('praChartMode')||'3d';}catch(e){return '3d';}})();
var PRA_DEPTH=9;
function praIs3D(){return PRA_MODE==='3d';}
function praShade(c,f){var m=/^#([0-9a-f]{6})$/i.exec(c||'');if(!m)return c;var n=parseInt(m[1],16),r=n>>16,g=(n>>8)&255,b=n&255;
  var t=function(v){return Math.round(f<0?v*(1+f):v+(255-v)*f);};return 'rgb('+t(r)+','+t(g)+','+t(b)+')';}
function praColorAt(ds,i){var c=ds.praColors||ds.hoverBackgroundColor||ds.backgroundColor;return Array.isArray(c)?c[i%c.length]:c;}
function praDepthOf(chart,el){var horiz=chart.options.indexAxis==='y';var s=horiz?el.height:el.width;return Math.max(3,Math.min(PRA_DEPTH,(s||20)*0.35));}

if (typeof Chart !== 'undefined') {
  // 3D extrusion: side/top faces for bars, a raised base for doughnuts. Drawn before the front faces.
  Chart.register({id:'praDepth', beforeDatasetsDraw:function(chart){
    if(!praIs3D())return; var ctx=chart.ctx, t=chart.config.type;
    ctx.save();
    chart.data.datasets.forEach(function(ds,di){
      var meta=chart.getDatasetMeta(di); if(meta.hidden)return;
      meta.data.forEach(function(el,i){
        var col=praColorAt(ds,i); if(typeof col!=='string')return;
        if(t==='doughnut'||t==='pie'){
          if(!chart.getDataVisibility(i))return;
          var D=PRA_DEPTH; ctx.fillStyle=praShade(col,-0.35);
          for(var k=D;k>=1;k--){ctx.beginPath();ctx.arc(el.x,el.y+k,el.outerRadius,el.startAngle,el.endAngle);
            ctx.arc(el.x,el.y+k,el.innerRadius,el.endAngle,el.startAngle,true);ctx.closePath();ctx.fill();}
          return;
        }
        if(t!=='bar')return;
        var horiz=chart.options.indexAxis==='y', d=praDepthOf(chart,el);
        var L,R,T,B;
        if(horiz){L=Math.min(el.x,el.base);R=Math.max(el.x,el.base);T=el.y-el.height/2;B=el.y+el.height/2;}
        else{L=el.x-el.width/2;R=el.x+el.width/2;T=Math.min(el.y,el.base);B=Math.max(el.y,el.base);}
        if(!(R-L>0.5)||!(B-T>0.5))return;
        ctx.fillStyle=praShade(col,0.35);
        ctx.beginPath();ctx.moveTo(L,T);ctx.lineTo(L+d,T-d);ctx.lineTo(R+d,T-d);ctx.lineTo(R,T);ctx.closePath();ctx.fill();
        ctx.fillStyle=praShade(col,-0.3);
        ctx.beginPath();ctx.moveTo(R,T);ctx.lineTo(R+d,T-d);ctx.lineTo(R+d,B-d);ctx.lineTo(R,B);ctx.closePath();ctx.fill();
      });
    });
    ctx.restore();
  }});

  // Count + percentage labels that never overlap: one line, two lines, count only, or hidden (tooltip still works).
  Chart.register({id:'praLabels', afterDatasetsDraw:function(chart){
    var o=chart.options.plugins&&chart.options.plugins.praLabels; if(!o)return;
    var ctx=chart.ctx, horiz=chart.options.indexAxis==='y', t=chart.config.type, d3=praIs3D();
    ctx.save(); ctx.textBaseline='middle';
    chart.data.datasets.forEach(function(ds,di){
      var meta=chart.getDatasetMeta(di); if(meta.hidden)return;
      var sum=(ds.data||[]).reduce(function(a,b){return a+Number(b||0);},0);
      meta.data.forEach(function(el,i){
        var c,p;
        if(o.grouped){c=Number(ds.data[i]||0); if(!c)return; p=sum?c/sum*100:0;}
        else{c=o.counts[i]; if(c==null)return; p=o.pcts?o.pcts[i]:(o.total?c/o.total*100:0);}
        if(t==='doughnut'){
          if(!chart.getDataVisibility(i)||p<7)return; var pos=el.tooltipPosition(); ctx.fillStyle='#fff'; ctx.textAlign='center';
          ctx.shadowColor='rgba(0,0,0,.35)'; ctx.shadowBlur=3;
          ctx.font='bold 11px Arial'; ctx.fillText(p.toFixed(0)+'%',pos.x,pos.y-6);
          ctx.font='10px Arial'; ctx.fillText(praFmtN(c),pos.x,pos.y+7); ctx.shadowBlur=0; return;
        }
        if(t==='line'){ctx.font='bold 10px Arial';ctx.fillStyle='#172033';ctx.textAlign='center';ctx.fillText(praFmtN(c)+' ('+p.toFixed(1)+'%)',el.x,el.y-12);return;}
        var dep=d3?praDepthOf(chart,el):0, full=praFmtN(c)+' ('+p.toFixed(1)+'%)', cnt=praFmtN(c), pc=p.toFixed(1)+'%';
        ctx.font='bold 10px Arial'; ctx.fillStyle='#172033';
        if(horiz){ctx.textAlign='left'; ctx.fillText(o.grouped?cnt:full,el.x+6+dep,el.y-dep/2); return;}
        ctx.textAlign='center';
        var slot=o.grouped?el.width+4:el.width/0.72, x=el.x+dep/2, y=el.y-10-dep;
        if(!o.grouped&&ctx.measureText(full).width<=slot-4){ctx.fillText(full,x,y);return;}
        if(!o.grouped&&Math.max(ctx.measureText(cnt).width,ctx.measureText(pc).width)<=slot-2){
          ctx.fillText(cnt,x,y-12); ctx.font='9px Arial'; ctx.fillStyle='#64748b'; ctx.fillText(pc,x,y); return;}
        if(ctx.measureText(cnt).width<=slot)ctx.fillText(cnt,x,y);
      });
    });
    ctx.restore();
  }});
}

function praAlpha(c,a){var m=/^#([0-9a-f]{6})$/i.exec(c);if(!m)return c;var n=parseInt(m[1],16);return 'rgba('+(n>>16)+','+((n>>8)&255)+','+(n&255)+','+a+')';}
function praGrad(colors,horiz){return function(c){var ch=c.chart,a=ch.chartArea;
  var col=Array.isArray(colors)?colors[c.dataIndex%colors.length]:colors; if(!a||typeof col!=='string')return col;
  if(praIs3D())return col;
  var g=horiz?ch.ctx.createLinearGradient(a.left,0,a.right,0):ch.ctx.createLinearGradient(0,a.bottom,0,a.top);
  g.addColorStop(0,praAlpha(col,0.45)); g.addColorStop(1,col); return g;};}
var PRA_TIP={backgroundColor:'rgba(23,32,51,.92)',padding:10,cornerRadius:8,titleFont:{size:11,weight:'bold'},bodyFont:{size:11},displayColors:true,boxPadding:4};
function praTrim(l,n){n=n||30;l=String(l);return l.length>n?l.slice(0,n-1)+'…':l;}
// Wrap an axis label onto at most two lines so it is never clipped at the canvas edge.
function praWrap(l,max){max=max||18;var words=String(l).split(/\s+/),lines=[''];
  words.forEach(function(w){var cur=lines[lines.length-1];if(cur&&(cur+' '+w).length>max)lines.push(w);else lines[lines.length-1]=cur?cur+' '+w:w;});
  if(lines.length>2){lines=[lines[0],lines.slice(1).join(' ')];}
  return lines.map(function(x){return praTrim(x,max+2);});}

// Chooses the chart form from the data: few parts of a whole -> doughnut, too many columns -> horizontal bars,
// a 1-2 point line -> column chart.
function praPickType(el,type,labels,options){
  var n=labels.length, horiz=options.indexAxis==='y';
  if(options.pie&&n>=2&&n<=options.pie)return {type:'doughnut',horiz:false};
  if(type==='line'&&n<3)return {type:'bar',horiz:false};
  if(type==='bar'&&!horiz){
    var w=(el.parentElement&&el.parentElement.clientWidth)||0;
    if(w?w/Math.max(n,1)<52:n>5)return {type:'bar',horiz:true};
  }
  return {type:type,horiz:horiz};
}

function praRebuildChart(canvasId,type,labels,values,options){
  options=options||{};
  var el=document.getElementById(canvasId); if(!el||typeof Chart==='undefined')return;
  try{var old=Chart.getChart(el); if(old)old.destroy();}catch(e){}
  var pick=praPickType(el,type,labels,options); type=pick.type;
  var bg=options.bg||PRA_PAL, counts=options.counts||null, total=options.total||0;
  var isDo=type==='doughnut', horiz=pick.horiz, d3=praIs3D();
  if(isDo&&options.pieValues)values=options.pieValues;
  if(isDo&&!Array.isArray(bg))bg=PRA_PAL;
  var pctOf=function(i){return options.pcts?options.pcts[i]:(total?counts[i]/total*100:0);};
  var lab=(isDo&&counts)?labels.map(function(l,i){return praTrim(l,20)+' · '+pctOf(i).toFixed(1)+'%';}):labels;
  var catAxis={grid:{display:false},afterFit:function(sc){if(horiz)sc.width+=8;},ticks:{color:'#334155',font:{size:9},autoSkip:false,maxRotation:horiz?0:40,
    callback:function(v){return horiz?praWrap(this.getLabelForValue(v),20):praWrap(this.getLabelForValue(v),labels.length<=3?24:14);}}};
  var valAxis={beginAtZero:true,grid:{color:'#edf0f4'},border:{display:false},ticks:{color:'#64748b',font:{size:9},
    callback:function(v){return options.percent?Number(v).toFixed(0)+'%':praFmtN(v);}}};
  if(options.percent&&!isDo){var mx=Math.max.apply(null,values.concat([1]));valAxis.suggestedMax=Math.min(100,mx*1.15);}
  var scales={}; if(!isDo){scales[horiz?'x':'y']=valAxis; scales[horiz?'y':'x']=catAxis;}
  var lc=options.border||praC(1);
  var lineFill=function(c){var a=c.chart.chartArea; if(!a)return praAlpha(lc,.14);
    var g=c.chart.ctx.createLinearGradient(0,a.top,0,a.bottom); g.addColorStop(0,praAlpha(lc,.35)); g.addColorStop(1,praAlpha(lc,0)); return g;};
  var dep=d3?PRA_DEPTH:0;
  new Chart(el,{type:type,
    data:{labels:lab,datasets:[{data:values,praColors:bg,
      backgroundColor:type==='bar'?praGrad(bg,horiz):(type==='line'?lineFill:bg),hoverBackgroundColor:type==='bar'?bg:undefined,
      borderColor:type==='line'?lc:(isDo?'#fff':bg),
      borderWidth:type==='line'?3:(isDo?2:0),borderRadius:isDo?0:(d3?0:6),borderSkipped:false,maxBarThickness:44,
      pointRadius:5,pointHoverRadius:7,pointBackgroundColor:'#fff',pointBorderColor:lc,pointBorderWidth:2,
      tension:0.35,fill:type==='line',hoverOffset:isDo?10:0}]},
    options:{responsive:true,maintainAspectRatio:false,indexAxis:horiz?'y':'x',cutout:isDo?(options.solidPie?'0%':'55%'):undefined,
      animation:{duration:900,easing:'easeOutQuart'},
      layout:{padding:isDo?{top:4,left:4,right:4,bottom:4+dep}:{right:horiz?96+dep:12+dep,top:horiz?4+dep:34+dep,left:10,bottom:4}},
      plugins:{
        legend:{display:isDo||!!options.legend,position:'bottom',
          labels:{color:'#475569',font:{size:9},boxWidth:10,usePointStyle:true,pointStyle:'circle',padding:8}},
        tooltip:Object.assign({},PRA_TIP,{callbacks:{label:function(c){var i=c.dataIndex;
          if(counts){return ' '+labels[i]+': '+praFmtN(counts[i])+' ('+pctOf(i).toFixed(1)+'%)';}
          return ' '+labels[i]+': '+c.raw+(options.percent?'%':'');}}}),
        praLabels:counts?{counts:counts,total:total,pcts:isDo?null:(options.pcts||null)}:false
      },
      scales:scales}});
  return {type:type,horiz:horiz};
}

function praRebuildGroupedChart(canvasId,labels,datasets,options){
  options=options||{};
  var el=document.getElementById(canvasId); if(!el||typeof Chart==='undefined')return;
  try{var old=Chart.getChart(el); if(old)old.destroy();}catch(e){}
  var horiz=options.indexAxis==='y', d3=praIs3D(), dep=d3?PRA_DEPTH:0;
  datasets.forEach(function(d,i){var col=d.praColors||d.backgroundColor||PRA_PAL[i%6];
    d.praColors=col; d.borderRadius=d3?0:5; d.borderSkipped=false; d.maxBarThickness=34;
    d.hoverBackgroundColor=col; d.backgroundColor=praGrad(col,horiz);});
  var scales={};
  scales[horiz?'x':'y']={beginAtZero:true,grid:{color:'#edf0f4'},border:{display:false},ticks:{color:'#64748b',font:{size:9},precision:0}};
  scales[horiz?'y':'x']={grid:{display:false},afterFit:function(sc){if(horiz)sc.width+=8;},ticks:{color:'#334155',font:{size:9},autoSkip:false,maxRotation:horiz?0:30,
    callback:function(v){return horiz?praWrap(this.getLabelForValue(v),20):praWrap(this.getLabelForValue(v),14);}}};
  new Chart(el,{type:'bar',data:{labels:labels,datasets:datasets},
    options:{responsive:true,maintainAspectRatio:false,indexAxis:horiz?'y':'x',
      animation:{duration:900,easing:'easeOutQuart'},
      layout:{padding:{right:(horiz?44:8)+dep,top:(horiz?4:20)+dep}},
      plugins:{legend:{display:!!options.legend,position:'bottom',
          labels:{color:'#475569',font:{size:9},boxWidth:10,usePointStyle:true,pointStyle:'circle',
            generateLabels:function(ch){return ch.data.datasets.map(function(ds,i){var c=Array.isArray(ds.praColors)?ds.praColors[0]:ds.praColors;
              return {text:praTrim(ds.label,32),fillStyle:c,strokeStyle:c,pointStyle:'circle',hidden:!ch.isDatasetVisible(i),datasetIndex:i};});}}},
        tooltip:Object.assign({},PRA_TIP,{callbacks:{label:function(c){
          var tot=(c.dataset.data||[]).reduce(function(a,b){return a+Number(b||0);},0);
          var p=tot?c.raw/tot*100:0;
          return ' '+c.dataset.label+': '+praFmtN(c.raw)+' ('+p.toFixed(1)+'% of this series)';}}}),
        praLabels:{grouped:true}},
      scales:scales}});
}

function praUpdateCharts(){
  var N=PRA.lead_count||1;
  var sum=function(a){return a.reduce(function(x,y){return x+Number(y||0);},0);};
  var pcts=function(a){return a.map(function(v){return +(v/N*100).toFixed(2);});};
  praRebuildChart('jobLevel','bar',PRA.job_labels,PRA.job_values,{bg:PRA_PAL,counts:PRA.job_values,total:N});
  // Job-level split is a part-of-whole view -> pie, so it differs from the Job Level columns.
  praRebuildChart('jobSplit','bar',PRA.job_split_labels,pcts(PRA.job_split_values),{bg:PRA_PAL,indexAxis:'y',percent:true,
    counts:PRA.job_split_values,total:N,pie:6,pieValues:PRA.job_split_values,solidPie:true});
  praRebuildChart('jobFunctions','line',PRA.func_labels,pcts(PRA.func_values),{bg:PRA_PAL,percent:true,counts:PRA.func_values,total:N});
  praRebuildChart('industries','bar',PRA.industry_labels,pcts(PRA.industry_values),{bg:PRA_PAL,indexAxis:'y',percent:true,counts:PRA.industry_values,total:N});
  var sz=praRebuildChart('employeeSize','bar',PRA.size_labels,pcts(PRA.size_values),{bg:PRA_PAL,percent:true,counts:PRA.size_values,total:N});
  var s4=document.getElementById('s4');
  if(s4&&sz&&sz.horiz){praRenameSection(s4,'Employee Size — Column','Employee Size');}
  var dt=sum(PRA.device_values)||1;
  praRebuildChart('devices','bar',PRA.device_labels,PRA.device_values.map(function(v){return +(v/dt*100).toFixed(2);}),
    {bg:[praC(2),praC(1),praC(0),praC(3)],indexAxis:'y',percent:true,counts:PRA.device_values,total:dt,pie:4,pieValues:PRA.device_values});
  praRebuildChart('assetSplit','doughnut',PRA.asset_labels.slice(0,6),PRA.asset_values.slice(0,6),{bg:PRA_PAL,legend:true,counts:PRA.asset_values.slice(0,6),total:N});

  var indDs=PRA.industry_labels.slice(0,2).map(function(ind,i){return {label:ind,
    data:PRA.asset_names.map(function(_,a){return PRA.asset_industry_matrix?PRA.asset_industry_matrix[a][i]:0;}),
    backgroundColor:praC(i)};});
  praRebuildGroupedChart('assetIndustry',PRA.asset_names,indDs,{legend:true});
  praRebuildGroupedChart('assetSize',PRA.size_names,PRA.asset_names.map(function(a,i){return {label:a,data:PRA.asset_size_matrix[i]||[],backgroundColor:PRA_PAL[i%6]};}),{legend:true});
  praRebuildGroupedChart('assetJob',PRA.job_names,PRA.asset_names.map(function(a,i){return {label:a,data:PRA.asset_job_matrix[i]||[],backgroundColor:PRA_PAL[i%6]};}),{legend:true,indexAxis:'y'});

  [['openChart',PRA.asset_open_labels,PRA.asset_open_values],['clickChart',PRA.asset_click_labels,PRA.asset_click_values],['conversionChart',PRA.asset_conv_labels,PRA.asset_conv_values]].forEach(function(x){
    praRebuildChart(x[0],'bar',x[1],x[2],{bg:PRA_PAL,indexAxis:'y',counts:x[2],total:sum(x[2])||1});
  });
  var rates=[PRA.delivery_rate||0,PRA.open_rate||0,PRA.click_rate||0,PRA.conversion_rate||0,PRA.bounce_rate||0];
  praRebuildChart('statsChart','bar',['Delivered','Open','Clicks','Conversion','Bounce'],rates,
    {bg:[praC(2),praC(0),praC(1),praC(3),praC(4)],indexAxis:'y',percent:true,
     counts:[PRA.delivered,PRA.opens,PRA.clicks,PRA.conversion,PRA.bounced],pcts:rates});
  if(document.getElementById('praGeoChart'))praRenderGeoChart();
}
function praRenderGeoChart(){
  praRebuildChart('praGeoChart','bar',PRA.country_labels,PRA.country_values,{bg:praC(1),indexAxis:'y',counts:PRA.country_values,total:PRA.geo_total||1});
}

/* 2D / 3D switch (floating, hidden when printing). */
function praSetMode(m){PRA_MODE=m; try{localStorage.setItem('praChartMode',m);}catch(e){}
  document.querySelectorAll('#praModeSwitch button').forEach(function(b){b.classList.toggle('on',b.dataset.mode===m);});
  praUpdateCharts();}
document.addEventListener('DOMContentLoaded',function(){
  var st=document.createElement('style');
  st.textContent='#praModeSwitch{position:fixed;left:18px;bottom:18px;z-index:9999;display:flex;gap:2px;padding:3px;border-radius:999px;'+
    'background:#fff;box-shadow:0 6px 20px rgba(15,23,42,.18);border:1px solid #e2e8f0;font:600 12px Arial}'+
    '#praModeSwitch span{padding:6px 8px 6px 10px;color:#64748b}'+
    '#praModeSwitch button{border:0;background:transparent;padding:6px 12px;border-radius:999px;cursor:pointer;color:#334155;font:inherit}'+
    '#praModeSwitch button.on{background:linear-gradient(90deg,#f47b20,#f59e0b);color:#fff}'+
    '@media print{#praModeSwitch{display:none}}'+
    '.grid>.panel,.grid>div{min-width:0}.chartbox{min-width:0;overflow:hidden}.chartbox canvas{max-width:100%}';
  document.head.appendChild(st);
  var sw=document.createElement('div'); sw.id='praModeSwitch';
  sw.innerHTML='<span>Charts</span><button data-mode="2d">2D</button><button data-mode="3d">3D</button>';
  sw.addEventListener('click',function(e){var b=e.target.closest('button'); if(b)praSetMode(b.dataset.mode);});
  document.body.appendChild(sw);
  sw.querySelectorAll('button').forEach(function(b){b.classList.toggle('on',b.dataset.mode===PRA_MODE);});
});

/* ---------------- Location slide: Google Maps ---------------- */
var praGMap=null, praMarkerList={}, praInfoWin=null, praMapRequested=false, praBounds=null;
function praS5Visible(){var s=document.getElementById('s5');return !!s&&!s.classList.contains('hidden');}

function praFallbackMap(note){
  var map=document.getElementById('geoMap'); if(!map)return;
  var first=PRA.country_labels[0]||'World';
  map.innerHTML='<div style="width:100%;height:200px"><canvas id="praGeoChart"></canvas></div>'+
    '<iframe id="praGoogleMap" title="Google Maps country location" src="https://www.google.com/maps?q='+encodeURIComponent(first)+'&output=embed" '+
    'style="width:100%;height:220px;border:0" loading="lazy" referrerpolicy="no-referrer-when-downgrade"></iframe>'+
    '<div style="font-size:10px;color:#64748b;padding:4px 8px">'+praEsc(note||'Free embedded Google map. Run with --google-maps-key to get the interactive bubble map.')+'</div>';
  map.style.height='450px'; map.style.overflow='hidden';
  praRenderGeoChart();
}
function praMapFail(msg){praGMap=null; praMapRequested=true; praFallbackMap(msg);}

function praAddMarker(name,lat,lng,cnt,max){
  var m=new google.maps.Marker({position:{lat:lat,lng:lng},map:praGMap,title:name+': '+cnt+' leads',
    label:{text:String(cnt),color:'#fff',fontWeight:'700',fontSize:'12px'},
    icon:{path:google.maps.SymbolPath.CIRCLE,scale:16+24*Math.sqrt(cnt/max),fillColor:'#f47b20',fillOpacity:0.85,strokeColor:'#fff',strokeWeight:2}});
  m.addListener('click',function(){praFocusCountry(name);});
  praMarkerList[name]={marker:m,count:cnt};
  praBounds.extend({lat:lat,lng:lng});
}
function praFitAll(){
  if(!praGMap||!praBounds||praBounds.isEmpty())return;
  praGMap.fitBounds(praBounds,50);
  google.maps.event.addListenerOnce(praGMap,'idle',function(){if(praGMap.getZoom()>5)praGMap.setZoom(4);});
}
function praBuildGoogleMap(){
  var el=document.getElementById('praGMap'); if(!el||praGMap)return;
  praGMap=new google.maps.Map(el,{center:{lat:30,lng:-20},zoom:2,minZoom:2,mapTypeControl:false,streetViewControl:false,
    styles:[{featureType:'poi',stylers:[{visibility:'off'}]},{featureType:'road',stylers:[{visibility:'off'}]}]});
  praInfoWin=new google.maps.InfoWindow(); praBounds=new google.maps.LatLngBounds();
  var pts=(PRA.geo_spec&&PRA.geo_spec.points)?PRA.geo_spec.points:[];
  var max=Math.max.apply(null,PRA.country_values.concat([1]));
  var have={};
  pts.forEach(function(p){have[p[0]]=1; praAddMarker(p[0],p[1],p[2],p[3],max);});
  var geocoder=new google.maps.Geocoder();
  PRA.country_labels.forEach(function(n,i){ if(have[n])return;
    geocoder.geocode({address:n},function(r,st){ if(st==='OK'){var l=r[0].geometry.location; praAddMarker(n,l.lat(),l.lng(),PRA.country_values[i],max); praFitAll();}});});
  praFitAll();
}
function praEnsureGoogleMap(){
  if(!PRA.maps_key||!document.getElementById('praGMap'))return;
  if(window.google&&google.maps&&google.maps.Map){
    if(!praGMap)praBuildGoogleMap(); else {google.maps.event.trigger(praGMap,'resize'); praFitAll();}
    return;
  }
  if(praMapRequested)return; praMapRequested=true;
  window.gm_authFailure=function(){praMapFail('Google rejected the API key. Check that Maps JavaScript API is enabled, billing is active and the key allows this page. Showing the free embedded map instead.');};
  window.praInitGoogleMap=function(){ if(praS5Visible())praBuildGoogleMap(); };
  var sc=document.createElement('script');
  sc.src='https://maps.googleapis.com/maps/api/js?key='+encodeURIComponent(PRA.maps_key)+'&callback=praInitGoogleMap&v=weekly';
  sc.async=true; sc.onerror=function(){praMapFail('Could not load Google Maps. Showing the free embedded map instead.');};
  document.head.appendChild(sc);
}
function praFocusCountry(name){
  var N=PRA.geo_total||1;
  document.querySelectorAll('#s5 .geo-item').forEach(function(it){
    var a=it.querySelector('a'); var on=a&&a.dataset.mapQuery===name;
    it.style.background=on?'#fff4ea':''; it.style.borderLeftWidth=on?'6px':'';
  });
  var rec=praMarkerList[name];
  if(praGMap&&rec){
    praGMap.panTo(rec.marker.getPosition()); praGMap.setZoom(4);
    praInfoWin.setContent('<div style="font:13px Arial"><b>'+praEsc(name)+'</b><br>'+praFmtN(rec.count)+' leads &middot; '+(rec.count/N*100).toFixed(1)+'%</div>');
    praInfoWin.open(praGMap,rec.marker); return;
  }
  var f=document.getElementById('praGoogleMap');
  if(f)f.src='https://www.google.com/maps?q='+encodeURIComponent(name)+'&output=embed';
}

function praUpdateSlide5(){
  var s=document.getElementById('s5'); if(!s)return;
  var names=PRA.country_labels, vals=PRA.country_values, N=PRA.geo_total||1, unique=names.length;
  var mv=s.querySelector('.metric .value'); if(mv)mv.textContent=String(unique);
  var ml=s.querySelector('.metric .label'); if(ml){var t=ml.lastChild; if(t&&t.nodeType===3)t.nodeValue='Unique Countries';}
  var title=s.querySelector('.title');
  if(title)title.textContent='Location Split \u00b7 '+(unique<=3?names.join(' \u00b7 '):names.slice(0,2).join(' \u00b7 ')+' +'+(unique-2)+' more');
  var navItems=document.querySelectorAll('.slide-nav-item'); if(navItems[4]){var sm=navItems[4].querySelector('small'); if(sm)sm.textContent='Google Map';}
  var geoList=s.querySelector('.geo-list');
  if(geoList){geoList.innerHTML=PRA.geo_rows; geoList.style.maxHeight='230px'; geoList.style.overflowY='auto';
    geoList.addEventListener('click',function(e){var a=e.target.closest('a[data-map-query]'); if(!a)return; e.preventDefault(); praFocusCountry(a.dataset.mapQuery);});}
  var callout=s.querySelector('.callout');
  if(callout){
    var k=Math.min(5,unique), top=vals.slice(0,k).reduce(function(a,b){return a+b;},0);
    callout.textContent=unique?('Top '+k+' location'+(k>1?'s':'')+' ('+names.slice(0,k).join(', ')+') account for '+praFmtN(top)+' of '+praFmtN(N)+' leads ('+(top/N*100).toFixed(1)+'%).'):'No geographic data is available in the raw lead file.';
  }
  var map=document.getElementById('geoMap');
  if(map){
    if(PRA.maps_key){
      map.innerHTML='<div id="praGMap" style="width:100%;height:100%;border-radius:10px"></div>';
      map.style.height='450px'; map.style.overflow='hidden';
    } else { praFallbackMap(); }
  }
  try{window.geoInitialized=true;}catch(e){}
  try{window.initMap=function(){return true;};}catch(e){}
}

/* ---------------- Decision-maker mix replaces the made-up device split ---------------- */
function praRenameSection(root,oldT,newT){
  root.querySelectorAll('.section').forEach(function(sec){
    var c=sec.cloneNode(true); var ic=c.querySelector('.section-icon'); if(ic)ic.remove();
    if(c.textContent.trim()!==oldT)return;
    var nodes=[].filter.call(sec.childNodes,function(n){return n.nodeType===3;});
    if(nodes.length)nodes[nodes.length-1].nodeValue=newT; else sec.appendChild(document.createTextNode(newT));
  });
}
document.addEventListener('DOMContentLoaded',function(){
  if(PRA.device_title==='Decision Maker Mix'){
    var s4=document.getElementById('s4');
    if(s4){
      praRenameSection(s4,'Devices','Decision Makers vs Influencers');
      praRenameSection(s4,'Device Mix','Decision Maker Mix');
      var t=s4.querySelector('.title'); if(t)t.textContent='Industry, Company Size & Decision Maker Profile';
    }
    var nav=document.querySelectorAll('.slide-nav-item')[3]; if(nav){var b=nav.querySelector('b'); if(b)b.textContent='Industry, Decision Makers & Size';}
  }
  var s5=document.getElementById('s5');
  if(s5){ new MutationObserver(function(){ if(praS5Visible())praEnsureGoogleMap(); }).observe(s5,{attributes:true,attributeFilter:['class']}); }
});
'''

def js_dynamic_layer(data: dict[str, Any], df: pd.DataFrame, geo_spec: dict[str, Any]) -> str:
    inputs_maps_key = data.get("maps_key", "")
    # JSON data used by browser-side JS for text + charts.
    top_job = data["job_levels"][0] if data["job_levels"] else ("N/A", 0)
    top_industries = data["industries"][:2]
    top_sizes = data["company_sizes"]
    top_assets = data["assets"]

    job_labels, job_values = chart_payload(data["job_levels"])
    job_split_labels, job_split_values = chart_payload(data["job_levels"])
    func_labels, func_values = chart_payload(data["job_functions"])
    industry_labels, industry_values = chart_payload(data["industries"][:6])
    size_labels, size_values = chart_payload(top_sizes)
    asset_labels, asset_values = chart_payload(top_assets)
    decision_labels, decision_values = chart_payload(data["decisions"])
    device_labels, device_values = chart_payload(data["devices"])
    asset_open_labels, asset_open_values = chart_payload(data["asset_opens"][:6])
    asset_click_labels, asset_click_values = chart_payload(data["asset_clicks"][:6])
    asset_conv_labels, asset_conv_values = chart_payload(data["asset_conversions"][:6])

    # Industry x Asset matrix for Slide 6. The table uses its first two columns;
    # the chart can use all six top industries.
    industry_names = [x[0] for x in data["industries"][:6]]
    asset_names = [x[0] for x in top_assets]
    matrix = build_matrix(df, data["asset_col"], data["industry_col"], asset_names, industry_names)

    # Asset x employee size / job level datasets
    company_size_col = find_column(df, ["Company Size", "Employee Size", "Company Size Range"])
    job_level_col = find_column(df, ["Job Level", "Seniority", "Seniority Level"])

    size_names = [x[0] for x in data["company_sizes"][:6]]
    asset_size_matrix = build_matrix(df, data["asset_col"], company_size_col, asset_names, size_names)

    job_names = [x[0] for x in data["job_levels"][:6]]
    asset_job_matrix = build_matrix(df, data["asset_col"], job_level_col, asset_names, job_names)

    # Build list HTML for dynamic slide 5.
    geo_sorted = sorted(data["geo"], key=lambda x: x[3], reverse=True)
    geo_top = geo_sorted[:5]
    geo_other_count = sum(x[3] for x in geo_sorted[5:])
    geo_other_pct = pct(geo_other_count, data["lead_count"]) or 0

    geo_rows = "".join(
      f'<div class="geo-item"><a href="https://www.google.com/maps/search/?api=1&amp;query={quote(name)}" '
      f'data-map-query="{esc(name)}" target="_blank" rel="noopener">{esc(name)}</a>'
      f'<b>{count:,} leads · {pct_text(pct(count, data["geo_total"]))}</b></div>'
      for name, count in data["country_counts"]
    )

    # Dynamic top two industries.
    industry_split_html = ""
    if top_industries:
        cols = []
        for i, (name, count) in enumerate(top_industries):
            color = "var(--blue)" if i == 0 else "var(--orange)"
            cols.append(
                f'<div><div style="font-size:30px;color:{color};font-weight:900">{pct_text(pct(count, data["lead_count"]))}</div>'
                f'<div style="font-size:10px;color:#64748b">{esc(name)}</div></div>'
            )
        industry_split_html = "".join(cols)
    else:
        industry_split_html = '<div><div style="font-size:14px;color:#64748b;font-weight:800">No industry data</div></div>'

    # Dynamic device mix text.
    device_total = sum(count for _, count in data["devices"])
    if data["devices"]:
        device_mix_text = " &nbsp; ".join(
            f'<b style="color:{"var(--cyan)" if i == 0 else "var(--orange)"}">{pct_text(pct(count, device_total))}</b> {esc(name)} ({count:,})'
            for i, (name, count) in enumerate(data["devices"])
        )
        if data["device_mix_is_estimate"]:
            device_mix_text += '<div style="margin-top:6px;color:#64748b;font-size:10px">Illustrative allocation; device type is not present in the source workbook.</div>'
    else:
        device_mix_text = '<span style="color:#64748b">Device data not available in the raw file.</span>'

    # Asset table on slide 6.
    table_asset_rows = ""
    for i, (asset, count) in enumerate(top_assets):
        industries_for_row = matrix[i] if i < len(matrix) else [0] * len(industry_names)
        c1 = industries_for_row[0] if len(industries_for_row) >= 1 else 0
        c2 = industries_for_row[1] if len(industries_for_row) >= 2 else 0
        table_asset_rows += (
            f"<tr><td>{esc(asset)}</td><td class='num'>{c1:,}</td><td class='num'>{c2:,}</td>"
            f"<td class='num'>{count:,}</td></tr>"
        )

    industry_headers = industry_names + ["Industry 2"]
    if len(industry_names) == 1:
        industry_headers = [industry_names[0], "—"]

    # Senior audience cards — retain the three-card format.
    senior_candidates = []
    for name in ["C-Level", "Director", "Vice President"]:
        for actual_name, count in data["job_levels"]:
            if actual_name.strip().lower() == name.strip().lower():
                senior_candidates.append((actual_name, count))
                break
    if len(senior_candidates) < 3:
        for pair in data["job_levels"]:
            if pair not in senior_candidates:
                senior_candidates.append(pair)
            if len(senior_candidates) == 3:
                break

    while len(senior_candidates) < 3:
        senior_candidates.append(("N/A", 0))

    senior_cards = "".join(
        f'<div class="audience-stat"><span class="pill">{esc(name)}</span><b>{count:,}</b>'
        f'<small>{pct_text(pct(count, data["lead_count"]))}</small></div>'
        for name, count in senior_candidates[:3]
    )

    top_job_name, top_job_count = top_job
    top_asset_name = top_assets[0][0] if top_assets else "N/A"
    top_asset_count = top_assets[0][1] if top_assets else 0
    top_industry_name = top_industries[0][0] if top_industries else "N/A"
    top_industry_count = top_industries[0][1] if top_industries else 0
    top_size_name = top_sizes[0][0] if top_sizes else "N/A"
    top_size_count = top_sizes[0][1] if top_sizes else 0

    priority_locations = ", ".join(x[0] for x in geo_top[:5]) if geo_top else "No geographic data"
    priority_location_share = pct(sum(x[3] for x in geo_top), data["lead_count"]) if geo_top else None

    duration_text = (
        f'{data["business_days"]} business days'
        if data["business_days"] is not None
        else "the supplied campaign period"
    )

    # These observations intentionally use only measured lead-level / campaign inputs.
    observations = [
        f'Campaign generated {data["lead_count"]:,} leads in {duration_text}.',
        f'{esc(top_job_name)} is the largest job-level segment with {top_job_count:,} leads ({pct_text(pct(top_job_count, data["lead_count"]))}).',
        f'{esc(top_industry_name)} is the largest industry segment with {top_industry_count:,} leads ({pct_text(pct(top_industry_count, data["lead_count"]))}).',
        f'{esc(top_size_name)} is the largest company-size segment with {top_size_count:,} leads ({pct_text(pct(top_size_count, data["lead_count"]))}).',
        f'{esc(top_asset_name)} is the largest assigned asset segment with {top_asset_count:,} leads ({pct_text(pct(top_asset_count, data["lead_count"]))}).',
        f'Delivery rate is {pct_text(data["delivery_rate"])}, open rate is {pct_text(data["open_rate"])}, click rate is {pct_text(data["click_rate"])}, and conversion rate is {pct_text(data["conversion_rate"])}.',
    ]

    recommendations = [
        f'Prioritize {esc(top_job_name)} and {esc(top_industry_name)} audiences in the next campaign cycle based on the current lead mix.',
        f'Use the strongest asset segment, {esc(top_asset_name)}, as a starting point for follow-up and asset planning.',
        f'Focus geographic follow-up on {esc(priority_locations)} where the top five locations represent {pct_text(priority_location_share)} of the raw lead base.' if priority_location_share is not None else 'Add complete location data before applying geographic prioritization.',
    ]

    # Encode everything into browser JS.
    data_json = safe_json({
        "campaign_name": data["campaign_name"],
        "report_date": data["report_date"],
        "period": data.get("period", data["report_date"]),
        "palette": data.get("palette") or [],
        "start_date": data["start_date"],
        "end_date": data["end_date"],
        "prepared_by": data["prepared_by"],
        "lead_count": data["lead_count"],
        "sent": data["sent"],
        "delivered": data["delivered"],
        "opens": data["opens"],
        "clicks": data["clicks"],
        "conversion": data["conversion"],
        "bounced": data["bounced"],
        "delivery_rate": data["delivery_rate"],
        "open_rate": data["open_rate"],
        "click_rate": data["click_rate"],
        "conversion_rate": data["conversion_rate"],
        "bounce_rate": data["bounce_rate"],
        "business_days": data["business_days"],
        "job_labels": json.loads(job_labels),
        "job_values": json.loads(job_values),
        "job_split_labels": json.loads(job_split_labels),
        "job_split_values": json.loads(job_split_values),
        "func_labels": json.loads(func_labels),
        "func_values": json.loads(func_values),
        "industry_labels": json.loads(industry_labels),
        "industry_values": json.loads(industry_values),
        "size_labels": json.loads(size_labels),
        "size_values": json.loads(size_values),
        "asset_labels": json.loads(asset_labels),
        "asset_values": json.loads(asset_values),
        "unique_asset_count": data["unique_asset_count"],
        "decision_labels": json.loads(decision_labels),
        "decision_values": json.loads(decision_values),
        "device_labels": json.loads(device_labels),
        "device_values": json.loads(device_values),
        "device_mix_is_estimate": data["device_mix_is_estimate"],
        "device_title": data["device_title"],
        "maps_key": inputs_maps_key,
        "asset_open_labels": json.loads(asset_open_labels),
        "asset_open_values": json.loads(asset_open_values),
        "asset_click_labels": json.loads(asset_click_labels),
        "asset_click_values": json.loads(asset_click_values),
        "asset_conv_labels": json.loads(asset_conv_labels),
        "asset_conv_values": json.loads(asset_conv_values),
        "country_labels": [name for name, _ in data["country_counts"]],
        "country_values": [count for _, count in data["country_counts"]],
        "asset_event_note": data["asset_event_note"],
        "geo_total": data["geo_total"],
        "top_job_name": top_job_name,
        "top_job_count": top_job_count,
        "top_industry_name": top_industry_name,
        "top_industry_count": top_industry_count,
        "top_size_name": top_size_name,
        "top_size_count": top_size_count,
        "top_asset_name": top_asset_name,
        "top_asset_count": top_asset_count,
        "senior_cards": senior_cards,
        "industry_split_html": industry_split_html,
        "device_mix_text": device_mix_text,
        "geo_rows": geo_rows,
        "table_asset_rows": table_asset_rows,
        "industry_names": industry_names,
        "asset_names": asset_names,
        "size_names": size_names,
        "asset_size_matrix": asset_size_matrix,
        "job_names": job_names,
        "asset_job_matrix": asset_job_matrix,
        "observations": observations,
        "recommendations": recommendations,
        "geo_spec": geo_spec,
    })

    # Use JSON inside a script tag; escape </script> if a campaign name contains it.
    data_json = data_json.replace("</script>", "<\\/script>")

    # JS helper code. The original HTML/CSS is untouched; this layer only updates
    # the existing DOM and recreates the existing Chart.js charts.
    dynamic_layer = f"""
<!-- ==========================================================
     DYNAMIC PRA LAYER - generated by generate_dynamic_pra.py
     ========================================================== -->
<script>
const PRA = {data_json};

function praText(root, exactText, newText) {{
  const walker = document.createTreeWalker(root || document.body, NodeFilter.SHOW_TEXT);
  let node;
  while ((node = walker.nextNode())) {{
    if (node.nodeValue.trim() === exactText) {{
      node.nodeValue = node.nodeValue.replace(exactText, String(newText));
      return true;
    }}
  }}
  return false;
}}

function praAllText(root, oldText, newText) {{
  const walker = document.createTreeWalker(root || document.body, NodeFilter.SHOW_TEXT);
  let node;
  while ((node = walker.nextNode())) {{
    if (node.nodeValue.includes(oldText)) {{
      node.nodeValue = node.nodeValue.split(oldText).join(String(newText));
    }}
  }}
}}

function praNumber(v) {{
  return v === null || v === undefined ? 'N/A' : Number(v).toLocaleString('en-US');
}}

function praPct(v) {{
  return v === null || v === undefined ? 'N/A' : Number(v).toFixed(2) + '%';
}}

function praRebuildGroupedChart(containerId, labels, datasets, options={{}}) {{
  const canvas = document.getElementById(containerId);
  if (!canvas || typeof Chart === 'undefined') return;
  try {{ const old = Chart.getChart(canvas); if (old) old.destroy(); }} catch (e) {{}}
  new Chart(canvas, {{
    type:'bar',
    data:{{labels:labels,datasets:datasets}},
    options:{{
      responsive:true,
      maintainAspectRatio:false,
      indexAxis:options.indexAxis || 'x',
      plugins:{{legend:{{display:!!options.legend,position:'bottom'}}}},
      scales:{{
        x:{{ticks:{{color:'#64748b',font:{{size:9}}}},grid:{{display:options.indexAxis !== 'y'}}}},
        y:{{beginAtZero:true,ticks:{{color:'#64748b',font:{{size:9}}}},grid:{{color:'#edf0f4'}}}}
      }}
    }}
  }});
}}

function praRebuildChart(canvasId, type, labels, values, options={{}}) {{
  const el = document.getElementById(canvasId);
  if (!el || typeof Chart === 'undefined') return;
  try {{ const old = Chart.getChart(el); if (old) old.destroy(); }} catch (e) {{}}
  const background = options.bg || ['#3f5bd8','#f47b20','#12a8b8','#18a878','#7c5ce5','#e66aa4'];
  const dataset = {{
    data:values,
    backgroundColor:background,
    borderColor:options.border || background,
    borderWidth:options.borderWidth || 0,
    borderRadius:options.radius || 4,
    pointRadius:4,tension:0.28,fill:options.fill || false
  }};
  new Chart(el, {{
    type:type,
    data:{{labels:labels,datasets:[dataset]}},
    options:{{
      responsive:true,
      maintainAspectRatio:false,
      indexAxis:options.indexAxis || 'x',
      plugins:{{legend:{{display:!!options.legend,position:'bottom'}}}},
      scales:type === 'doughnut' ? {{}} : {{
        x:options.indexAxis === 'y'
          ? {{beginAtZero:true,grid:{{color:'#edf0f4'}},ticks:{{callback:function(v){{return options.percent ? Number(v).toFixed(2) + '%' : v;}}}}}}
          : {{grid:{{display:false}}}},
        y:options.indexAxis === 'y'
          ? {{grid:{{display:false}}}}
          : {{beginAtZero:true,grid:{{color:'#edf0f4'}},ticks:{{callback:function(v){{return options.percent ? Number(v).toFixed(2) + '%' : v;}}}}}}
      }}
    }}
  }});
}}

function praUpdateTopbarAndCover() {{
  document.title = 'VAIS | ' + PRA.campaign_name + ' | Program Analysis Report';

  const topH1 = document.querySelector('.brand-copy h1');
  const topPeriod = document.querySelector('.meta-block:nth-of-type(1) b');
  const topPrepared = document.querySelector('.meta-block:nth-of-type(2) b');

  if (topH1) topH1.textContent = PRA.campaign_name;
  if (topPeriod) topPeriod.textContent = PRA.period;
  if (topPrepared) topPrepared.textContent = PRA.prepared_by;

  const coverCampaign = document.querySelector('#s1 .campaign');
  const coverPrepared = document.querySelector('#s1 .prepared');
  if (coverCampaign) coverCampaign.textContent = PRA.campaign_name;
  if (coverPrepared) coverPrepared.innerHTML = 'Prepared by: ' + PRA.prepared_by + '<br/>' + PRA.report_date;
}}

function praUpdateSlide3() {{
  const s = document.getElementById('s3');
  if (!s) return;

  // Lead count
  const leadValue = s.querySelector('.metric .value');
  if (leadValue) leadValue.textContent = praNumber(PRA.lead_count);

  // Senior audience cards
  const stats = s.querySelectorAll('.audience-stat');
  const cards = [
    ...[
      ['C-Level', 'c-level'], ['Director', 'director'], ['Vice President', 'vice president']
    ]
  ];
  const raw = PRA.job_labels.map((n,i)=>({{name:n,count:PRA.job_values[i]}}));
  const preferred = ['c-level','director','vice president'];
  let selected = preferred
    .map(k => raw.find(x => x.name.toLowerCase().trim() === k))
    .filter(Boolean);

  raw.forEach(x => {{
    if (selected.length < 3 && !selected.some(y => y.name === x.name)) selected.push(x);
  }});
  while (selected.length < 3) selected.push({{name:'N/A',count:0}});

  stats.forEach((el, i) => {{
    const pair = selected[i];
    const pill = el.querySelector('.pill');
    const val = el.querySelector('b');
    const small = el.querySelector('small');
    if (pill) pill.textContent = pair.name;
    if (val) val.textContent = praNumber(pair.count);
    if (small) small.textContent = praPct((pair.count / PRA.lead_count) * 100);
  }});

  const callout = s.querySelector('.callout');
  if (callout) {{
    callout.textContent = 'The campaign generated ' + praNumber(PRA.lead_count) +
      ' leads, with ' + PRA.top_job_name + ' contacts representing the largest job-level segment at ' +
      praPct((PRA.top_job_count / PRA.lead_count) * 100) + '.';
  }}
}}

function praUpdateSlide4() {{
  const s = document.getElementById('s4');
  if (!s) return;

  const metricValue = s.querySelector('.metric .value');
  const metricNote = s.querySelector('.metric .note');
  if (metricValue) metricValue.textContent = String(PRA.industry_labels.length);
  if (metricNote) metricNote.textContent = PRA.industry_labels.slice(0,3).join(' • ') || 'No industry data';

  const industrySplit = s.querySelector('.panel .section + .grid.g2');
  const panels = s.querySelectorAll('.panel');
  const sectionTitle = panel => {{
    const section = panel.querySelector('.section');
    if (!section) return '';
    const copy = section.cloneNode(true);
    copy.querySelector('.section-icon')?.remove();
    return copy.textContent.trim();
  }};
  const industryBox = [...panels].find(x => sectionTitle(x) === 'Industry Split');
  const deviceBox = [...panels].find(x => sectionTitle(x) === 'Device Mix');

  if (industryBox) {{
    const grid = industryBox.querySelector('.grid.g2');
    if (grid) grid.innerHTML = PRA.industry_split_html;
  }}
  if (deviceBox) {{
    const wrap = deviceBox.querySelector('div[style*="font-size:13px"]');
    if (wrap) wrap.innerHTML = PRA.device_mix_text;
  }}
}}

function praUpdateSlide5() {{
  const s = document.getElementById('s5');
  if (!s) return;
  const metricValue = s.querySelector('.metric .value');
  if (metricValue) metricValue.textContent = String(PRA.geo_total);

  const geoList = s.querySelector('.geo-list');
  if (geoList) geoList.innerHTML = PRA.geo_rows;

  const callout = s.querySelector('.callout');
  if (callout) {{
    const top = PRA.geo_rows ? 'Geographic concentration is based on the top five locations in the raw lead file.' :
      'No geographic data is available in the raw lead file.';
    callout.textContent = top;
  }}

  const map = document.getElementById('geoMap');
  if (map) {{
    const firstCountry = PRA.geo_rows ? (PRA.geo_rows.match(/data-map-query="([^"]+)"/) || [])[1] : '';
    const mapQuery = encodeURIComponent(firstCountry || 'North America');
    map.innerHTML = '<div style="width:100%;height:210px"><canvas id="praGeoChart"></canvas></div>' +
      '<iframe id="praGoogleMap" title="Google Maps country location" ' +
      'src="https://www.google.com/maps?q=' + mapQuery + '&output=embed" ' +
      'style="width:100%;height:210px;border:0" loading="lazy" referrerpolicy="no-referrer-when-downgrade"></iframe>';
    map.style.height = '430px';
    map.style.overflow = 'hidden';
    const geoCanvas = document.getElementById('praGeoChart');
    if (geoCanvas && typeof Chart !== 'undefined') {{
      new Chart(geoCanvas, {{
        type:'bar',
        data:{{
          labels:PRA.country_labels,
          datasets:[{{label:'Leads',data:PRA.country_values,backgroundColor:'#f47b20',borderRadius:3}}]
        }},
        options:{{indexAxis:'y',responsive:true,maintainAspectRatio:false,
          plugins:{{legend:{{display:false}},tooltip:{{callbacks:{{label:c=>c.raw + ' leads'}}}}}},
          scales:{{x:{{beginAtZero:true,ticks:{{precision:0,color:'#64748b'}},grid:{{color:'#edf0f4'}}}},
            y:{{ticks:{{color:'#334155',font:{{size:10}}}},grid:{{display:false}}}}}}
        }}
      }});
    }}

    if (geoList) geoList.addEventListener('click', function(event) {{
      const link = event.target.closest('a[data-map-query]');
      const frame = document.getElementById('praGoogleMap');
      if (link && frame) frame.src = 'https://www.google.com/maps?q=' + encodeURIComponent(link.dataset.mapQuery) + '&output=embed';
    }});
  }}

  // Prevent the old Leaflet initializer from running on this map.
  try {{ window.geoInitialized = true; }} catch (e) {{}}
  try {{ window.initMap = function() {{ return true; }}; }} catch (e) {{}}
}}

function praUpdateSlide6() {{
  const s = document.getElementById('s6');
  if (!s) return;

  const uniqueAssets = s.querySelector('.metric .value');
  const assetNote = s.querySelector('.metric .note');
  if (uniqueAssets) uniqueAssets.textContent = String(PRA.unique_asset_count);
  if (assetNote) assetNote.textContent = 'Unique assets in source data';

  const table = s.querySelector('table tbody');
  if (table) table.innerHTML = PRA.table_asset_rows;

  const headers = s.querySelectorAll('table th');
  if (headers.length >= 4) {{
    headers[0].textContent = 'Asset';
    headers[1].textContent = PRA.industry_labels[0] || 'Industry 1';
    headers[2].textContent = PRA.industry_labels[1] || 'Industry 2';
    headers[3].textContent = 'Total';
  }}

  const chartGrid = s.querySelector('.grid.g3');
  const industryPanel = s.querySelector('#assetIndustry')?.closest('.panel');
  if (chartGrid && industryPanel) {{
    chartGrid.style.gridTemplateColumns = 'repeat(3, minmax(0, 1fr))';
    chartGrid.style.alignItems = 'start';
    industryPanel.style.gridColumn = 'auto';
    const industryChart = industryPanel.querySelector('.chartbox');
    if (industryChart) industryChart.style.height = '300px';
    if (!document.getElementById('praIndustryChartLayout')) {{
      const style = document.createElement('style');
      style.id = 'praIndustryChartLayout';
      style.textContent = '@media(max-width:760px){{#s6 .grid.g3{{grid-template-columns:1fr!important}}}}';
      document.head.appendChild(style);
    }}
  }}
}}

function praUpdateSlide7() {{
  const s = document.getElementById('s7');
  if (!s) return;
  const callout = s.querySelector('.callout');
  if (callout) {{
    callout.textContent = 'Audience and asset matrices are calculated directly from the raw lead file. The top assigned asset is ' +
      PRA.top_asset_name + ' with ' + praNumber(PRA.top_asset_count) + ' leads.';
  }}
}}

function praUpdateAssetSlide(slideId, metricValue, metricLabel, noteText, tableMode) {{
  const s = document.getElementById(slideId);
  if (!s) return;
  const value = s.querySelector('.asset-hero .value');
  const note = s.querySelector('.asset-hero .note');
  if (value) value.textContent = praNumber(metricValue);
  if (note) {{
    if ((slideId === 's8' || slideId === 's9') && !noteText) {{
      note.remove();
    }} else {{
      note.textContent = noteText;
    }}
  }}

  if (slideId === 's8' || slideId === 's9' || slideId === 's10') {{
    const layout = s.querySelector('.asset-split-layout');
    const legendPanel = s.querySelector('.asset-legend-panel');
    if (legendPanel) legendPanel.remove();
    if (layout) layout.style.gridTemplateColumns = '175px minmax(0, 1fr)';
  }}

  const tbody = s.querySelector('table tbody');
  if (tbody) {{
    const source = tableMode === 'open' ? PRA.asset_open_values :
                  tableMode === 'click' ? PRA.asset_click_values :
                  PRA.asset_conv_values;
    tbody.innerHTML = PRA.asset_labels.slice(0, source.length).map((name,i) => {{
      const share = metricValue ? (source[i] / metricValue * 100) : 0;
      const valueLabel = Number(source[i] || 0).toLocaleString('en-US');
      return '<tr><td>' + name + '</td><td>' + valueLabel + '</td><td>' + share.toFixed(2) + '%</td></tr>';
    }}).join('');
  }}

  const callout = s.querySelector('.callout');
  if (callout) callout.textContent = PRA.asset_event_note;
}}

function praUpdateSlide11() {{
  const s = document.getElementById('s11');
  if (!s) return;
  const cards = s.querySelectorAll('.stats-card');
  const values = [PRA.sent, PRA.delivered, PRA.opens, PRA.clicks, PRA.conversion, PRA.bounced];
  const notes = ['', praPct(PRA.delivery_rate), praPct(PRA.open_rate), praPct(PRA.click_rate), praPct(PRA.conversion_rate), praPct(PRA.bounce_rate)];

  cards.forEach((card, i) => {{
    const value = card.querySelector('.value');
    const note = card.querySelector('.note');
    if (value) value.textContent = praNumber(values[i]);
    if (note && i > 0) note.textContent = notes[i];
  }});

  const readout = s.querySelector('.exec-readout .callout');
  if (readout) {{
    readout.textContent =
      'Delivery rate is ' + praPct(PRA.delivery_rate) +
      ', while the campaign generated ' + praNumber(PRA.opens) +
      ' opens, ' + praNumber(PRA.clicks) +
      ' clicks and ' + praNumber(PRA.conversion) +
      ' conversions from ' + praNumber(PRA.sent) + ' sends.';
  }}
}}

function praUpdateSlide12() {{
  const s = document.getElementById('s12');
  if (!s) return;

  const list = s.querySelector('.observation-list');
  if (list) list.innerHTML = PRA.observations.map(x => '<li>' + x + '</li>').join('');

  const callouts = s.querySelectorAll('.panel:nth-of-type(2) .callout');
  PRA.recommendations.forEach((text, i) => {{
    if (callouts[i]) callouts[i].textContent = text;
  }});
}}

function praUpdateCharts() {{
  praRebuildChart('jobLevel', 'bar',
    PRA.job_labels,
    PRA.job_values,
    {{bg:['#f28a3d','#4e7fd8','#12e7d4','#18a878','#7c5ce5'], indexAxis:'x'}}
  );

  praRebuildChart('jobSplit', 'bar',
    PRA.job_split_labels,
    PRA.job_split_values.map(v => PRA.lead_count ? (v/PRA.lead_count*100) : 0),
    {{bg:['#12e7d4','#8d6de7','#f28a3d','#4e7fd8'], indexAxis:'y', percent:true}}
  );

  praRebuildChart('jobFunctions', 'line',
    PRA.func_labels,
    PRA.func_values.map(v => PRA.lead_count ? (v/PRA.lead_count*100) : 0),
    {{border:'#f47b20', bg:'rgba(244,123,32,.12)', borderWidth:3, fill:true, percent:true}}
  );

  praRebuildChart('industries', 'bar',
    PRA.industry_labels,
    PRA.industry_values.map(v => PRA.lead_count ? (v/PRA.lead_count*100) : 0),
    {{bg:['#4e7fd8','#f28a3d','#12a8b8','#18a878','#7c5ce5','#e66aa4'], indexAxis:'y', percent:true}}
  );

  praRebuildChart('employeeSize', 'bar',
    PRA.size_labels,
    PRA.size_values.map(v => PRA.lead_count ? (v/PRA.lead_count*100) : 0),
    {{bg:'#f28a3d', percent:true}}
  );

  const deviceTotal = PRA.device_values.reduce((sum, value) => sum + Number(value || 0), 0) || 1;
  const deviceShares = PRA.device_values.map(value => Number(value || 0) / deviceTotal * 100);
  praRebuildChart('devices', 'bar', PRA.device_labels, deviceShares,
    {{bg:['#12a8b8','#f47b20','#4e7fd8','#18a878','#7c5ce5','#e66aa4'],indexAxis:'y',percent:true}});

  praRebuildChart('assetSplit', 'doughnut',
    PRA.asset_labels.slice(0,6),
    PRA.asset_values.slice(0,6),
    {{bg:['#4e7fd8','#f28a3d','#12a8b8','#18a878','#7c5ce5','#e66aa4'], legend:true, percent:true}}
  );

  // Asset industry matrix
  const industryDatasets = PRA.industry_labels.slice(0,6).map((industry, i) => ({{
    label: industry,
    data: PRA.asset_names.map((_, a) => PRA.asset_industry_matrix?.[a]?.[i] || 0),
    backgroundColor: PRA_PAL[i % PRA_PAL.length]
  }}));
  praRebuildGroupedChart('assetIndustry', PRA.asset_names, industryDatasets, {{legend:true}});

  // Asset x employee size
  const sizeDatasets = PRA.asset_names.map((asset, i) => ({{
    label: asset,
    data: PRA.asset_size_matrix[i] || [],
    backgroundColor: i === 0 ? '#3f5bd8' : '#f47b20'
  }}));
  praRebuildGroupedChart('assetSize', PRA.size_names, sizeDatasets, {{legend:true}});

  // Asset x job level
  const jobDatasets = PRA.asset_names.map((asset, i) => ({{
    label: asset,
    data: PRA.asset_job_matrix[i] || [],
    backgroundColor: PRA_PAL[i % PRA_PAL.length]
  }}));
  praRebuildGroupedChart('assetJob', PRA.job_names, jobDatasets, {{legend:true,indexAxis:'y'}});

  praRebuildChart('openChart', 'bar',
    PRA.asset_open_labels,
    PRA.asset_open_values,
    {{bg:['#4e7fd8','#f28a3d','#12a8b8','#18a878','#7c5ce5','#e66aa4'], indexAxis:'y'}}
  );

  praRebuildChart('clickChart', 'bar',
    PRA.asset_click_labels,
    PRA.asset_click_values,
    {{bg:['#4e7fd8','#f28a3d','#12a8b8','#18a878','#7c5ce5','#e66aa4'], indexAxis:'y'}}
  );

  praRebuildChart('conversionChart', 'bar',
    PRA.asset_conv_labels,
    PRA.asset_conv_values,
    {{bg:['#4e7fd8','#f28a3d','#12a8b8','#18a878','#7c5ce5','#e66aa4'], indexAxis:'y'}}
  );

  praRebuildChart('statsChart', 'bar',
    ['Delivered','Open','Clicks','Conversion','Bounce'],
    [PRA.delivery_rate || 0, PRA.open_rate || 0, PRA.click_rate || 0, PRA.conversion_rate || 0, PRA.bounce_rate || 0],
    {{bg:['#12e7d4','#4e7fd8','#f28a3d','#69d86e','#8d6de7'], indexAxis:'y', percent:true}}
  );
}}

function praUpdateSlide8to10() {{
  // Open
  praUpdateAssetSlide(
    's8',
    PRA.opens,
    'Total Open Count',
    '',
    'open'
  );

  // Click
  praUpdateAssetSlide(
    's9',
    PRA.clicks,
    'Total Click Count',
    '',
    'click'
  );

  // Conversion
  praUpdateAssetSlide(
    's10',
    PRA.conversion,
    'Total Conversion Count',
    PRA.asset_conv_values.slice(0,2).join(' + ') + ' conversions',
    'conversion'
  );
}}

function praApplyAll() {{
  praUpdateTopbarAndCover();
  praUpdateSlide3();
  praUpdateSlide4();
  praUpdateSlide5();
  praUpdateSlide6();
  praUpdateSlide7();
  praUpdateSlide8to10();
  praUpdateSlide11();

  // Observations/recommendations use safe HTML strings generated by Python.
  const obs = document.querySelector('#s12 .observation-list');
  if (obs) obs.innerHTML = PRA.observations.map(x => '<li>' + x + '</li>').join('');

  const recs = document.querySelectorAll('#s12 .panel');
  const recPanel = [...recs].find(p => p.querySelector('.section')?.textContent.trim() === 'Recommendations');
  if (recPanel) {{
    const c = recPanel.querySelectorAll('.callout');
    PRA.recommendations.forEach((r,i)=>{{ if(c[i]) c[i].innerHTML = r; }});
  }}

  // Rebuild charts after the existing template charts.
  setTimeout(() => {{
    praUpdateChartsDataOnly();
  }}, 0);
}}

function praUpdateChartsDataOnly() {{
  // Add calculated matrix data which is generated in Python.
  PRA.asset_industry_matrix = {safe_json(matrix)};
  praUpdateChartsWithMatrices();
}}

function praUpdateChartsWithMatrices() {{
  // Re-run the same chart generation function but with matrix data available.
  praUpdateCharts();
}}

document.addEventListener('DOMContentLoaded', function() {{
  praApplyAll();
}});
</script>
<script>
""" + PRA_EXTRA_JS + """
</script>
"""

    return dynamic_layer


# ============================================================
# FINAL BUILDER
# ============================================================

def collect_inputs() -> dict[str, Any]:
    print("\n=== PROGRAM ANALYSIS REPORT INPUTS ===\n")
    campaign_name = input("1. Campaign Name: ").strip()
    while not campaign_name:
        print("Campaign Name cannot be blank.")
        campaign_name = input("1. Campaign Name: ").strip()

    report_date = parse_date_input("2. Report Date (e.g. 21-Jul-2026): ")
    start_date = parse_date_input("3. Start Date (e.g. 01-Jul-2026): ")
    end_date = parse_date_input("4. End Date (e.g. 31-Jul-2026): ")

    sent = input_int("5. Sent: ")
    delivered = input_int("6. Delivered: ")
    opens = input_int("7. Opens: ")
    clicks = input_int("8. Clicks: ")
    conversion = input_int("9. Conversion: ")

    print("\nDerived automatically:")
    print("Bounce = Sent - Delivered")
    print("Delivery Rate = Delivered / Sent")
    print("Open Rate = Opens / Delivered")
    print("Click Rate = Clicks / Delivered")
    print("Conversion Rate = Conversion / Delivered")
    print("Bounce Rate = Bounce / Sent\n")

    prepared_by = input("Prepared By [Quality Department]: ").strip() or "Quality Department"
    maps_key = input("Google Maps API key (press Enter to use the free embedded map): ").strip()

    return {
        "campaign_name": campaign_name,
        "report_date": report_date,
        "start_date": start_date,
        "end_date": end_date,
        "sent": sent,
        "delivered": delivered,
        "opens": opens,
        "clicks": clicks,
        "conversion": conversion,
        "prepared_by": prepared_by,
        "maps_key": maps_key,
    }


DATE_FORMATS = {
    "month_year": "%B %Y",        # September 2026
    "mon_year": "%b %Y",          # Sep 2026
    "day_month_year": "%d-%b-%Y",  # 01-Sep-2026
}

HEX_COLOR = re.compile(r"^#[0-9a-fA-F]{6}$")


def apply_display_options(data: dict[str, Any], inputs: dict[str, Any]) -> None:
    """Re-format the dates shown in the report and attach the chosen chart palette."""
    fmt_code = DATE_FORMATS.get(inputs.get("date_format") or "month_year", "%B %Y")
    shown = {}
    for key in ("report_date", "start_date", "end_date"):
        dt = pd.to_datetime(inputs[key], dayfirst=True, errors="coerce")
        shown[key] = dt.strftime(fmt_code) if pd.notna(dt) else inputs[key]
    data.update(shown)
    data["period"] = (shown["start_date"] if shown["start_date"] == shown["end_date"]
                      else f'{shown["start_date"]} – {shown["end_date"]}')
    palette = [c for c in (inputs.get("palette") or []) if HEX_COLOR.match(str(c))]
    data["palette"] = palette if len(palette) >= 2 else []


def build_report(template_path: Path, excel_path: Path, inputs: dict[str, Any],
                 output_dir: Path | None = None) -> Path:
    template = template_path.read_text(encoding="utf-8")
    df = read_raw_leads(excel_path)
    data = prepare_data(df, inputs)
    apply_display_options(data, inputs)

    geo_spec = build_plotly_geo_spec(df)
    inputs = dict(inputs)
    inputs.setdefault("maps_key", "")
    # Plotly CDN is added only because the map is embedded dynamically.
    dynamic_layer = js_dynamic_layer(data, df, geo_spec)

    # Replace any previously generated data layer while retaining the template design.
    output_html = re.sub(
      r"<!--\s*=+\s*DYNAMIC PRA LAYER.*?-->\s*(?:<script[^>]*></script>\s*)?<script>.*?</script>\s*(?:<script>\s*//\s*PRA-EXTRA.*?</script>\s*)?",
      "",
      template,
      flags=re.DOTALL,
    )

    # Add Plotly CDN script + dynamic layer immediately before the final </body>.
    marker = "</body>"
    if marker not in output_html:
        raise ValueError("The supplied HTML template does not contain </body>.")

    output_html = output_html.replace(
        marker,
        dynamic_layer + "\n" + marker,
        1
    )

    output_html = add_summary_slide(output_html, column_summaries(df), len(df))

    if inputs.get("logo_data_uri"):
        logo = inputs["logo_data_uri"]
        output_html = re.sub(r'<img src="[^"]*"([^>]*class="brand-logo")',
                             lambda m: f'<img src="{logo}"{m.group(1)}', output_html)

    # Make page title dynamic without changing layout.
    output_html = re.sub(
        r"<title>.*?</title>",
        f"<title>VAIS | {esc(data['campaign_name'])} | Program Analysis Report</title>",
        output_html,
        count=1,
        flags=re.DOTALL,
    )

    output_name = f"PRA_{slugify(data['campaign_name'])}.html"
    output_path = (output_dir or template_path.parent) / output_name
    output_path.write_text(output_html, encoding="utf-8")

    print("\n=== REPORT GENERATED ===")
    print(f"Campaign       : {data['campaign_name']}")
    print(f"Raw Leads      : {fmt(data['lead_count'])}")
    print(f"Sent           : {fmt(data['sent'])}")
    print(f"Delivered      : {fmt(data['delivered'])}")
    print(f"Bounce         : {fmt(data['bounced'])}")
    print(f"Delivery Rate  : {pct_text(data['delivery_rate'])}")
    print(f"Open Rate      : {pct_text(data['open_rate'])}")
    print(f"Click Rate     : {pct_text(data['click_rate'])}")
    print(f"Conversion Rate: {pct_text(data['conversion_rate'])}")
    print(f"Output         : {output_path}")

    return output_path


def main() -> None:
    parser = argparse.ArgumentParser(description="Generate dynamic PRA using the supplied HTML template.")
    parser.add_argument("--template", default=DEFAULT_TEMPLATE)
    parser.add_argument("--excel", default=DEFAULT_EXCEL)

    # Optional non-interactive inputs.
    parser.add_argument("--campaign")
    parser.add_argument("--report-date")
    parser.add_argument("--start-date")
    parser.add_argument("--end-date")
    parser.add_argument("--sent", type=int)
    parser.add_argument("--delivered", type=int)
    parser.add_argument("--opens", type=int)
    parser.add_argument("--clicks", type=int)
    parser.add_argument("--conversion", type=int)
    parser.add_argument("--prepared-by", default="Quality Department")
    parser.add_argument("--date-format", choices=sorted(DATE_FORMATS), default="month_year",
                        help="How dates appear in the report (default: month_year, e.g. September 2026)")
    parser.add_argument("--palette", default="",
                        help="Comma-separated chart colours, e.g. '#3f5bd8,#f47b20,#12a8b8'")
    parser.add_argument("--google-maps-key", default=os.environ.get("GOOGLE_MAPS_API_KEY", ""),
                        help="Google Maps JavaScript API key (or set GOOGLE_MAPS_API_KEY)")

    args = parser.parse_args()

    template_path = Path(args.template)
    excel_path = Path(args.excel)

    if not template_path.exists():
        raise FileNotFoundError(f"HTML template not found: {template_path}")
    if not excel_path.exists():
        raise FileNotFoundError(f"Excel file not found: {excel_path}")

    cli_complete = all(
        v is not None
        for v in [
            args.campaign,
            args.report_date,
            args.start_date,
            args.end_date,
            args.sent,
            args.delivered,
            args.opens,
            args.clicks,
            args.conversion,
        ]
    )

    if cli_complete:
        inputs = {
            "campaign_name": args.campaign,
            "report_date": pd.to_datetime(args.report_date).strftime("%d-%b-%Y"),
            "start_date": pd.to_datetime(args.start_date).strftime("%d-%b-%Y"),
            "end_date": pd.to_datetime(args.end_date).strftime("%d-%b-%Y"),
            "sent": args.sent,
            "delivered": args.delivered,
            "opens": args.opens,
            "clicks": args.clicks,
            "conversion": args.conversion,
            "prepared_by": args.prepared_by,
            "maps_key": args.google_maps_key,
        }
    else:
        inputs = collect_inputs()

    if not inputs.get("maps_key"):
        inputs["maps_key"] = args.google_maps_key
    inputs.setdefault("date_format", args.date_format)
    inputs.setdefault("palette", [c.strip() for c in args.palette.split(",") if c.strip()])

    build_report(template_path, excel_path, inputs)


if __name__ == "__main__":
    main()