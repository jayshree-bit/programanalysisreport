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
        '<div class="footer">VALASYS MEDIA</div></section>'
    )


def remove_summary_slide(html: str) -> str:
    """Removes the optional all-lead-fields summary page and its references."""
    html = re.sub(r"<style>\s*#s13 .*?</style>\s*<section class=\"slide hidden\" id=\"s13\">.*?</section>\s*", "", html, flags=re.DOTALL)
    html = re.sub(r"<button class=\"slide-nav-item\" onclick=\"goToSlide\(12\)\">.*?</button>\s*", "", html, flags=re.DOTALL)
    html = re.sub(r"<div class=\"toc-row\"><span>08</span><b>Lead Data Summary</b><span>13</span></div>", "", html)
    return html


def remove_single_asset_slides(html: str) -> str:
  """Hides multi-asset comparison slides when the report has at most one asset."""
  removed_indices = {5, 7, 8, 9}
  removed_slide_ids = {6, 8, 9, 10}

  html = re.sub(
    r'<section\b(?=[^>]*\bid="s(?:6|8|9|10)")[^>]*>.*?</section>\s*',
    "",
    html,
    flags=re.DOTALL,
  )

  def update_navigation(match: re.Match[str]) -> str:
    button = match.group(0)
    index_match = re.search(r'goToSlide\((\d+)\)', button)
    if not index_match:
      return button
    old_index = int(index_match.group(1))
    if old_index in removed_indices:
      return ""
    new_index = old_index - sum(index < old_index for index in removed_indices)
    button = re.sub(r'goToSlide\(\d+\)', f"goToSlide({new_index})", button, count=1)
    return re.sub(
      r'(<span class="thumb">)\d+(</span>)',
      rf"\g<1>{new_index + 1:02d}\g<2>",
      button,
      count=1,
    )

  html = re.sub(
    r'<button class="slide-nav-item"[^>]*>.*?</button>\s*',
    update_navigation,
    html,
    flags=re.DOTALL,
  )

  def renumber_slide(match: re.Match[str]) -> str:
    slide_id = int(match.group(2))
    contents = match.group(3)
    new_number = slide_id - sum(removed_id < slide_id for removed_id in removed_slide_ids)
    contents = re.sub(
      r'(<div class="slideNo">)\d+(</div>)',
      rf"\g<1>{new_number}\g<2>",
      contents,
      count=1,
    )
    return f'<section{match.group(1)}>{contents}</section>'

  html = re.sub(
    r'<section\b([^>]*\bid="s(\d+)"[^>]*)>(.*?)</section>',
    renumber_slide,
    html,
    flags=re.DOTALL,
  )

  removed_toc_titles = {
    "Asset-wise Open Split",
    "Asset-wise Click Split",
    "Asset-wise Conversion Split",
  }

  def rewrite_toc(match: re.Match[str]) -> str:
    rows = re.findall(r'<div class="toc-row">.*?</div>', match.group(2), flags=re.DOTALL)
    kept_rows = []
    for row in rows:
      title_match = re.search(r'<b>(.*?)</b>', row, flags=re.DOTALL)
      title = title_match.group(1).strip() if title_match else ""
      if title in removed_toc_titles:
        continue
      if title == "Asset Dashboard":
        title, section_number, page_number = "Asset Engagement", "02", "6"
      elif title == "Campaign Statistics":
        section_number, page_number = "03", "7"
      elif title == "Observations & Recommendations":
        section_number, page_number = "04", "8"
      else:
        kept_rows.append(row)
        continue

      row = re.sub(r'<b>.*?</b>', f"<b>{title}</b>", row, count=1, flags=re.DOTALL)
      span_index = 0
      spans = list(re.finditer(r'<span>.*?</span>', row, flags=re.DOTALL))

      def replace_toc_span(span_match: re.Match[str]) -> str:
        nonlocal span_index
        current = span_index
        span_index += 1
        if current == 0:
          return f"<span>{section_number}</span>"
        if current == len(spans) - 1:
          return f"<span>{page_number}</span>"
        return span_match.group(0)

      row = re.sub(r'<span>.*?</span>', replace_toc_span, row, flags=re.DOTALL)
      kept_rows.append(row)
    return match.group(1) + "".join(kept_rows) + match.group(3)

  html = re.sub(
    r'(<div class="toc">)(.*?)(</div><div class="footer">)',
    rewrite_toc,
    html,
    count=1,
    flags=re.DOTALL,
  )
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
    # Count distinct assets ignoring case/extra spaces and blank or placeholder cells,
    # so "eBook" / "ebook " or "N/A" never make a single-asset file look multi-asset.
    _asset_norm = series_clean(df, asset_col).str.lower().str.replace(r"\s+", " ", regex=True)
    _asset_norm = _asset_norm[~_asset_norm.isin(["", "nan", "none", "null", "n/a", "na", "-", "--"])]
    unique_asset_count = int(_asset_norm.nunique())
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
        "unique_industry_count": int(series_clean(df, industry_col).replace("", pd.NA).nunique()),
        "unique_country_count": len([c for c, _ in country_counts if c]),
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
    _country_names = [c for c, _ in top_counts(df, find_column(df, ["Country", "Country Name"]), 6)]

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
    lead_type_labels = [name for name, _ in data.get("lead_types", [])]
    has_bant = any(re.search(r"\bBANT\b", name, flags=re.IGNORECASE) for name in lead_type_labels)
    has_cs = any(re.search(r"\bCS\b", name, flags=re.IGNORECASE) for name in lead_type_labels)

    if has_bant and has_cs:
        engagement_observation = (
            "BANT potential customers were engaged through telephonic calls and promotional emails; "
            "CS potential customers were engaged through promotional emails."
        )
    elif has_bant:
        engagement_observation = "Potential customers were engaged through telephonic calls and promotional emails."
    elif has_cs:
        engagement_observation = "Potential customers were engaged through promotional emails."
    else:
        engagement_observation = None

    # These observations intentionally use only measured lead-level / campaign inputs.
    observations = [
        f'Campaign generated {data["lead_count"]:,} leads in {duration_text}.',
        *([engagement_observation] if engagement_observation else []),
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
        "unique_industry_count": data["unique_industry_count"],
        "unique_country_count": data["unique_country_count"],
        "country_names": _country_names,
        "asset_country_matrix": build_matrix(df, data["asset_col"], find_column(df, ["Country", "Country Name"]), asset_names, _country_names),
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


# ---------------------------------------------------------------------------
# Download-as-PDF / Download-as-PPT layer (runs in the browser, no server needed)
# ---------------------------------------------------------------------------
PDF_LAYER_START = "<!-- PRA-PDFLAYOUT-LAYER:START -->"
PDF_LAYER_END = "<!-- PRA-PDFLAYOUT-LAYER:END -->"
PDF_LAYER_CSS = r"""#praModeSwitch{display:none!important}
.pdfx .slide .kicker{display:none}
.pdfx #s1.slide .title{text-transform:uppercase;font-size:68px!important;letter-spacing:.01em;margin-bottom:14px!important}
@media(max-width:760px){.pdfx #s1.slide .title{font-size:52px!important}}
html[data-theme="dark"] .slide{background:linear-gradient(rgba(33,41,54,.93),rgba(27,35,47,.96)),repeating-linear-gradient(90deg,rgba(255,255,255,.04) 0 38px,transparent 38px 70px,rgba(255,255,255,.022) 70px 96px,transparent 96px 140px)!important;background-color:#222a36!important}
html[data-theme="dark"] .cover{background:#222833!important}
html[data-theme="dark"]{--trk:rgba(0,0,0,.32);--ink:#e6edf5}
html[data-theme="light"]{--trk:rgba(23,32,51,.09);--ink:#334155}
.pdf-grid{display:grid;gap:14px;margin-bottom:14px}
.pdf-panel h4,.pdf-kpi h4{margin:0 0 8px;text-align:center;font-size:14px;letter-spacing:.09em;text-transform:uppercase;color:var(--ink)}
.pdf-kpi{display:flex;flex-direction:column;align-items:center;justify-content:center;gap:10px;text-align:center;min-height:200px}
.pdf-kpi .ic{font-size:46px;line-height:1}
.pdf-kpi .num{font:900 54px/1 "Courier New",monospace;color:#12e7d4;text-shadow:0 0 12px rgba(18,231,212,.5)}
html[data-theme="light"] .pdf-kpi .num{color:#0e9aa7;text-shadow:none}
.pdf-cv{position:relative;width:100%}
.pdf-hbar{display:grid;grid-template-columns:130px 1fr 52px;align-items:center;gap:10px;margin:8px 0;font-size:12px;color:var(--ink);font-weight:700}
.pdf-hbar span:first-child{text-align:right}
.pdf-track{height:30px;background:var(--trk);border-radius:3px;overflow:hidden}
.pdf-fill{height:100%;border-radius:0 14px 14px 0}
.pdf-dev{display:flex;justify-content:space-around;align-items:center;height:100%;min-height:210px;text-align:center;color:var(--ink);font-weight:800}
.pdf-dev .big{font:900 34px/1.1 "Courier New",monospace;color:#12e7d4;text-shadow:0 0 10px rgba(18,231,212,.45)}
html[data-theme="light"] .pdf-dev .big{color:#0e9aa7;text-shadow:none}
.pdf-dev .ic{font-size:64px}
.pdf-stats{display:grid;grid-template-columns:repeat(6,1fr);gap:10px;margin-bottom:14px}
.pdf-stat{text-align:center;color:var(--ink)}.pdf-stat .ic{font-size:30px}.pdf-stat b{display:block;font-size:30px}.pdf-stat small{font-size:12px;font-weight:800;letter-spacing:.05em}
.pdf-obs{margin:0;padding:6px 10px;list-style:none;color:var(--ink);font-size:14px;line-height:1.75}
.pdf-obs li{padding-left:22px;position:relative;margin-bottom:6px}.pdf-obs li:before{content:"\27A2";position:absolute;left:0;color:#f47b20}
.pdf-thanks{flex:1;display:flex;align-items:center;justify-content:center;font-size:56px;font-weight:900;color:var(--ink)}
.pdf-offices{display:grid;grid-template-columns:repeat(3,1fr);gap:20px;color:var(--ink);font-size:12px;line-height:1.6;padding:0 40px 40px}
.pdf-offices b{display:block;font-size:13px}
#geoMap{height:300px!important}
.pdf-svg{color:var(--ink)}
.pdf-leg{display:flex;flex-wrap:wrap;gap:6px 16px;justify-content:center;margin-top:8px;font-size:11px;font-weight:700;color:var(--ink)}
.pdf-leg i{display:inline-block;width:10px;height:10px;margin-right:6px;vertical-align:-1px}
.pdf-panel{display:flex;flex-direction:column}.pdf-panel>.pdf-svg{margin:auto 0}
"""
PDF_LAYER_JS = r"""(function () {
  'use strict';
  function build() {
    var P = (typeof PRA !== 'undefined') ? PRA : window.PRA; if (!P) return;
    var COL = (P.palette && P.palette.length) ? P.palette : ['#4472c4', '#ed7d31', '#a5a5a5', '#ffc000', '#5b9bd5', '#70ad47', '#9e480e', '#7c5ce5'];
    var $ = function (id) { return document.getElementById(id); };
    var sum = function (a) { return a.reduce(function (x, y) { return x + (+y || 0); }, 0); };
    var pc = function (v, t, d) { return t ? (v / t * 100).toFixed(d == null ? 0 : d) : '0'; };
    var esc = function (s) { return String(s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); };
    var FONT = 'Calibri, Segoe UI, Arial, sans-serif', uid = 0;
    try { if (window.Chart && Chart.instances) Object.values(Chart.instances).forEach(function (c) { try { c.destroy(); } catch (e) {} }); } catch (e) {}
    document.documentElement.classList.add('pdfx');

    /* ---------- tiny SVG chart kit (no external libraries) ---------- */
    function svg(w, h, inner) { return '<svg class="pdf-svg" viewBox="0 0 ' + w + ' ' + h + '" width="100%" style="display:block;max-height:' + h + 'px" xmlns="http://www.w3.org/2000/svg" font-family="' + FONT + '">' + inner + '</svg>'; }
    function T(x, y, t, o) { o = o || {}; return '<text x="' + x + '" y="' + y + '" text-anchor="' + (o.a || 'middle') + '" font-size="' + (o.s || 12) + '" font-weight="' + (o.w || 700) + '" fill="' + (o.c || 'currentColor') + '">' + esc(t) + '</text>'; }
    function grad(id, c1, c2, vert) { return '<linearGradient id="' + id + '" x1="0" y1="' + (vert ? 1 : 0) + '" x2="' + (vert ? 0 : 1) + '" y2="0"><stop offset="0" stop-color="' + c1 + '"/><stop offset="1" stop-color="' + c2 + '"/></linearGradient>'; }
    function wrap(t, n) {
      t = String(t); if (t.length <= n) return [t];
      var a = '', b = ''; t.split(' ').forEach(function (x) { if (!b && (a + ' ' + x).trim().length <= n) a = (a + ' ' + x).trim(); else b = (b + ' ' + x).trim(); });
      if (b.length > n) b = b.slice(0, n - 1) + '\u2026';
      return [a || t.slice(0, n), b].filter(Boolean);
    }
    function lines(x, y, t, n, o) { var ls = wrap(t, n), s = ''; ls.forEach(function (l, k) { s += T(x, y + (ls.length > 1 ? (k ? 7 : -6) : 4), l, o); }); return s; }
    function legend(items) { return '<div class="pdf-leg">' + items.map(function (s) { return '<span><i style="background:' + s.color + '"></i>' + esc(s.name) + '</span>'; }).join('') + '</div>'; }

    function vbar(labels, vals, o) {
      o = o || {}; var W = o.w || 560, H = o.h || 210, n = labels.length || 1, pl = 10, pr = 10, pt = 26, pb = 40, id = 'g' + (++uid);
      var max = Math.max.apply(null, vals.concat([1])), slot = (W - pl - pr) / n, bw = Math.min(o.bw || slot * .5, 70), inner = '', defs = o.grad ? grad(id, o.grad[0], o.grad[1], true) : '';
      labels.forEach(function (l, i) {
        var h = vals[i] / max * (H - pt - pb), x = pl + slot * i + (slot - bw) / 2, y = H - pb - h;
        var f = o.colors ? o.colors[i] : (o.grad ? 'url(#' + id + ')' : (o.color || '#4472c4'));
        inner += '<rect x="' + x + '" y="' + y + '" width="' + bw + '" height="' + Math.max(h, 1.5) + '" rx="2" fill="' + f + '"/>' + T(x + bw / 2, y - 7, o.fmt ? o.fmt(vals[i], i) : vals[i], { s: 12 });
        wrap(l, slot < 120 ? 13 : 24).forEach(function (t, k) { inner += T(pl + slot * i + slot / 2, H - pb + 17 + k * 13, t, { s: 11 }); });
      });
      return svg(W, H, defs + inner + '<line x1="' + pl + '" x2="' + (W - pr) + '" y1="' + (H - pb) + '" y2="' + (H - pb) + '" stroke="rgba(160,174,192,.6)"/>');
    }
    function hbar(labels, vals, o) {
      o = o || {}; var n = labels.length || 1, rowH = o.rowH || 40, lw = o.lw || 150, W = o.w || 620, pr = o.pr || 56, aw = W - lw - pr, H = n * rowH + (o.axis ? 28 : 6);
      var max = o.max || Math.max.apply(null, vals.concat([1])), id = 'g' + (++uid), defs = o.grad ? grad(id, o.grad[0], o.grad[1]) : '', inner = '';
      labels.forEach(function (l, i) {
        var yc = i * rowH + rowH / 2 + 2, bh = Math.min(o.bh || 26, rowH - 8), w = Math.max(3, vals[i] / max * aw);
        var f = o.outline ? 'rgba(68,114,196,.22)' : (o.colors ? o.colors[i] : (o.grad ? 'url(#' + id + ')' : (o.color || '#4472c4')));
        inner += lines(lw - 10, yc, l, o.wrap || 26, { a: 'end', s: 11 }) +
          '<rect x="' + lw + '" y="' + (yc - bh / 2) + '" width="' + w + '" height="' + bh + '" rx="2" fill="' + f + '"' + (o.outline ? ' stroke="#3b82f6" stroke-width="1.5"' : '') + '/>' +
          (o.inside ? T(lw + w / 2, yc + 4, o.fmt ? o.fmt(vals[i], i) : vals[i], { c: '#fff', s: 13 }) : T(lw + w + 6, yc + 4, o.fmt ? o.fmt(vals[i], i) : vals[i], { a: 'start', s: 12 }));
      });
      if (o.axis) { inner += '<line x1="' + lw + '" x2="' + lw + '" y1="0" y2="' + (n * rowH) + '" stroke="rgba(160,174,192,.6)"/>'; [0, 20, 40, 60, 80, 100].forEach(function (t) { inner += T(lw + aw * t / 100, H - 6, t + '%', { s: 10, w: 600 }); }); }
      return svg(W, H, defs + inner);
    }
    function clustered(labels, series, o) {
      o = o || {}; var W = o.w || 560, H = o.h || 220, n = labels.length || 1, k = series.length || 1, pl = 10, pt = 24, pb = 40, slot = (W - 2 * pl) / n, gw = slot * .82, bw = Math.min(gw / k, 36), inner = '';
      var max = Math.max.apply(null, [1].concat.apply([], series.map(function (s) { return s.data || []; })));
      labels.forEach(function (l, i) {
        var x0 = pl + slot * i + (slot - bw * k) / 2;
        series.forEach(function (s, j) {
          var v = (s.data || [])[i] || 0, h = v / max * (H - pt - pb), x = x0 + j * bw;
          inner += '<rect x="' + x + '" y="' + (H - pb - h) + '" width="' + (bw - 2) + '" height="' + Math.max(h, v ? 1.5 : 0) + '" fill="' + s.color + '"/>' + (v ? T(x + bw / 2 - 1, H - pb - h - 6, v, { s: 11 }) : '');
        });
        wrap(l, slot < 110 ? 12 : 22).forEach(function (t, q) { inner += T(pl + slot * i + slot / 2, H - pb + 17 + q * 13, t, { s: 11 }); });
      });
      return svg(W, H, inner + '<line x1="' + pl + '" x2="' + (W - pl) + '" y1="' + (H - pb) + '" y2="' + (H - pb) + '" stroke="rgba(160,174,192,.6)"/>') + legend(series);
    }
    function stacked(rows, series, o) {
      o = o || {}; var W = o.w || 680, rowH = 46, lw = 210, pr = 20, aw = W - lw - pr, H = rows.length * rowH + 6, inner = '';
      var tots = rows.map(function (_, i) { return sum(series.map(function (s) { return s.data[i] || 0; })); }), max = Math.max.apply(null, tots.concat([1]));
      rows.forEach(function (r, i) {
        var yc = i * rowH + rowH / 2 + 2, x = lw; inner += lines(lw - 10, yc, r, 34, { a: 'end', s: 11 });
        series.forEach(function (s) { var v = s.data[i] || 0; if (!v) return; var w = v / max * aw; inner += '<rect x="' + x + '" y="' + (yc - 14) + '" width="' + w + '" height="28" fill="' + s.color + '"/>' + (w > 14 ? T(x + w / 2, yc + 4, v, { c: '#fff', s: 12 }) : ''); x += w; });
      });
      return svg(W, H, inner) + legend(series);
    }
    function donut(p, color, text) {
      var r = 62, c = 2 * Math.PI * r, d = Math.max(0, Math.min(100, p)) / 100 * c;
      return svg(240, 230, '<ellipse cx="120" cy="206" rx="80" ry="10" fill="rgba(18,231,212,.45)"/><circle cx="120" cy="105" r="' + r + '" fill="none" stroke="rgba(160,174,192,.45)" stroke-width="34"/><circle cx="120" cy="105" r="' + r + '" fill="none" stroke="' + color + '" stroke-width="34" stroke-dasharray="' + d + ' ' + (c - d) + '" transform="rotate(-90 120 105)"/>' + T(120, 112, text, { s: 20, w: 800 }));
    }
    function pie(labels, vals, colors) {
      var W = 640, H = 300, cx = 320, cy = 125, r = 100, ry = .6, tot = sum(vals) || 1, a0 = -Math.PI / 2, slices = [], lab = '';
      vals.forEach(function (v, i) {
        var a1 = a0 + v / tot * 2 * Math.PI, large = (a1 - a0) > Math.PI ? 1 : 0, x1 = r * Math.cos(a0), y1 = r * Math.sin(a0), x2 = r * Math.cos(a1), y2 = r * Math.sin(a1);
        slices.push({ d: v / tot >= .9999 ? 'M0,' + (-r) + ' A' + r + ',' + r + ' 0 1 1 0,' + r + ' A' + r + ',' + r + ' 0 1 1 0,' + (-r) + 'Z' : 'M0,0 L' + x1 + ',' + y1 + ' A' + r + ',' + r + ' 0 ' + large + ' 1 ' + x2 + ',' + y2 + 'Z', c: colors[i % colors.length] });
        var m = (a0 + a1) / 2, lx = cx + Math.cos(m) * (r + 22), ly = cy + Math.sin(m) * r * ry + (Math.sin(m) > 0 ? 24 : 0) + Math.sin(m) * 18, an = Math.cos(m) >= 0 ? 'start' : 'end';
        if (v) lab += T(lx, ly - 6, labels[i], { a: an, s: 11 }) + T(lx, ly + 8, pc(v, tot) + '%', { a: an, s: 11 });
        a0 = a1;
      });
      var depth = slices.map(function (s) { return '<path d="' + s.d + '" fill="' + s.c + '"/><path d="' + s.d + '" fill="#000" fill-opacity=".38"/>'; }).join('');
      var top = slices.map(function (s) { return '<path d="' + s.d + '" fill="' + s.c + '" stroke="rgba(255,255,255,.55)" stroke-width="1"/>'; }).join('');
      return svg(W, H, '<g transform="translate(' + cx + ',' + (cy + 16) + ') scale(1,' + ry + ')">' + depth + '</g><g transform="translate(' + cx + ',' + cy + ') scale(1,' + ry + ')">' + top + '</g>' + lab);
    }

    /* ---------- page helpers ---------- */
    var panel = function (h, body, st) { return '<div class="panel pdf-panel"' + (st ? ' style="' + st + '"' : '') + '>' + (h ? '<h4>' + h + '</h4>' : '') + body + '</div>'; };
    var kpi = function (t, ic, v) { return '<div class="panel pdf-kpi"><h4>' + t + '</h4><div class="ic">' + ic + '</div><div class="num">' + v + '</div></div>'; };
    var grid = function (cols, inner) { return '<div class="pdf-grid" style="grid-template-columns:' + cols + '">' + inner + '</div>'; };
    function slide(id, title, body) {
      var s = $(id); if (!s) return;
      var n = (s.querySelector('.slideNo') || {}).textContent || '';
      s.innerHTML = '<h2 class="title">' + title + '</h2>' + body + '<div class="footer">VALASYS MEDIA\u2122</div><div class="slideNo">' + n + '</div>';
    }
    function hrows(rows, grads) {
      return rows.map(function (r, i) { var g = grads[i % grads.length];
        return '<div class="pdf-hbar"><span>' + esc(r[0]) + '</span><div class="pdf-track"><div class="pdf-fill" style="width:' + Math.max(2, Math.min(100, r[1])) + '%;background:linear-gradient(90deg,' + g[0] + ',' + g[1] + ')"></div></div><span>' + r[2] + '</span></div>'; }).join('');
    }
    var sizeKey = function (l) { var m = String(l).replace(/,/g, '').match(/\d+/); return m ? +m[0] : 1e9; };
    var order = function (names, fn) { return names.map(function (_, i) { return i; }).sort(function (a, b) { return fn(names[a], names[b], a, b); }); };
    var acol = function (name) { var i = P.asset_names.indexOf(name); return COL[(i < 0 ? 0 : i) % COL.length]; };
    var leads = P.lead_count || sum(P.job_values || []) || 1;

    /* ---- Slide 3 ---- */
    var ji = order(P.job_labels, function (a, b, i, j) { return P.job_values[i] - P.job_values[j]; });
    var jl = ji.map(function (i) { return P.job_labels[i]; }), jv = ji.map(function (i) { return P.job_values[i]; }), jt = sum(jv) || 1;
    var split = jl.map(function (l, i) { return [l, jv[i] / jt * 100, pc(jv[i], jt) + '%']; }).reverse();
    var dm = 0;
    if (P.decision_labels && P.decision_labels.length) { P.decision_labels.forEach(function (l, i) { if (/decision/i.test(l) && !/non|not/i.test(l)) dm += P.decision_values[i]; }); jt = sum(P.decision_values) || jt; }
    else jl.forEach(function (l, i) { if (/c-?level|chief|\bc[a-z]o\b|vice|\bvp\b|president|director|head|owner|founder|decision/i.test(l)) dm += jv[i]; });
    var dmP = jt ? dm / jt * 100 : 0, ft = sum(P.func_values) || 1;
    slide('s3', 'Campaign Dashboard',
      grid('1fr 2fr 2fr', kpi('Leads Generated', '\uD83C\uDFAF', leads) + panel('Job Level', vbar(jl, jv, { w: 520, h: 200, bw: 46, grad: ['#a8801f', '#b8f5c0'] })) + panel('Job Level Split', hrows(split, [['#a21caf', '#f0f'], ['#0284c7', '#22e5ff'], ['#f59e0b', '#fde047'], ['#ea580c', '#fbbf24']]))) +
      grid('2fr 1.3fr 1.3fr', panel('Job Functions', hbar(P.func_labels, P.func_values.map(function (v) { return v / ft * 100; }), { w: 520, rowH: 70, bh: 56, lw: 130, pr: 20, max: 100, axis: true, inside: true, wrap: 16, color: '#4472c4', fmt: function (v) { return v.toFixed(2) + '%'; } })) + panel('Decision Makers', donut(dmP, '#12e7d4', dmP.toFixed(2) + '%')) + panel('Recommender', donut(100 - dmP, '#ff7a00', (100 - dmP).toFixed(2) + '%'))));

    /* ---- Slide 4 ---- */
    var ind = P.industry_labels.slice(0, 5), indv = P.industry_values.slice(0, 5);
    var si = order(P.size_labels, function (a, b) { return sizeKey(b) - sizeKey(a); });
    var sl = si.map(function (i) { return P.size_labels[i]; }), sv = si.map(function (i) { return P.size_values[i]; }), stt = sum(sv) || 1;
    var di = P.device_labels.findIndex(function (l) { return /desk|laptop|pc/i.test(l); });
    var d0 = di >= 0 ? di : 0, d1 = P.device_labels.length > 1 ? (d0 === 0 ? 1 : 0) : -1, dt = sum(P.device_values) || 1;
    var devHtml = '<div class="pdf-dev"><div><div class="ic">' + (di >= 0 ? '\uD83D\uDDA5' : '\u25D0') + '</div><div class="big">' + pc(P.device_values[d0], dt, 1) + '%</div>' + esc(P.device_labels[d0] || '') + '</div>' +
      (d1 >= 0 ? '<div><div class="ic">' + (di >= 0 ? '\uD83D\uDCF1' : '\u25D1') + '</div><div class="big">' + pc(P.device_values[d1], dt, 1) + '%</div>' + esc(P.device_labels[d1]) + '</div>' : '') + '</div>';
    slide('s4', 'Campaign Dashboard',
      grid('1fr 3fr', kpi('Unique Industries', '\uD83C\uDFED', P.unique_industry_count || P.industry_labels.length) + panel('Top 5 Industries', vbar(ind.map(function (l) { return l.toUpperCase(); }), indv, { w: 760, h: 210, bw: 34, color: '#4472c4', fmt: function (v) { return pc(v, leads) + '%'; } }))) +
      grid('1.4fr 1fr', panel('Employee Size', hbar(sl, sv, { w: 560, rowH: 34, bh: 24, lw: 80, pr: 50, grad: ['#7c3aed', '#ff3d8b'], fmt: function (v) { return pc(v, stt) + '%'; } })) + panel(P.device_title || 'Devices', devHtml)));

    /* ---- Slide 5 (keeps the Google map below) ---- */
    var s5 = $('s5');
    if (s5) {
      var t5 = s5.querySelector('.title'); if (t5) t5.textContent = 'Campaign Dashboard';
      var cl = P.country_labels.slice(0, 7), cvv = P.country_values.slice(0, 7), rest = (P.geo_total || sum(P.country_values)) - sum(cvv);
      if (rest > 0) { cl.push('Others'); cvv.push(rest); }
      var row = document.createElement('div');
      row.innerHTML = grid('1fr 3fr', kpi('Unique Geo Locations', '\uD83D\uDCCD', P.unique_country_count || P.country_labels.length) + panel('Location Split', pie(cl, cvv, ['#70ad47', '#ffc000', '#ed7d31', '#4472c4', '#5b9bd5', '#a5a5a5', '#9e480e', '#7c5ce5'])));
      if (t5) t5.after(row.firstChild);
    }

    /* ---- Slide 6 ---- */
    var ai = order(P.asset_labels, function (a, b, i, j) { return P.asset_values[i] - P.asset_values[j]; }), at = sum(P.asset_values) || 1;
    var seriesBy = function (mat) { return P.asset_names.map(function (a, i) { return { name: a, color: COL[i % COL.length], data: (mat && mat[i]) || [] }; }); };
    slide('s6', 'Asset Dashboard',
      grid('1fr 3fr', kpi('Unique Assets', '\uD83D\uDCF0', P.unique_asset_count) + panel('Asset Split', hbar(ai.map(function (i) { return P.asset_labels[i]; }).reverse(), ai.map(function (i) { return P.asset_values[i] / at * 100; }).reverse(), { w: 700, rowH: 44, lw: 260, wrap: 38, bh: 24, pr: 56, max: 100, colors: ai.map(function (i) { return acol(P.asset_labels[i]); }).reverse(), fmt: function (v) { return v.toFixed(1) + '%'; } }))) +
      grid('1fr 1fr', panel('Asset Engagement by Top Industries', clustered(P.industry_names.slice(0, 5), seriesBy(P.asset_industry_matrix))) + panel('Asset Engagement by Country', clustered(P.country_names || [], seriesBy(P.asset_country_matrix)))));

    /* ---- Slide 7 ---- */
    var szI = order(P.size_names, function (a, b) { return sizeKey(a) - sizeKey(b); });
    var sizeSeries = szI.map(function (j, k) { return { name: P.size_names[j], color: COL[k % COL.length], data: P.asset_names.map(function (_, a) { return (P.asset_size_matrix && P.asset_size_matrix[a]) ? P.asset_size_matrix[a][j] : 0; }) }; });
    var jobSeries = P.job_names.map(function (n, j) { return { name: n, color: COL[j % COL.length], data: P.asset_names.map(function (_, a) { return (P.asset_job_matrix && P.asset_job_matrix[a]) ? P.asset_job_matrix[a][j] : 0; }) }; });
    slide('s7', 'Asset Dashboard', panel('Asset Engagement by Employee Size', stacked(P.asset_names, sizeSeries)) + '<div style="height:14px"></div>' + panel('Asset Engagement by Job Level', stacked(P.asset_names, jobSeries)));

    /* ---- Slides 8-10 ---- */
    function splitSlide(id, title, word, labels, values) {
      var t = sum(values) || 1;
      slide(id, title, grid('1fr 3fr', kpi('Unique Assets', '\uD83D\uDCF0', P.unique_asset_count) + panel(word + ' Split', hbar(labels, values, { w: 700, rowH: 40, lw: 270, wrap: 40, bh: 22, pr: 50, outline: true }))) +
        panel(word + ' Percentage', vbar(labels, values, { w: 900, h: 230, bw: 120, colors: labels.map(acol), fmt: function (v) { return pc(v, t) + '%'; } })));
    }
    splitSlide('s8', 'Asset Wise Open Split', 'Open', P.asset_open_labels, P.asset_open_values);
    splitSlide('s9', 'Asset Wise Click Split', 'Click', P.asset_click_labels, P.asset_click_values);
    splitSlide('s10', 'Asset Wise Conversion Split', 'Conversion', P.asset_conv_labels, P.asset_conv_values);

    /* ---- Slide 11 ---- */
    var st = [['\u2708\uFE0F', P.sent, 'SENT'], ['\uD83D\uDCEC', P.delivered, 'DELIVERED'], ['\u2709\uFE0F', P.opens, 'OPENS'], ['\uD83D\uDDB1\uFE0F', P.clicks, 'CLICKS'], ['\u2B07\uFE0F', P.conversion, 'CONVERSION'], ['\uD83D\uDEAB', P.bounced, 'BOUNCED']];
    var rates = [['Bounce', P.bounce_rate], ['Conversion', P.conversion_rate], ['Clicks', P.click_rate], ['Open', P.open_rate], ['Delivered', P.delivery_rate]];
    slide('s11', 'Campaign Statistics',
      '<div class="panel"><div class="pdf-stats">' + st.map(function (x) { return '<div class="pdf-stat"><div class="ic">' + x[0] + '</div><b>' + Number(x[1] || 0).toLocaleString() + '</b><small>' + x[2] + '</small></div>'; }).join('') + '</div></div><div style="height:14px"></div>' +
      panel('Statistics Split', rates.map(function (r) { return '<div class="pdf-hbar"><span>' + r[0] + '</span><div class="pdf-track"><div class="pdf-fill" style="width:' + Math.max(2, Math.min(100, r[1])) + '%;background:linear-gradient(90deg,#00b050,#fff200,#f59a23)"></div></div><span>' + Number(r[1]).toFixed(1) + '%</span></div>'; }).join('')));

    /* ---- Slide 12 + Thank you ---- */
    slide('s12', 'Observations &amp; Recommendations', panel('', '<ul class="pdf-obs">' + (P.observations || []).concat(P.recommendations || []).map(function (o) { return '<li>' + o + '</li>'; }).join('') + '</ul>'));
    var deck = $('deckView');
    if (deck && !$('s14')) {
      var ty = document.createElement('section'); ty.className = 'slide hidden'; ty.id = 's14';
      ty.innerHTML = '<div class="pdf-thanks">Thank You!</div><div class="pdf-offices">' +
        '<div><b><span class="office-flag">🇺🇸</span> USA Office</b>111 Town Square Place, Suite 1203, Jersey City, NJ 07310<br><b>Ph.: +1 303-960-0264</b><br>255 S Orange Avenue, Suite 104 #2185, Orlando, FL 32801</div>' +
        '<div><b><span class="office-flag">🇦🇪</span> Dubai Office</b>Unit No: 492, DMCC Business Centre, Level No 1, Jewellery &amp; Gemplex 3, Dubai - United Arab Emirates<br><b>Ph.: +971-544570526</b></div>' +
        '<div><b><span class="office-flag">🇮🇳</span> India Office</b>801, 8th Floor, Cerebrum IT Park, B-3 Building, Kalyani Nagar, Pune - 411014</div></div>' +
        '<div class="footer">VALASYS MEDIA\u2122</div>';
      deck.appendChild(ty);
      try { slides.push(ty); } catch (e) {}
    }
    try { update(); } catch (e) {}
  }
  function start() { setTimeout(function () { try { build(); } catch (e) { console.error('PDF layout failed', e); } }, 150); }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', start); else start();
})();
"""


def pdf_layout_layer() -> str:
    """Re-lays the dashboard slides out like the reference PDF (KPI tiles, bar/donut/pie/stacked charts)."""
    return (
        PDF_LAYER_START + "\n<style>" + PDF_LAYER_CSS + "</style>\n<script>" + PDF_LAYER_JS + "</script>\n"
        + PDF_LAYER_END + "\n"
    )


THEME_LAYER_START = "<!-- PRA-THEME-LAYER:START -->"
THEME_LAYER_END = "<!-- PRA-THEME-LAYER:END -->"

THEME_LAYER_CSS = r""".pra-theme-btn{border:1px solid #dfe5ec;background:#fff;color:#344054;border-radius:10px;height:38px;padding:0 13px;font:800 11px Arial,sans-serif;letter-spacing:.04em;cursor:pointer;box-shadow:0 2px 7px rgba(16,24,40,.04);white-space:nowrap}
.pra-theme-btn:hover{border-color:#f47b20;color:#f47b20}
html[data-theme="dark"]{color-scheme:dark;--line:#2a3850;--white:#e6edf5;--muted:#9fb0c3}
html[data-theme="dark"],html[data-theme="dark"] body{background:#0e1520!important;color:#e6edf5!important}
html[data-theme="dark"] .toolbar{background:rgba(15,23,35,.98)!important;border-bottom:1px solid #243044!important;box-shadow:0 3px 16px rgba(0,0,0,.4)!important}
html[data-theme="dark"] .brand-divider{background:#2a3850}
html[data-theme="dark"] .meta-block{border-right-color:#2a3850}
html[data-theme="dark"] .meta-icon{background:rgba(244,81,50,.16)}
html[data-theme="dark"] .brand-copy h1,html[data-theme="dark"] .title,html[data-theme="dark"] .cover .title,html[data-theme="dark"] .cover .campaign,html[data-theme="dark"] .audience-stat b,html[data-theme="dark"] .section,html[data-theme="dark"] .meta-block b,html[data-theme="dark"] .brand h1{color:#f1f5f9!important}
html[data-theme="dark"] .brand-copy p,html[data-theme="dark"] .subtitle,html[data-theme="dark"] .metric .label,html[data-theme="dark"] .metric .note,html[data-theme="dark"] .meta-block small,html[data-theme="dark"] .slide-indicator,html[data-theme="dark"] .cover .prepared,html[data-theme="dark"] .slideNo,html[data-theme="dark"] .toc-row span:last-child{color:#9fb0c3!important}
html[data-theme="dark"] .slideNo{color:#fff!important}
html[data-theme="dark"] .slide{background:linear-gradient(145deg,#182233 0%,#131c2a 76%,#1b1d29 100%)!important;border-color:#26334a!important;box-shadow:0 16px 48px rgba(0,0,0,.45)!important}
html[data-theme="dark"] .cover{background:linear-gradient(135deg,#182233 0%,#131c2a 58%,#241d22 100%)!important}
html[data-theme="dark"] .slide:before{border-color:rgba(244,123,32,.14)}
html[data-theme="dark"] .panel{background:rgba(26,37,54,.92)!important;border-color:#2a3850!important;box-shadow:0 8px 22px rgba(0,0,0,.30)!important}
html[data-theme="dark"] .callout{background:rgba(244,123,32,.10)!important;color:#d7dee8!important}
html[data-theme="dark"] .table th{color:#9fb0c3!important}
html[data-theme="dark"] .table td,html[data-theme="dark"] .observation-list{color:#d0d9e4!important}
html[data-theme="dark"] .table th,html[data-theme="dark"] .table td{border-bottom-color:#263349!important}
html[data-theme="dark"] .pill,html[data-theme="dark"] .geo-item{background:#1d2a3d!important;border-color:#2a3850!important;color:#c5d0de!important}
html[data-theme="dark"] .toc-row{border-bottom-color:#263349!important}
html[data-theme="dark"] .legend,html[data-theme="dark"] .legend.left-legend .legend-item{color:#c5d0de!important}
html[data-theme="dark"] .stats-card .label{color:#c5d0de!important}
html[data-theme="dark"] .stats-card .note{color:#9fb0c3!important}
html[data-theme="dark"] .audience-stat small{color:#e6edf5!important}
html[data-theme="dark"] .progress{background:#26334a}
html[data-theme="dark"] .section .section-icon,html[data-theme="dark"] .stats-card .stat-icon{background:rgba(244,123,32,.16)!important;box-shadow:inset 0 0 0 1px rgba(244,123,32,.30)!important}
html[data-theme="dark"] .icon-btn,html[data-theme="dark"] .nav-btn,html[data-theme="dark"] .btn,html[data-theme="dark"] .pra-theme-btn{background:#1d2a3d;border-color:#2a3850;color:#d0d9e4}
html[data-theme="dark"] .icon-btn:hover,html[data-theme="dark"] .nav-btn:hover,html[data-theme="dark"] .pra-theme-btn:hover{border-color:#f47b20;color:#ff9a5a;background:#243349}
html[data-theme="dark"] #geoMap,html[data-theme="dark"] .leaflet-container{background:#1b2a39!important;border-color:#2a3850!important}
html[data-theme="dark"] [style*="color:#64748b"]{color:#9fb0c3!important}
html[data-theme="dark"] [style*="color:#344054"],html[data-theme="dark"] [style*="color:#172033"]{color:#e6edf5!important}
@media print{.pra-theme-btn{display:none!important}}
"""

THEME_LAYER_JS = r"""(function () {
  'use strict';
  var KEY = 'praTheme', root = document.documentElement;

  function saved() { try { return localStorage.getItem(KEY); } catch (e) { return null; } }
  function isDark() { return root.getAttribute('data-theme') === 'dark'; }

  function themeCharts() {
    if (!window.Chart) return;
    var dark = isDark();
    var tick = dark ? '#aab6c5' : '#64748b';
    var grid = dark ? 'rgba(255,255,255,.09)' : '#edf0f4';
    try { Chart.defaults.color = tick; } catch (e) {}
    var list = [];
    try { list = Chart.instances ? Object.values(Chart.instances) : []; } catch (e) {}
    list.forEach(function (c) {
      try {
        var o = c.options || {};
        Object.keys(o.scales || {}).forEach(function (k) {
          var s = o.scales[k];
          if (!s) return;
          s.ticks = s.ticks || {}; s.ticks.color = tick;
          s.grid = s.grid || {};
          if (s.grid.display !== false) s.grid.color = grid;
          if (s.title && s.title.display) s.title.color = tick;
        });
        if (o.plugins && o.plugins.legend && o.plugins.legend.labels) o.plugins.legend.labels.color = tick;
        c.update('none');
      } catch (e) {}
    });
  }

  function label() { return isDark() ? '\u2600 Light' : '\u263E Dark'; }

  function apply(theme, persist) {
    if (theme === 'dark') root.setAttribute('data-theme', 'dark'); else root.setAttribute('data-theme', 'light');
    if (persist) { try { localStorage.setItem(KEY, theme); } catch (e) {} }
    var b = document.getElementById('praThemeBtn');
    if (b) { b.textContent = label(); b.setAttribute('aria-pressed', String(isDark())); }
    themeCharts();
  }

  function addButton() {
    var host = document.querySelector('.report-meta');
    if (!host || document.getElementById('praThemeBtn')) return;
    var b = document.createElement('button');
    b.id = 'praThemeBtn'; b.type = 'button'; b.className = 'pra-theme-btn';
    b.title = 'Switch between light and dark mode';
    b.textContent = label();
    b.addEventListener('click', function () { apply(isDark() ? 'light' : 'dark', true); });
    host.insertBefore(b, host.querySelector('.icon-btn') || null);
  }

  // Start in the saved theme (light by default).
  apply(saved() === 'light' ? 'light' : 'dark', false);

  function init() {
    addButton();
    apply(isDark() ? 'dark' : 'light', false);
    // Charts are created at different moments by the template and the data layer, so re-theme a few times.
    [400, 1200, 2800].forEach(function (ms) { setTimeout(themeCharts, ms); });
  }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init); else init();
  window.addEventListener('load', function () { setTimeout(themeCharts, 300); });
  // Slide changes can reveal charts that were rebuilt after load.
  document.addEventListener('click', function () { setTimeout(themeCharts, 150); });
  document.addEventListener('keydown', function () { setTimeout(themeCharts, 150); });
})();
"""


def theme_layer() -> str:
    """Light / dark mode toggle. Light is the default; the choice is remembered in the browser."""
    return (
        THEME_LAYER_START + "\n"
        "<style>" + THEME_LAYER_CSS + "</style>\n"
        "<script>" + THEME_LAYER_JS + "</script>\n"
        + THEME_LAYER_END + "\n"
    )


EXPORT_LAYER_START = "<!-- PRA-EXPORT-LAYER:START -->"
EXPORT_LAYER_END = "<!-- PRA-EXPORT-LAYER:END -->"

EXPORT_LAYER_JS = r"""(function () {
  'use strict';
  var PAGE_W = 1280, MIN_H = 720, SCALE = 2, busy = false;

  function fileBase() {
    var parts = (document.title || '').split('|');
    var name = (parts[1] || 'Report').trim().replace(/[^A-Za-z0-9]+/g, '_').replace(/^_+|_+$/g, '');
    return 'PRA_' + (name || 'Report');
  }

  function addButtons() {
    var host = document.querySelector('.report-meta');
    if (!host || document.getElementById('praExportBtns')) return;
    var wrap = document.createElement('div');
    wrap.id = 'praExportBtns';
    wrap.className = 'pra-export';
    wrap.innerHTML =
      '<button type="button" data-fmt="pdf" title="Download the whole report as a PDF">\u2B07 PDF</button>' +
      '<button type="button" data-fmt="pptx" title="Download the whole report as a PowerPoint">\u2B07 PPT</button>';
    var anchor = host.querySelector('.icon-btn');
    host.insertBefore(wrap, anchor || null);
    wrap.addEventListener('click', function (e) {
      var b = e.target.closest('button[data-fmt]');
      if (b) exportReport(b.getAttribute('data-fmt'));
    });
  }

  function overlay() {
    var o = document.createElement('div');
    o.id = 'praExportOverlay';
    o.innerHTML = '<div class="pra-box"><div class="pra-spin"></div><b id="praExportMsg">Preparing\u2026</b>' +
      '<small>Please keep this tab open. It takes about 20\u201340 seconds.</small></div>';
    document.body.appendChild(o);
    return o;
  }

  function wait(ms) { return new Promise(function (r) { setTimeout(r, ms); }); }

  async function settle(el) {
    try {
      if (window.Chart && Chart.instances) {
        Object.values(Chart.instances).forEach(function (c) {
          if (c && c.canvas && el.contains(c.canvas)) {
            c.options.animation = false; c.resize(); c.update('none');
          }
        });
      }
    } catch (e) {}
    try {
      if (window.Plotly) el.querySelectorAll('.js-plotly-plot').forEach(function (p) { Plotly.Plots.resize(p); });
    } catch (e) {}
    await wait(700);
  }

  function onClone(doc) {
    // Cross-origin iframes (embedded Google map) cannot be rasterised, so show a clear placeholder instead of a blank box.
    var ink = (getComputedStyle(document.documentElement).getPropertyValue('--ink') || '').trim() || '#e6edf5';
    doc.querySelectorAll('svg.pdf-svg').forEach(function (s) { s.style.color = ink; });
    doc.querySelectorAll('iframe').forEach(function (f) {
      var d = doc.createElement('div');
      d.style.cssText = 'display:flex;align-items:center;justify-content:center;width:100%;min-height:320px;' +
        'background:#eef2f7;color:#64748b;font:600 13px Arial,sans-serif;text-align:center;border-radius:8px;padding:16px;';
      d.textContent = 'Interactive map \u2013 open the HTML report to explore locations';
      f.parentNode.replaceChild(d, f);
    });
  }

  async function captureSlides(setMsg) {
    var slides = Array.prototype.slice.call(document.querySelectorAll('#deckView .slide'));
    var visible = slides.map(function (s) { return !s.classList.contains('hidden'); });
    var scrollY = window.scrollY;
    var out = [];
    window.scrollTo(0, 0);
    try {
      for (var i = 0; i < slides.length; i++) {
        setMsg('Capturing slide ' + (i + 1) + ' of ' + slides.length + '\u2026');
        slides.forEach(function (s, k) { s.classList.toggle('hidden', k !== i); });
        var el = slides[i], saved = el.style.cssText;
        el.style.cssText += ';width:' + PAGE_W + 'px;max-width:none;margin:0;min-height:' + MIN_H + 'px;';
        await settle(el);
        var canvas = await html2canvas(el, {
          scale: SCALE, useCORS: true, backgroundColor: document.documentElement.getAttribute('data-theme') === 'dark' ? '#131c2a' : '#ffffff', logging: false,
          windowWidth: 1440, onclone: onClone
        });
        out.push({ data: canvas.toDataURL('image/jpeg', 0.92), w: canvas.width / SCALE, h: canvas.height / SCALE });
        el.style.cssText = saved;
      }
    } finally {
      slides.forEach(function (s, k) { s.classList.toggle('hidden', !visible[k]); s.style.removeProperty('width'); });
      window.dispatchEvent(new Event('resize'));
      window.scrollTo(0, scrollY);
    }
    return out;
  }

  function buildPdf(imgs, name) {
    var JsPDF = window.jspdf.jsPDF;
    var pdf = new JsPDF({ orientation: 'l', unit: 'px', format: [imgs[0].w, imgs[0].h], hotfixes: ['px_scaling'], compress: true });
    imgs.forEach(function (im, i) {
      if (i > 0) pdf.addPage([im.w, im.h], im.w >= im.h ? 'l' : 'p');
      pdf.addImage(im.data, 'JPEG', 0, 0, im.w, im.h, undefined, 'FAST');
    });
    pdf.save(name + '.pdf');
  }

  function buildPptx(imgs, name) {
    var pptx = new PptxGenJS();
    pptx.layout = 'LAYOUT_WIDE'; // 13.33 x 7.5 in
    pptx.title = (document.title || 'Program Analysis Report');
    var BW = 13.333, BH = 7.5;
    imgs.forEach(function (im) {
      var s = pptx.addSlide();
      s.background = { color: document.documentElement.getAttribute('data-theme') === 'dark' ? '131C2A' : 'FFFFFF' };
      var r = Math.min(BW / im.w, BH / im.h), w = im.w * r, h = im.h * r;
      s.addImage({ data: im.data, x: (BW - w) / 2, y: (BH - h) / 2, w: w, h: h });
    });
    return pptx.writeFile({ fileName: name + '.pptx' });
  }

  async function exportReport(fmt) {
    if (busy) return;
    if (!window.html2canvas || (fmt === 'pdf' && !window.jspdf) || (fmt === 'pptx' && !window.PptxGenJS)) {
      alert('Export libraries could not be loaded. Please check your internet connection and reload the report.');
      return;
    }
    busy = true;
    var btns = document.querySelectorAll('#praExportBtns button');
    btns.forEach(function (b) { b.disabled = true; });
    var o = overlay(), msgEl = o.querySelector('#praExportMsg');
    var setMsg = function (t) { msgEl.textContent = t; };
    try {
      var imgs = await captureSlides(setMsg);
      setMsg(fmt === 'pdf' ? 'Building PDF\u2026' : 'Building PowerPoint\u2026');
      await wait(50);
      if (fmt === 'pdf') buildPdf(imgs, fileBase()); else await buildPptx(imgs, fileBase());
    } catch (err) {
      console.error(err);
      alert('Could not create the ' + (fmt === 'pdf' ? 'PDF' : 'PowerPoint') + ': ' + (err && err.message ? err.message : err));
    } finally {
      o.remove();
      btns.forEach(function (b) { b.disabled = false; });
      busy = false;
    }
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', addButtons); else addButtons();
})();
"""

EXPORT_LAYER_CSS = """
.topbar{gap:12px!important}
.brand-large{flex:1 1 0!important;min-width:0!important}
.brand-large .brand-logo{flex:0 0 auto;width:clamp(180px,22vw,320px)!important}
.brand-copy{min-width:0!important;overflow:hidden}
.brand-copy h1,.brand-copy p{white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.report-meta{gap:10px!important}
@media(max-width:1500px){.report-meta .meta-block:first-of-type{display:none}.pra-export button,.pra-theme-btn{padding-left:10px!important;padding-right:10px!important}}
@media(max-width:1250px){.report-meta .meta-block{display:none}.brand-copy{display:none}}
.pra-export{display:flex;gap:8px;align-items:center}
.pra-export button{background:#f45132;color:#fff;border:0;border-radius:8px;padding:8px 13px;font:800 11px Arial,sans-serif;letter-spacing:.04em;cursor:pointer;box-shadow:0 4px 12px rgba(244,81,50,.25);white-space:nowrap}
.pra-export button:nth-child(2){background:#3f5bd8;box-shadow:0 4px 12px rgba(63,91,216,.25)}
.pra-export button:hover{filter:brightness(1.08)}
.pra-export button:disabled{opacity:.55;cursor:wait}
#praExportOverlay{position:fixed;inset:0;z-index:99999;background:rgba(23,32,51,.78);display:flex;align-items:center;justify-content:center;font-family:Arial,sans-serif}
#praExportOverlay .pra-box{background:#fff;border-radius:14px;padding:26px 34px;text-align:center;box-shadow:0 20px 60px rgba(0,0,0,.35);min-width:300px}
#praExportOverlay b{display:block;color:#172033;font-size:15px;margin:12px 0 6px}
#praExportOverlay small{color:#667085;font-size:11px}
.pra-spin{width:34px;height:34px;margin:0 auto;border:4px solid #e4e9ef;border-top-color:#f45132;border-radius:50%;animation:praSpin .8s linear infinite}
@keyframes praSpin{to{transform:rotate(360deg)}}
@media print{.pra-export,#praExportOverlay{display:none!important}}
"""


def export_layer() -> str:
    return (
        EXPORT_LAYER_START + "\n"
        "<style>" + EXPORT_LAYER_CSS + "</style>\n"
        '<script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script>\n'
        '<script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js"></script>\n'
        '<script src="https://cdnjs.cloudflare.com/ajax/libs/pptxgenjs/3.12.0/pptxgen.bundle.js"></script>\n'
        "<script>" + EXPORT_LAYER_JS + "</script>\n"
        + EXPORT_LAYER_END + "\n"
    )



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

    # Drop any export layer left over from an earlier generated file (it is re-added below).
    output_html = re.sub(
        re.escape(EXPORT_LAYER_START) + r".*?" + re.escape(EXPORT_LAYER_END) + r"\s*",
        "",
        output_html,
        flags=re.DOTALL,
    )

    output_html = re.sub(
        re.escape(PDF_LAYER_START) + r".*?" + re.escape(PDF_LAYER_END) + r"\s*",
        "",
        output_html,
        flags=re.DOTALL,
    )
    output_html = re.sub(
        re.escape(THEME_LAYER_START) + r".*?" + re.escape(THEME_LAYER_END) + r"\s*",
        "",
        output_html,
        flags=re.DOTALL,
    )

    # Add Plotly CDN script + dynamic layer immediately before the final </body>.
    marker = "</body>"
    if marker not in output_html:
        raise ValueError("The supplied HTML template does not contain </body>.")

    output_html = output_html.replace(
        marker,
        dynamic_layer + "\n" + theme_layer() + pdf_layout_layer() + export_layer() + marker,
        1
    )

    output_html = remove_summary_slide(output_html)
    if data["unique_asset_count"] <= 1:
      output_html = remove_single_asset_slides(output_html)

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