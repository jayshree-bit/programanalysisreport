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
    if excel_path.suffix.lower() == ".ods":
        try:
            sheets = pd.ExcelFile(excel_path, engine="odf").sheet_names
            sheet = "Sheet1" if "Sheet1" in sheets else sheets[0]
            return pd.read_excel(excel_path, engine="odf", sheet_name=sheet)
        except Exception:
            pass

    sheets = pd.ExcelFile(excel_path).sheet_names
    sheet = "Sheet1" if "Sheet1" in sheets else sheets[0]
    return pd.read_excel(excel_path, sheet_name=sheet)


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
    device_mix_is_estimate = not bool(devices)
    if device_mix_is_estimate:
      devices = [("Desktop/Laptop", 72), ("Mobile", 25), ("Other", 3)]
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
        "lead_types": lead_types,
        "asset_opens": asset_opens,
        "asset_clicks": asset_clicks,
        "asset_conversions": asset_conversions,
        "asset_event_note": asset_event_note,
        "geo": geo,
        "country_counts": country_counts,
        "geo_total": geo_total,
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


def js_dynamic_layer(data: dict[str, Any], df: pd.DataFrame, geo_spec: dict[str, Any]) -> str:
    # JSON data used by browser-side JS for text + charts.
    top_job = data["job_levels"][0] if data["job_levels"] else ("N/A", 0)
    top_industries = data["industries"][:6]
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

    # Industry x Asset matrix for Slide 6
    industry_names = [x[0] for x in top_industries]
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
            f'<b style="color:{"var(--cyan)" if i == 0 else "var(--orange)"}">{pct_text(pct(count, device_total))}</b> {esc(name)}'
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
<script src="https://cdn.plot.ly/plotly-2.35.2.min.js"></script>
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
  if (topPeriod) topPeriod.textContent = PRA.report_date;
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

  const industryPanel = [...s.querySelectorAll('.panel')].find(panel =>
    panel.querySelector('.section')?.textContent.includes('Asset Engagement by Top Industries')
  );
  const topGrid = industryPanel?.parentElement;
  if (industryPanel && topGrid?.classList.contains('g3')) {{
    topGrid.style.gridTemplateColumns = 'repeat(2, minmax(0, 1fr))';
    industryPanel.style.gridColumn = '1 / -1';
    const chartBox = industryPanel.querySelector('.chartbox');
    if (chartBox) chartBox.style.height = '430px';
    topGrid.insertAdjacentElement('afterend', industryPanel);
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
  if (note) note.textContent = noteText;

  const legend = s.querySelector('.left-legend');
  if (legend) {{
    legend.innerHTML = PRA.asset_labels.slice(0,2).map((name, i) =>
      '<div class="legend-item"><span class="dot" style="background:' +
      (i === 0 ? '#3f5bd8' : '#f47b20') + '"></span>' + name + '</div>'
    ).join('');
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
  const industryColors = ['#3f5bd8','#f47b20','#12a8b8','#18a878','#7c5ce5','#e66aa4'];
  const industryDatasets = PRA.industry_labels.map((industry, i) => ({{
    label: industry,
    data: PRA.asset_names.map((_, a) => PRA.asset_industry_matrix ? PRA.asset_industry_matrix[a][i] : 0),
    backgroundColor: industryColors[i % industryColors.length]
  }}));
  praRebuildGroupedChart('assetIndustry', PRA.asset_names, industryDatasets, {{legend:true}});

  // Asset x employee size
  const sizeColors = ['#3f5bd8','#f47b20','#12a8b8','#18a878','#7c5ce5','#e66aa4'];
  const assetOpacity = ['ff','cc','99','77','55'];
  const sizeDatasets = PRA.asset_names.map((asset, i) => ({{
    label: asset,
    data: PRA.asset_size_matrix[i] || [],
    backgroundColor: PRA.size_names.map((_, sizeIndex) =>
      sizeColors[sizeIndex % sizeColors.length] + assetOpacity[i % assetOpacity.length]
    )
  }}));
  praRebuildGroupedChart('assetSize', PRA.size_names, sizeDatasets, {{legend:true}});

  // Asset x job level
  const assetColors = ['#3f5bd8','#f47b20','#12a8b8','#18a878','#9a58bd'];
  const jobDatasets = PRA.asset_names.map((asset, i) => ({{
    label: asset,
    data: PRA.asset_job_matrix[i] || [],
    backgroundColor: assetColors[i % assetColors.length]
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
    PRA.asset_open_values.slice(0,2).join(' + ') + ' opens',
    'open'
  );

  // Click
  praUpdateAssetSlide(
    's9',
    PRA.clicks,
    'Total Click Count',
    PRA.asset_click_values.slice(0,2).join(' + ') + ' clicks',
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
    }


def build_report(template_path: Path, excel_path: Path, inputs: dict[str, Any]) -> Path:
    template = template_path.read_text(encoding="utf-8")
    df = read_raw_leads(excel_path)
    data = prepare_data(df, inputs)

    geo_spec = build_plotly_geo_spec(df)
    # Plotly CDN is added only because the map is embedded dynamically.
    dynamic_layer = js_dynamic_layer(data, df, geo_spec)

    # Replace any previously generated data layer while retaining the template design.
    output_html = re.sub(
      r"<!--\s*=+\s*DYNAMIC PRA LAYER.*?-->\s*<script[^>]*>\s*</script>\s*<script>.*?</script>\s*",
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

    # Make page title dynamic without changing layout.
    output_html = re.sub(
        r"<title>.*?</title>",
        f"<title>VAIS | {esc(data['campaign_name'])} | Program Analysis Report</title>",
        output_html,
        count=1,
        flags=re.DOTALL,
    )

    output_name = f"PRA_{slugify(data['campaign_name'])}.html"
    output_path = template_path.parent / output_name
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
        }
    else:
        inputs = collect_inputs()

    build_report(template_path, excel_path, inputs)


if __name__ == "__main__":
    main()
