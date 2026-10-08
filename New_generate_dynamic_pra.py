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
    "USA": "United States", "US": "United States", "U.S.": "United States", "U.S.A.": "United States",
    "United States of America": "United States", "America": "United States",
    "UK": "United Kingdom", "U.K.": "United Kingdom", "Great Britain": "United Kingdom", "England": "United Kingdom",
    "Scotland": "United Kingdom", "Wales": "United Kingdom", "Northern Ireland": "United Kingdom",
    "UAE": "United Arab Emirates", "The Netherlands": "Netherlands", "Holland": "Netherlands",
    "Russian Federation": "Russia", "Korea, South": "South Korea", "Republic of Korea": "South Korea", "Korea": "South Korea",
    "Korea, North": "North Korea", "Czechia": "Czech Republic", "Viet Nam": "Vietnam", "Turkiye": "Turkey",
    "Türkiye": "Turkey", "Hong Kong SAR": "Hong Kong", "Macau": "Macao", "Burma": "Myanmar",
    "Cote d'Ivoire": "Ivory Coast", "Côte d'Ivoire": "Ivory Coast", "Swaziland": "Eswatini",
    "Democratic Republic of the Congo": "DR Congo", "Congo (Kinshasa)": "DR Congo", "Congo, Democratic Republic of the": "DR Congo",
    "Republic of the Congo": "Congo", "Congo (Brazzaville)": "Congo", "Republic of Ireland": "Ireland",
    "Slovak Republic": "Slovakia", "Macedonia": "North Macedonia", "Bosnia": "Bosnia and Herzegovina",
    "Taiwan, Province of China": "Taiwan", "Palestinian Territory": "Palestine", "Cape Verde": "Cabo Verde",
}

# Approximate country centre points (lat, lon) - enough to place one bubble per country.
COUNTRY_COORDS = {
    "Afghanistan": (33.9, 67.7), "Albania": (41.2, 20.2), "Algeria": (28.0, 1.7), "Andorra": (42.5, 1.5),
    "Angola": (-11.2, 17.9), "Argentina": (-38.4, -63.6), "Armenia": (40.1, 45.0), "Australia": (-25.3, 133.8),
    "Austria": (47.5, 14.6), "Azerbaijan": (40.1, 47.6), "Bahamas": (25.0, -77.4), "Bahrain": (26.0, 50.6),
    "Bangladesh": (23.7, 90.4), "Barbados": (13.2, -59.5), "Belarus": (53.7, 27.9), "Belgium": (50.5, 4.5),
    "Belize": (17.2, -88.5), "Benin": (9.3, 2.3), "Bhutan": (27.5, 90.4), "Bolivia": (-16.3, -63.6),
    "Bosnia and Herzegovina": (43.9, 17.7), "Botswana": (-22.3, 24.7), "Brazil": (-14.2, -51.9), "Brunei": (4.5, 114.7),
    "Bulgaria": (42.7, 25.5), "Burkina Faso": (12.2, -1.6), "Burundi": (-3.4, 29.9), "Cambodia": (12.6, 104.9),
    "Cameroon": (7.4, 12.4), "Canada": (56.1304, -106.3468), "Cabo Verde": (16.0, -24.0), "Central African Republic": (6.6, 20.9),
    "Chad": (15.5, 18.7), "Chile": (-35.7, -71.5), "China": (35.9, 104.2), "Colombia": (4.6, -74.3),
    "Comoros": (-11.9, 43.9), "Congo": (-0.2, 15.8), "Costa Rica": (9.7, -83.8), "Croatia": (45.1, 15.2),
    "Cuba": (21.5, -77.8), "Cyprus": (35.1, 33.4), "Czech Republic": (49.8, 15.5), "Denmark": (56.3, 9.5),
    "Djibouti": (11.8, 42.6), "Dominican Republic": (18.7, -70.2), "DR Congo": (-4.0, 21.8), "Ecuador": (-1.8, -78.2),
    "Egypt": (26.8, 30.8), "El Salvador": (13.8, -88.9), "Equatorial Guinea": (1.7, 10.3), "Eritrea": (15.2, 39.8),
    "Estonia": (58.6, 25.0), "Eswatini": (-26.5, 31.5), "Ethiopia": (9.1, 40.5), "Fiji": (-17.7, 178.1),
    "Finland": (61.9, 25.7), "France": (46.2276, 2.2137), "Gabon": (-0.8, 11.6), "Gambia": (13.4, -15.3),
    "Georgia": (42.3, 43.4), "Germany": (51.1657, 10.4515), "Ghana": (7.9, -1.0), "Greece": (39.1, 21.8),
    "Greenland": (71.7, -42.6), "Guatemala": (15.8, -90.2), "Guinea": (9.9, -9.7), "Guinea-Bissau": (11.8, -15.2),
    "Guyana": (4.9, -58.9), "Haiti": (18.9, -72.3), "Honduras": (15.2, -86.2), "Hong Kong": (22.3, 114.2),
    "Hungary": (47.2, 19.5), "Iceland": (64.96, -19.0), "India": (20.6, 78.96), "Indonesia": (-0.8, 113.9),
    "Iran": (32.4, 53.7), "Iraq": (33.2, 43.7), "Ireland": (53.4, -8.2), "Israel": (31.0, 34.9),
    "Italy": (41.9, 12.6), "Ivory Coast": (7.5, -5.5), "Jamaica": (18.1, -77.3), "Japan": (36.2, 138.3),
    "Jordan": (30.6, 36.2), "Kazakhstan": (48.0, 66.9), "Kenya": (0.0, 37.9), "Kosovo": (42.6, 20.9),
    "Kuwait": (29.3, 47.5), "Kyrgyzstan": (41.2, 74.8), "Laos": (19.9, 102.5), "Latvia": (56.9, 24.6),
    "Lebanon": (33.9, 35.9), "Lesotho": (-29.6, 28.2), "Liberia": (6.4, -9.4), "Libya": (26.3, 17.2),
    "Liechtenstein": (47.2, 9.6), "Lithuania": (55.2, 23.9), "Luxembourg": (49.8, 6.1), "Macao": (22.2, 113.5),
    "Madagascar": (-18.8, 46.9), "Malawi": (-13.3, 34.3), "Malaysia": (4.2, 102.0), "Maldives": (3.2, 73.2),
    "Mali": (17.6, -4.0), "Malta": (35.9, 14.4), "Mauritania": (21.0, -10.9), "Mauritius": (-20.3, 57.6),
    "Mexico": (23.6, -102.6), "Moldova": (47.4, 28.4), "Monaco": (43.7, 7.4), "Mongolia": (46.9, 103.8),
    "Montenegro": (42.7, 19.4), "Morocco": (31.8, -7.1), "Mozambique": (-18.7, 35.5), "Myanmar": (21.9, 95.96),
    "Namibia": (-22.96, 18.5), "Nepal": (28.4, 84.1), "Netherlands": (52.1326, 5.2913), "New Zealand": (-40.9, 174.9),
    "Nicaragua": (12.9, -85.2), "Niger": (17.6, 8.1), "Nigeria": (9.1, 8.7), "North Korea": (40.3, 127.5),
    "North Macedonia": (41.6, 21.7), "Norway": (60.5, 8.5), "Oman": (21.5, 55.9), "Pakistan": (30.4, 69.3),
    "Palestine": (31.9, 35.2), "Panama": (8.5, -80.8), "Papua New Guinea": (-6.3, 143.96), "Paraguay": (-23.4, -58.4),
    "Peru": (-9.2, -75.0), "Philippines": (12.9, 121.8), "Poland": (51.9, 19.1), "Portugal": (39.4, -8.2),
    "Puerto Rico": (18.2, -66.6), "Qatar": (25.4, 51.2), "Romania": (45.9, 24.97), "Russia": (61.5, 105.3),
    "Rwanda": (-1.9, 29.9), "Saudi Arabia": (23.9, 45.1), "Senegal": (14.5, -14.5), "Serbia": (44.0, 21.0),
    "Seychelles": (-4.7, 55.5), "Sierra Leone": (8.5, -11.8), "Singapore": (1.35, 103.8), "Slovakia": (48.7, 19.7),
    "Slovenia": (46.2, 14.99), "Somalia": (5.2, 46.2), "South Africa": (-30.6, 22.9), "South Korea": (35.9, 127.8),
    "South Sudan": (7.9, 29.7), "Spain": (40.5, -3.7), "Sri Lanka": (7.9, 80.8), "Sudan": (12.9, 30.2),
    "Suriname": (3.9, -56.0), "Sweden": (60.1, 18.6), "Switzerland": (46.8, 8.2), "Syria": (34.8, 38.99),
    "Taiwan": (23.7, 121.0), "Tajikistan": (38.9, 71.3), "Tanzania": (-6.4, 34.9), "Thailand": (15.9, 100.99),
    "Timor-Leste": (-8.9, 125.7), "Togo": (8.6, 0.8), "Trinidad and Tobago": (10.7, -61.2), "Tunisia": (33.9, 9.5),
    "Turkey": (38.96, 35.2), "Turkmenistan": (38.97, 59.6), "Uganda": (1.4, 32.3), "Ukraine": (48.4, 31.2),
    "United Arab Emirates": (23.4, 53.8), "United Kingdom": (55.3781, -3.4360), "United States": (39.8283, -98.5795),
    "Uruguay": (-32.5, -55.8), "Uzbekistan": (41.4, 64.6), "Venezuela": (6.4, -66.6), "Vietnam": (14.1, 108.3),
    "Yemen": (15.6, 48.5), "Zambia": (-13.1, 27.8), "Zimbabwe": (-19.0, 29.2),
}
_COORD_LOOKUP = {k.lower(): v for k, v in COUNTRY_COORDS.items()}
_ALIAS_LOOKUP = {k.lower(): v for k, v in COUNTRY_ALIAS.items()}


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
    name = str(country).strip()
    name = COUNTRY_ALIAS.get(name) or _ALIAS_LOOKUP.get(name.lower()) or name
    hit = COUNTRY_COORDS.get(name) or _COORD_LOOKUP.get(name.lower())
    if hit:
        return hit
    if CountryInfo is None:
        return None
    try:
        value = CountryInfo(name).info().get("latlng")
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
    country_col = find_column(df, ["Country", "Country Name"])
    all_names = [n for n, _ in top_counts(df, country_col, len(df))]
    placed = {p[0] for p in points}
    unplaced = [n for n in all_names if n not in placed]

    if not points:
        return {
            "available": False,
            "unplaced": unplaced,
            "points": [],
            "message": "No valid geographic data was found in the raw lead file."
        }

    max_size = max(x[3] for x in points) if points else 1

    return {
        "available": True,
        "unplaced": unplaced,
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
    device_is_pct = False
    if use_decision_mix and inputs.get("desktop_pct") is not None and inputs.get("mobile_pct") is not None:
        # Desktop / Mobile usage typed into the form replaces the Decision Maker Mix card.
        devices = sorted([("Desktop", int(inputs["desktop_pct"])), ("Mobile", int(inputs["mobile_pct"]))], key=lambda x: -x[1])
        device_title, use_decision_mix, device_is_pct = "Device Usage", False, True
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
        "device_is_pct": device_is_pct,
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


PRA_EXTRA_JS = r'''// PRA-EXTRA: count + percentage labels, SVG location map slide, decision-maker mix
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

/* ---------------- Location slide: self-contained SVG bubble map ----------------
   One bubble per country in the Location Split. Colours are exactly the pie's: ranks 1-7 get their own slice
   colour and every other country shares the "Others" colour. Plain SVG = no Google key, no iframe, no internet,
   and it is captured correctly by the PDF / PowerPoint export. */
var PRA_PIE_COLORS=['#70ad47','#ffc000','#ed7d31','#4472c4','#5b9bd5','#a5a5a5','#9e480e','#7c5ce5'];
var PRA_LAND={"na":[[-168,66],[-162,70],[-156,71.3],[-141,69.6],[-128,70],[-115,68.5],[-108,68],[-95,68],[-90,69],[-85,69.5],[-82,67],[-86,64],[-93,61],[-94.5,58.7],[-91,57],[-85,55.3],[-82,52.5],[-79,51.5],[-79.5,54.5],[-77,57],[-78,62],[-72,62.3],[-65,60.5],[-62,58],[-60,55.5],[-56,52.5],[-60,50.2],[-66,50],[-70,47],[-65,49.2],[-64.5,46.2],[-61,45.7],[-66,44.3],[-70,43.5],[-70.5,41.8],[-74,40.5],[-76,37],[-75.5,35.2],[-78,33.8],[-81,31.5],[-80,27],[-80.2,25.2],[-81.8,26.2],[-82.8,28.8],[-84,30],[-86.5,30.3],[-89.5,30.2],[-91,29.3],[-94,29.6],[-97.2,27.8],[-97.5,24.5],[-97.8,22],[-96,19],[-94.5,18.2],[-91.5,18.5],[-90.5,21],[-87,21.5],[-88,18.5],[-88.5,16],[-84,15.8],[-83.2,14.5],[-83.7,11],[-81.8,9],[-79.5,9.5],[-77.4,8.6],[-77.5,8.3],[-79.5,8.8],[-81.5,8],[-83.5,8.7],[-85.7,10],[-87.5,13],[-91.5,14],[-94.5,16],[-97,15.8],[-101,17.3],[-105.5,20],[-105.5,23],[-109,25.5],[-112.5,29.5],[-114.7,31.7],[-117.1,32.5],[-118.5,34],[-120.6,34.6],[-122.5,37.5],[-124.3,40.3],[-124,46],[-124.7,48.4],[-123,49],[-127,51],[-130,54.5],[-134,58],[-139,59.8],[-146,60.7],[-152,59],[-156,57],[-162,55],[-158,58.5],[-162,60],[-165,62],[-161,64.5],[-168,65.7]],"baja":[[-114.7,31.7],[-115.8,30.5],[-114,28],[-112.2,25.5],[-110,23],[-109.5,23.4],[-111.5,26],[-113,29],[-114.7,31.7]],"sa":[[-77.3,8.5],[-75.5,10.8],[-72,12],[-71,11],[-68,10.6],[-64,10.6],[-61.5,10.5],[-60,8.5],[-57,6],[-54,5.8],[-51.5,4.2],[-50,1.5],[-48,-0.8],[-44.5,-2.4],[-40,-2.8],[-35.2,-5.5],[-35,-9],[-38.5,-13.2],[-39,-17.8],[-41,-22],[-44,-23.2],[-48.5,-26],[-48.7,-28.5],[-52,-32],[-54,-34.8],[-57,-35],[-58.4,-34.3],[-57.2,-38],[-62,-39],[-62.3,-41],[-65,-41],[-64.3,-43],[-67.5,-46],[-66,-48],[-69,-51],[-68.5,-53],[-71,-54],[-74,-52],[-75.5,-47],[-73.7,-42],[-73.5,-37],[-71.7,-33],[-71.5,-28],[-70.3,-18],[-76,-14],[-79,-8],[-81.2,-5.8],[-80,-3],[-80.2,-1],[-79.8,1.5],[-78,2.5],[-77.5,6]],"gl":[[-73,78.5],[-65,81.5],[-40,83.4],[-20,82],[-18,77],[-20,72],[-22,70],[-26,68],[-33,66.5],[-40,65],[-43,60],[-48,61],[-52,65],[-54,69],[-58,75],[-68,76.2]],"eu":[[-9.5,37],[-9.3,39],[-9,43],[-2,43.5],[-1.5,46],[-4.5,48.3],[-1.5,49.7],[1.5,50.5],[4,51.5],[8,53.8],[8.5,57],[10.5,57.7],[10.5,55],[12,54.3],[14,54],[19,54.5],[21,57],[24,57],[24,59.4],[28,59.5],[30,60],[23,60],[21.5,61],[21,63],[25,65],[22,65.8],[18,63],[17,61],[19,60],[16.5,57],[14,55.5],[12.5,56.3],[11,59],[8,58],[5.5,59],[5,62],[10,64],[14,67.5],[18,69.5],[25,71],[31,70],[33,69],[41,67],[34,66.4],[35,64.5],[38,64.5],[41,66.5],[44,66],[44,68.5],[53,68.5],[58,68.8],[60,69.8],[68,68.5],[69,73],[73,72.5],[80,73.5],[87,75],[100,76.2],[104,77.7],[113,74],[128,72],[140,72.5],[150,71.3],[160,69.5],[170,70],[180,69],[180,65],[178,64.5],[172,64.5],[170,60],[163,59.8],[163,57],[156.7,51],[156,57],[160,61],[155,59.5],[143,59.3],[137,54],[141,52.5],[140.5,48.5],[135,43.8],[131,42.5],[129.5,41],[128,39],[129.5,36],[129,35],[126.5,34.5],[126.5,37.5],[125,39.5],[121.5,39],[121.5,40.8],[118,39],[119,37.2],[122.5,37],[119.5,35],[121.5,32],[122,30],[121,28],[119,25],[116.5,22.7],[113,22],[110.5,21.2],[109.5,19.7],[108,21.5],[106.7,20],[105.8,18.5],[109,15],[109.2,11.7],[107,10.4],[105,8.6],[104.8,10.2],[103,11],[100.5,13.3],[99.5,11],[100,8.5],[102,6],[103.5,4],[103.5,1.3],[101,2.8],[100.3,5.5],[98.3,8.3],[98.5,10.5],[98.5,13.5],[97.6,16.5],[95.3,15.8],[94.3,18],[92.2,21.5],[91.5,22.8],[90,22],[87,21.5],[86.5,19.5],[84,18],[82,16.5],[80.2,15.5],[80,13],[79.8,10.3],[78,8.5],[77,8],[76,10],[75,12.5],[73.5,16],[72.8,19],[72.7,21.7],[70.5,20.8],[69,22.3],[70.5,23.1],[68.5,23.7],[67,24.8],[66.5,25.4],[61.5,25.2],[57.5,25.7],[56.5,27],[54,26.7],[51.5,27.9],[50,30],[48.8,30.2],[48.5,28.5],[50.5,26.2],[51.5,24.5],[54,24.2],[56,26.2],[56.5,24.5],[58.7,23.5],[59.8,22.5],[58.5,20.5],[55.5,17.5],[52,16.2],[48,14],[45,12.8],[43.3,12.7],[42.8,15],[41,19],[39,21.5],[37,25],[35,28],[34.8,29.5],[34.3,31.2],[35,33],[36,34.5],[36,36.5],[34,36.2],[31,36.8],[28.5,36.7],[27,38.5],[26.5,40.3],[24,40.5],[23,39.5],[24,38],[22.5,36.5],[21.5,37],[21,39],[19.5,41],[19,42],[15.5,45],[13.7,45.2],[12.3,45.4],[12.5,44],[14,42.5],[16,41.8],[18.5,40.2],[17,39],[16.5,38],[15.7,38.2],[16,40],[14,40.8],[12,42],[10.5,43.5],[8.8,44.4],[7,43.7],[4,43.5],[3,42.5],[0.5,40.5],[-0.3,38.5],[-2,36.7],[-5.3,36],[-6.3,36.8],[-8.8,37.2]],"af":[[-17,21],[-16.5,24],[-13,27.7],[-10,29.5],[-9.8,31.5],[-6.8,34],[-5.9,35.8],[-2,35.1],[3,36.8],[10,37.3],[11,35],[10.2,33.5],[15,32.3],[20,32],[20,30.5],[24,32],[29,31],[32.3,31.3],[32.6,29.9],[33.5,28],[35,24],[37.2,21],[38.5,18],[41,14.5],[43.2,12.7],[44.5,10.5],[51,11.8],[51,10],[48,5],[44,1],[41.5,-1.8],[39.2,-5],[39,-8],[40.5,-11],[40.5,-15],[37,-17.8],[35,-20],[35.5,-24],[32.8,-26],[32.5,-28.5],[30,-31.5],[27,-33.7],[22,-34.2],[19,-34.8],[17.7,-32],[15.2,-27],[14.5,-22.5],[11.8,-17.2],[13.5,-12],[13,-9],[12,-5],[9,-1],[9.5,3.5],[8,4.5],[5,5.8],[2,6.3],[-2,4.8],[-5,5.2],[-7.7,4.4],[-10.5,6.3],[-13.2,8.3],[-15.2,11.2],[-16.8,13],[-17.5,14.7],[-16.5,16.5],[-16.2,19.5]],"au":[[113.5,-22],[114.5,-26],[115,-33.8],[118,-35],[123,-34],[126,-32.2],[131,-31.5],[135,-34.8],[138,-35.2],[140,-37.5],[144,-38.5],[147,-38.8],[150,-37.3],[153,-31],[153.2,-27],[150.5,-22.5],[146.2,-19],[145.4,-14.8],[143.5,-14],[142.3,-10.8],[141.5,-13],[139.5,-17.5],[136.5,-15.8],[135.5,-12],[132.5,-11.4],[130.5,-12.5],[129.2,-15],[126,-14],[122.3,-17.5],[121,-19.6],[116.5,-20.7]],"uk":[[-5.5,50],[1.3,51.2],[1.7,52.8],[0,53.5],[-1.5,55.5],[-2,57.5],[-4,57.7],[-3,58.6],[-5.2,58.6],[-6,56.5],[-4.8,55],[-3,54.8],[-3.2,53.3],[-4.7,52.7],[-5.2,51.7],[-3,51.4],[-4.5,50.4]],"ie":[[-10,51.8],[-6,52.2],[-6,54.2],[-8,55.3],[-10,54],[-9.5,52.5]],"is":[[-24,65.5],[-22,66.4],[-16,66.5],[-13.5,65],[-18,63.5],[-22.5,63.8]],"jp":[[130.8,31],[132,34],[135,34.5],[136.8,34.3],[140,35],[141,38.5],[142,40.5],[141.2,41.5],[140,40],[139.5,38],[137,37],[136,36],[133,35.5],[131,34.5],[130,33.5]],"hk":[[140,42],[141.5,45.4],[145.5,43.3],[143,42],[141,42]],"tw":[[120.2,23],[121.5,25.2],[122,24.5],[120.8,22]],"lz":[[120.5,18.5],[122.2,18.4],[121.5,15],[124,13],[121.5,13.8],[120.6,14.2],[119.9,16.3]],"md":[[122,7],[125.5,9.8],[126.5,7],[125.5,5.8],[123,7.3]],"bo":[[109,1.5],[109.5,-0.5],[110.5,-3],[114,-4],[116,-3.5],[116.5,-1],[117.8,1],[119,5],[117,7],[115.5,5],[113,3],[111,1.8]],"su":[[95.3,5.6],[98,4],[100.5,2],[104,-1],[106,-3],[105.7,-5.9],[102,-4],[99.5,-1],[97,2.5]],"ja":[[105.2,-6.8],[108,-6.3],[111,-6.5],[114.4,-7.7],[114.5,-8.7],[110,-8.3],[106.5,-7.4]],"sl":[[119.5,-5.5],[120.5,-3],[121.5,-1],[120,1],[124.8,1.5],[123,0.5],[121.2,-1.8],[123,-4],[122,-5],[120.5,-5.6]],"ng":[[131,-0.8],[134,-0.7],[138,-1.8],[141,-2.6],[145,-4.5],[147.5,-6],[150.5,-10.5],[147,-10],[143.5,-8.5],[141,-9.2],[138,-8.4],[138.5,-7],[135,-4.5],[132.5,-3.3]],"nzn":[[172.7,-34.5],[175,-37],[178.5,-37.7],[177,-39.5],[175.2,-41.5],[174.8,-39.3],[173.8,-39.2],[174.8,-37]],"nzs":[[172.7,-40.5],[174.3,-41.7],[173,-43.5],[171,-44.5],[169,-46.6],[166.5,-46],[168.5,-44],[171.5,-41.8]],"mg":[[49.3,-12],[50.4,-15.5],[49.5,-17.5],[47.3,-24.8],[45,-25.6],[43.3,-22],[44.4,-16.5],[47,-15]],"lk":[[79.8,9.8],[81.8,7.5],[81,6],[80,6.2],[79.8,8]],"cu":[[-85,21.9],[-82,23.1],[-77,21.5],[-74.2,20.3],[-77.5,19.9],[-80,21.7]],"hi":[[-74.4,19.8],[-71.7,19.9],[-68.4,18.6],[-70,18.2],[-73.5,18.2]],"nf":[[-59.3,47.7],[-56,51.5],[-53,47],[-55.5,46.9]],"baf":[[-80,73.5],[-68,70],[-62,67],[-65,63],[-72,64.5],[-77,66],[-85,70.5]],"vic":[[-118,70],[-105,73],[-101,69],[-112,68.5]],"ell":[[-90,76.5],[-75,79],[-62,82.5],[-80,83],[-92,81],[-95,77]],"nz2":[[51.5,71.5],[56,75],[68,77],[60,75],[55,71.5]]};
var PRA_WATER={"black":[[28,41.2],[29,45],[31,46.6],[33.5,46],[33,44.6],[36.5,45.3],[38,47],[39.5,43.5],[41.5,41.7],[37,41],[33,42],[29,41.2]],"casp":[[47,45],[49,46.5],[53,46.5],[53.5,44],[51,41.5],[53,40],[54,37.5],[51,36.8],[49,38.5],[49,40.5],[47.5,43]]};
function praS5Visible(){var s=document.getElementById('s5');return !!s&&!s.classList.contains('hidden');}
function praGeoColor(i){return PRA_PIE_COLORS[i<7?i:7];}
function praTextOn(hex){var m=/^#([0-9a-f]{6})$/i.exec(hex||'');if(!m)return '#fff';var n=parseInt(m[1],16);
  return (0.299*(n>>16)+0.587*((n>>8)&255)+0.114*(n&255))>165?'#1e293b':'#fff';}

function praRenderGeoMap(){
  var host=document.getElementById('geoMap'); if(!host)return;
  var names=PRA.country_labels||[], vals=PRA.country_values||[], N=PRA.geo_total||1;
  var pts={}; ((PRA.geo_spec&&PRA.geo_spec.points)||[]).forEach(function(p){pts[String(p[0]).toLowerCase()]={lat:p[1],lon:p[2]};});
  var rows=[], missing=((PRA.geo_spec&&PRA.geo_spec.unplaced)||[]).slice();
  names.forEach(function(n,i){var p=pts[String(n).toLowerCase()];
    if(p)rows.push({name:n,v:vals[i],i:i,lat:p.lat,lon:p.lon}); else if(missing.indexOf(n)<0)missing.push(n);});
  host.innerHTML=''; host.style.overflow='hidden';
  if(!rows.length){host.innerHTML='<div style="padding:24px;font:13px Arial;color:#64748b">No mappable countries were found in the raw lead file.</div>';return;}

  /* frame the map around the plotted countries (whole world when they are far apart) */
  var lons=rows.map(function(r){return r.lon;}), lats=rows.map(function(r){return r.lat;});
  var minLo=Math.min.apply(null,lons), maxLo=Math.max.apply(null,lons), minLa=Math.min.apply(null,lats), maxLa=Math.max.apply(null,lats);
  var ASPECT=2.4, sLon=Math.max((maxLo-minLo)*1.4+30,80), sLat=Math.max((maxLa-minLa)*1.4+20,40);
  if(sLon/sLat<ASPECT)sLon=sLat*ASPECT; else sLat=sLon/ASPECT;
  if(sLat>142){sLat=142;sLon=Math.min(360,sLat*ASPECT);}
  var x0=Math.min(Math.max((minLo+maxLo)/2-sLon/2,-180),180-sLon);
  var top=Math.min(84,Math.max((minLa+maxLa)/2+sLat/2,-58+sLat));
  var W=1000, S=W/sLon, MH=sLat*S;
  var X=function(lon){return (lon-x0)*S;}, Y=function(lat){return (top-lat)*S;};
  var d=function(poly){return 'M'+poly.map(function(p){return X(p[0]).toFixed(1)+','+Y(p[1]).toFixed(1);}).join('L')+'Z';};

  var grid='';
  for(var lo=Math.ceil(x0/30)*30;lo<=x0+sLon;lo+=30)grid+='<line x1="'+X(lo).toFixed(1)+'" y1="0" x2="'+X(lo).toFixed(1)+'" y2="'+MH.toFixed(1)+'"/>';
  for(var la=Math.ceil((top-sLat)/30)*30;la<=top;la+=30)grid+='<line x1="0" y1="'+Y(la).toFixed(1)+'" x2="'+W+'" y2="'+Y(la).toFixed(1)+'"/>';
  var holes=Object.keys(PRA_WATER).map(function(k){return d(PRA_WATER[k]);}).join('');
  var land=Object.keys(PRA_LAND).map(function(k){return '<path fill-rule="evenodd" d="'+d(PRA_LAND[k])+(k==='eu'?holes:'')+'"/>';}).join('');

  var max=Math.max.apply(null,rows.map(function(r){return r.v;}).concat([1]));
  rows.forEach(function(r){r.rad=11+17*Math.sqrt(r.v/max); r.cx=X(r.lon); r.cy=Y(r.lat); r.x0=r.cx; r.y0=r.cy;});
  /* nearby countries (e.g. UK / Netherlands / France / Germany) would sit on top of each other: nudge them
     apart just enough to stay readable, never more than ~45 map units from their true position */
  for(var it=0;it<80;it++){var moved=false;
    for(var a=0;a<rows.length;a++)for(var b=a+1;b<rows.length;b++){
      var A=rows[a],B=rows[b],dx=B.cx-A.cx,dy=B.cy-A.cy,dist=Math.sqrt(dx*dx+dy*dy)||0.01,need=(A.rad+B.rad)*0.92;
      if(dist<need){var push=(need-dist)/2,ux=dx/dist,uy=dy/dist; if(dist<0.5){ux=1;uy=0;}
        A.cx-=ux*push;A.cy-=uy*push;B.cx+=ux*push;B.cy+=uy*push;moved=true;}}
    rows.forEach(function(r){var ox=r.cx-r.x0,oy=r.cy-r.y0,od=Math.sqrt(ox*ox+oy*oy); if(od>45){r.cx=r.x0+ox/od*45;r.cy=r.y0+oy/od*45;}});
    if(!moved)break;}
  var dots=rows.slice().sort(function(a,b){return b.v-a.v;}).map(function(r){
    var col=praGeoColor(r.i), rad=r.rad, cx=r.cx, cy=r.cy, pc=(r.v/N*100).toFixed(1);
    return '<g class="pra-geo-dot" data-name="'+praEsc(r.name)+'" data-count="'+r.v+'" style="cursor:pointer">'+
      '<title>'+praEsc(r.name)+': '+praFmtN(r.v)+' leads ('+pc+'%)</title>'+
      '<circle cx="'+cx.toFixed(1)+'" cy="'+cy.toFixed(1)+'" r="'+rad.toFixed(1)+'" fill="'+col+'" fill-opacity=".92" stroke="#fff" stroke-width="2"/>'+
      (rad>=14?'<text x="'+cx.toFixed(1)+'" y="'+(cy+5).toFixed(1)+'" text-anchor="middle" font-size="15" font-weight="700" fill="'+praTextOn(col)+'" style="pointer-events:none">'+praFmtN(r.v)+'</text>':'')+
      '</g>';}).join('');

  /* legend = the pie's legend: top 7 + Others */
  var items=names.slice(0,7).map(function(n,i){return {t:n+' \u00b7 '+praFmtN(vals[i]),c:praGeoColor(i)};});
  if(names.length>7){var rest=vals.slice(7).reduce(function(a,b){return a+b;},0); items.push({t:'Others ('+(names.length-7)+') \u00b7 '+praFmtN(rest),c:praGeoColor(7)});}
  var lx=14,ly=MH+24,leg='';
  items.forEach(function(it){var w=24+it.t.length*7.4+18; if(lx+w>W-8){lx=14;ly+=22;}
    leg+='<circle cx="'+(lx+7)+'" cy="'+(ly-5)+'" r="7" fill="'+it.c+'"/><text x="'+(lx+20)+'" y="'+ly+'" font-size="15" font-weight="600" fill="currentColor">'+praEsc(it.t)+'</text>'; lx+=w;});
  if(missing.length){ly+=22; leg+='<text x="14" y="'+ly+'" font-size="13" fill="#b45309">Not plotted (no map position for): '+praEsc(missing.join(', '))+'</text>';}
  var H=ly+12;

  host.innerHTML='<svg viewBox="0 0 '+W+' '+H.toFixed(0)+'" preserveAspectRatio="xMidYMid meet" xmlns="http://www.w3.org/2000/svg" role="img" '+
    'aria-label="Map of lead locations" style="display:block;width:100%;height:100%;color:inherit" font-family="Calibri,Segoe UI,Arial,sans-serif">'+
    '<defs><clipPath id="praGeoClip"><rect x="0" y="0" width="'+W+'" height="'+MH.toFixed(1)+'"/></clipPath></defs>'+
    '<g clip-path="url(#praGeoClip)"><g stroke="currentColor" stroke-opacity=".12" stroke-width="1">'+grid+'</g>'+
    '<g fill="currentColor" fill-opacity=".2" stroke="currentColor" stroke-opacity=".28" stroke-width="1" stroke-linejoin="round">'+land+'</g>'+
    '</g>'+
    '<g id="praGeoDots">'+dots+'</g><g id="praGeoTip"></g>'+leg+'</svg>';
  var svg=host.querySelector('svg');
  svg.addEventListener('click',function(e){var g=e.target.closest&&e.target.closest('.pra-geo-dot'); if(g)praFocusCountry(g.getAttribute('data-name'));});
}

function praFocusCountry(name){
  var N=PRA.geo_total||1;
  document.querySelectorAll('#s5 .geo-item').forEach(function(it){
    var a=it.querySelector('a'); var on=a&&a.dataset.mapQuery===name;
    it.style.background=on?'#fff4ea':''; it.style.borderLeftWidth=on?'6px':'';
  });
  var host=document.getElementById('geoMap'), svg=host&&host.querySelector('svg'); if(!svg)return;
  var hit=null;
  svg.querySelectorAll('.pra-geo-dot').forEach(function(g){var on=g.getAttribute('data-name')===name, c=g.querySelector('circle');
    c.setAttribute('stroke',on?'#0f172a':'#fff'); c.setAttribute('stroke-width',on?4:2); if(on)hit=g;});
  var tip=svg.querySelector('#praGeoTip'); tip.innerHTML=''; if(!hit)return;
  svg.querySelector('#praGeoDots').appendChild(hit);
  var c=hit.querySelector('circle'), cx=+c.getAttribute('cx'), cy=+c.getAttribute('cy'), r=+c.getAttribute('r'), n=+hit.getAttribute('data-count');
  var txt=name+' \u00b7 '+praFmtN(n)+' leads \u00b7 '+(n/N*100).toFixed(1)+'%', w=txt.length*7.6+22;
  var x=Math.min(Math.max(cx-w/2,4),996-w), y=cy-r-34; if(y<4)y=cy+r+8;
  tip.innerHTML='<rect x="'+x.toFixed(1)+'" y="'+y.toFixed(1)+'" width="'+w.toFixed(1)+'" height="26" rx="6" fill="#0f172a" fill-opacity=".92"/>'+
    '<text x="'+(x+w/2).toFixed(1)+'" y="'+(y+18).toFixed(1)+'" text-anchor="middle" font-size="15" font-weight="700" fill="#fff">'+praEsc(txt)+'</text>';
}

function praUpdateSlide5(){
  var s=document.getElementById('s5'); if(!s)return;
  var names=PRA.country_labels, vals=PRA.country_values, N=PRA.geo_total||1, unique=names.length;
  var mv=s.querySelector('.metric .value'); if(mv)mv.textContent=String(unique);
  var ml=s.querySelector('.metric .label'); if(ml){var t=ml.lastChild; if(t&&t.nodeType===3)t.nodeValue='Unique Countries';}
  var title=s.querySelector('.title');
  if(title)title.textContent='Location Split \u00b7 '+(unique<=3?names.join(' \u00b7 '):names.slice(0,2).join(' \u00b7 ')+' +'+(unique-2)+' more');
  var navItems=document.querySelectorAll('.slide-nav-item'); if(navItems[4]){var sm=navItems[4].querySelector('small'); if(sm)sm.textContent='Location Map';}
  var geoList=s.querySelector('.geo-list');
  if(geoList){geoList.innerHTML=PRA.geo_rows; geoList.style.maxHeight='230px'; geoList.style.overflowY='auto';
    geoList.addEventListener('click',function(e){var a=e.target.closest('a[data-map-query]'); if(!a)return; e.preventDefault(); praFocusCountry(a.dataset.mapQuery);});}
  var callout=s.querySelector('.callout');
  if(callout){
    var k=Math.min(5,unique), top=vals.slice(0,k).reduce(function(a,b){return a+b;},0);
    callout.textContent=unique?('Top '+k+' location'+(k>1?'s':'')+' ('+names.slice(0,k).join(', ')+') account for '+praFmtN(top)+' of '+praFmtN(N)+' leads ('+(top/N*100).toFixed(1)+'%).'):'No geographic data is available in the raw lead file.';
  }
  praRenderGeoMap();
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
        "device_is_pct": data.get("device_is_pct", False),
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
.pdfx .slide{padding-bottom:82px!important}
.pdfx #s1.slide .title{text-transform:uppercase;font-size:68px!important;letter-spacing:.01em;margin-bottom:14px!important}
@media(max-width:760px){.pdfx #s1.slide .title{font-size:52px!important}}
html[data-theme="dark"] .slide{background:linear-gradient(rgba(33,41,54,.93),rgba(27,35,47,.96)),repeating-linear-gradient(90deg,rgba(255,255,255,.04) 0 38px,transparent 38px 70px,rgba(255,255,255,.022) 70px 96px,transparent 96px 140px)!important;background-color:#222a36!important}
html[data-theme="dark"] .cover{background:#222833!important}
html[data-theme="dark"]{--trk:rgba(0,0,0,.32);--ink:#e6edf5}
html[data-theme="light"]{--trk:rgba(23,32,51,.09);--ink:#334155}
.pdf-grid{display:grid;gap:14px;margin-bottom:14px}
.pdf-panel h4,.pdf-kpi h4{margin:0 0 8px;text-align:center;font-size:14px;letter-spacing:.09em;text-transform:uppercase;color:var(--ink)}
.pdf-kpi{display:flex;flex-direction:column;align-items:center;justify-content:center;gap:10px;text-align:center;min-height:200px}
.pdf-kpi .num{font:900 54px/1 "Courier New",monospace;color:#12e7d4;text-shadow:0 0 12px rgba(18,231,212,.5)}
html[data-theme="light"] .pdf-kpi .num{color:#0e9aa7;text-shadow:none}
.pdf-cv{position:relative;width:100%}
.pdf-hbar{display:grid;grid-template-columns:96px 1fr 56px;align-items:center;gap:12px;margin:10px 0;font-size:13px;color:var(--ink);font-weight:700}
.pdf-hbar .lb,.pdf-hbar span:first-child{display:flex;align-items:center;justify-content:flex-end;gap:9px;text-align:right}
.pdf-hbar .vl,.pdf-hbar span:last-child{font:800 14px Arial,sans-serif}
.pdf-track{height:30px;background:var(--trk);border-radius:16px;overflow:hidden;box-shadow:inset 0 3px 7px rgba(0,0,0,.42),0 1px 0 rgba(255,255,255,.06)}
.pdf-fill{position:relative;height:100%;min-width:18px;border-radius:16px;box-shadow:inset 0 2px 0 rgba(255,255,255,.42),inset 0 -6px 9px rgba(0,0,0,.24)}
.pdf-fill:after{content:"";position:absolute;left:10px;right:14px;top:4px;height:9px;border-radius:9px;background:linear-gradient(rgba(255,255,255,.42),rgba(255,255,255,0))}
.ibadge{border-radius:50%;display:flex;align-items:center;justify-content:center;flex:none}
.ibadge svg{width:100%;height:100%;display:block;filter:drop-shadow(0 2px 2px rgba(0,0,0,.4))}
.pdf-kpi .ic{display:flex;justify-content:center;font-size:inherit}
.pdf-stat .ic{display:flex;justify-content:center;margin-bottom:8px;font-size:inherit}
.pdf-dev{display:flex;flex-direction:column;justify-content:center;gap:22px;height:100%;min-height:250px;padding:6px 10px 2px;color:var(--ink)}
.pdf-dev .tiles{display:flex;justify-content:space-around;align-items:flex-start;gap:14px}
.pdf-dev .tile{flex:1;min-width:0;display:flex;flex-direction:column;align-items:center;gap:7px;text-align:center}
.pdf-dev .badge{width:104px;height:104px;border-radius:50%;display:flex;align-items:center;justify-content:center;margin-bottom:4px}
.pdf-dev .badge svg{width:60px;height:60px;filter:drop-shadow(0 3px 3px rgba(0,0,0,.4))}
.pdf-dev .big{font:900 40px/1.05 "Courier New",monospace;letter-spacing:.01em}
html[data-theme="light"] .pdf-dev .big{filter:brightness(.72) saturate(1.25);text-shadow:none!important}
.pdf-dev .nm{font-size:16px;font-weight:800;letter-spacing:.03em}
.pdf-dev .ct{font-size:11px;font-weight:700;opacity:.65;letter-spacing:.05em;text-transform:uppercase}
.pdf-dev .bar{display:flex;height:30px;border-radius:16px;overflow:hidden;background:var(--trk);margin:0 6px;box-shadow:inset 0 3px 7px rgba(0,0,0,.4),0 10px 16px -10px rgba(0,0,0,.65)}
.pdf-dev .bar span{display:flex;align-items:center;justify-content:center;font:900 12px Arial,sans-serif;color:#fff;text-shadow:0 1px 3px rgba(0,0,0,.6);border-right:2px solid rgba(0,0,0,.28);box-shadow:inset 0 2px 0 rgba(255,255,255,.35)}
.pdf-dev .bar span:last-child{border-right:0}
@media(max-width:760px){.pdf-dev .badge{width:76px;height:76px}.pdf-dev .badge svg{width:44px;height:44px}.pdf-dev .big{font-size:28px}.pdf-dev .nm{font-size:13px}}
.pdf-stats{display:grid;grid-template-columns:repeat(6,1fr);gap:10px;margin-bottom:14px}
.pdf-stat{text-align:center;color:var(--ink)}.pdf-stat b{display:block;font-size:30px}.pdf-stat small{font-size:12px;font-weight:800;letter-spacing:.05em}
.pdf-obs{margin:0;padding:6px 10px;list-style:none;color:var(--ink);font-size:14px;line-height:1.75}
.pdf-obs li{padding-left:22px;position:relative;margin-bottom:6px}.pdf-obs li:before{content:"\27A2";position:absolute;left:0;color:#f47b20}
.pdf-thanks{flex:1;display:flex;align-items:center;justify-content:center;font-size:56px;font-weight:900;color:var(--ink)}
.pdf-offices{display:grid;grid-template-columns:repeat(3,1fr);gap:20px;color:var(--ink);font-size:12px;line-height:1.6;padding:0 40px 40px}
.pdf-offices b{display:block;font-size:13px}
#geoMap{height:300px!important}
.pdf-svg{color:var(--ink)}
.pdf-leg{display:flex;flex-wrap:wrap;gap:6px 16px;justify-content:center;margin-top:8px;font-size:11px;font-weight:700;color:var(--ink)}
.pdf-leg i{display:inline-block;width:12px;height:12px;margin-right:6px;vertical-align:-2px;border-radius:4px;box-shadow:inset 0 2px 0 rgba(255,255,255,.4),0 2px 4px rgba(0,0,0,.35)}
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
    function shade(c, f) {
      var m = /^#([0-9a-f]{6})$/i.exec(c || ''); if (!m) return c;
      var n = parseInt(m[1], 16), r = n >> 16, g = (n >> 8) & 255, b = n & 255;
      var t = function (v) { return Math.round(f < 0 ? v * (1 + f) : v + (255 - v) * f); };
      return 'rgb(' + t(r) + ',' + t(g) + ',' + t(b) + ')';
    }
    function rgba(c, a) {
      var m = /^#([0-9a-f]{6})$/i.exec(c || ''); if (!m) return c;
      var n = parseInt(m[1], 16); return 'rgba(' + (n >> 16) + ',' + ((n >> 8) & 255) + ',' + (n & 255) + ',' + a + ')';
    }
    var KC = { target: '#12e7d4', factory: '#ff9a1f', pin: '#ff5a7a', doc: '#7c5ce5', plane: '#3b82f6', inbox: '#10b981', mail: '#06b6d4', cursor: '#8b5cf6', check: '#f59e0b', ban: '#ef4444' };
    function ico(k) {
      var W = '#fff', D = 'rgba(8,16,30,.36)', b = '';
      if (k === 'target') b = '<circle cx="30" cy="34" r="22" fill="none" stroke="' + W + '" stroke-width="5"/><circle cx="30" cy="34" r="12" fill="none" stroke="' + W + '" stroke-width="5"/><circle cx="30" cy="34" r="4.5" fill="' + W + '"/><path d="M32 32L56 8" stroke="' + W + '" stroke-width="4.5" stroke-linecap="round"/><path d="M46 6v12h12" fill="none" stroke="' + W + '" stroke-width="3.5" stroke-linecap="round" stroke-linejoin="round"/>';
      else if (k === 'factory') b = '<path d="M5 57V29l15 9v-9l15 9V9h11v28h13v20z" fill="' + W + '"/><rect x="13" y="45" width="7" height="7" rx="1" fill="' + D + '"/><rect x="28" y="45" width="7" height="7" rx="1" fill="' + D + '"/><rect x="43" y="45" width="7" height="7" rx="1" fill="' + D + '"/>';
      else if (k === 'pin') b = '<path d="M32 3C19.5 3 11 12.5 11 24c0 15.5 21 37 21 37s21-21.5 21-37C53 12.5 44.5 3 32 3z" fill="' + W + '"/><circle cx="32" cy="24" r="8.5" fill="' + D + '"/>';
      else if (k === 'doc') b = '<path d="M13 3h27l15 15v43H13z" fill="' + W + '"/><path d="M40 3v15h15z" fill="' + D + '"/><rect x="21" y="29" width="26" height="4" rx="2" fill="' + D + '"/><rect x="21" y="38" width="26" height="4" rx="2" fill="' + D + '"/><rect x="21" y="47" width="17" height="4" rx="2" fill="' + D + '"/>';
      else if (k === 'plane') b = '<path d="M59 5L6 26l18 7 8 19z" fill="' + W + '"/><path d="M24 33L59 5 31 41z" fill="' + D + '"/>';
      else if (k === 'inbox') b = '<rect x="4" y="12" width="48" height="34" rx="5" fill="' + W + '"/><path d="M7 17l21 16L49 17" fill="none" stroke="' + D + '" stroke-width="3.5" stroke-linecap="round" stroke-linejoin="round"/><circle cx="48" cy="46" r="14" fill="#22c55e" stroke="' + W + '" stroke-width="3"/><path d="M41 46l5 5 9-10" fill="none" stroke="' + W + '" stroke-width="4" stroke-linecap="round" stroke-linejoin="round"/>';
      else if (k === 'mail') b = '<path d="M6 27L32 6l26 21z" fill="' + W + '" fill-opacity=".78"/><rect x="14" y="16" width="36" height="26" rx="2" fill="' + D + '"/><path d="M6 27l26 18 26-18v26a5 5 0 0 1-5 5H11a5 5 0 0 1-5-5z" fill="' + W + '"/>';
      else if (k === 'cursor') b = '<path d="M14 4l36 28-16 3 10 19-9 4-10-19-11 11z" fill="' + W + '" stroke="' + D + '" stroke-width="2" stroke-linejoin="round"/>';
      else if (k === 'check') b = '<circle cx="32" cy="32" r="27" fill="' + W + '"/><path d="M19 33l9.5 9.5L46 22" fill="none" stroke="' + D + '" stroke-width="6.5" stroke-linecap="round" stroke-linejoin="round"/>';
      else if (k === 'ban') b = '<circle cx="32" cy="32" r="25" fill="none" stroke="' + W + '" stroke-width="7"/><path d="M14.5 49.5l35-35" stroke="' + W + '" stroke-width="7" stroke-linecap="round"/>';
      return '<svg viewBox="0 0 64 64" xmlns="http://www.w3.org/2000/svg">' + b + '</svg>';
    }
    function badge(k, c, sz) {
      c = c || KC[k] || '#12e7d4'; var q = function (d) { return Math.round(sz / d); }, g = Math.round(sz * .58);
      return '<div class="ibadge" style="width:' + sz + 'px;height:' + sz + 'px;background:radial-gradient(circle at 30% 22%,' + shade(c, .55) + ',' + c + ' 52%,' + shade(c, -.4) + ');box-shadow:0 0 0 ' + q(16) + 'px ' + rgba(c, .14) + ',0 ' + q(5) + 'px ' + q(3.2) + 'px -' + q(6) + 'px ' + c + ',inset 0 -' + q(10) + 'px ' + q(5.5) + 'px rgba(0,0,0,.3),inset 0 ' + q(12) + 'px ' + q(7) + 'px rgba(255,255,255,.42)"><div style="width:' + g + 'px;height:' + g + 'px">' + ico(k) + '</div></div>';
    }
    function pill(cx, cy, t) {
      t = String(t); var w = t.length * 7.6 + 18;
      return '<rect x="' + (cx - w / 2) + '" y="' + (cy - 11) + '" width="' + w + '" height="21" rx="10.5" fill="rgba(148,163,184,.17)" stroke="rgba(148,163,184,.4)"/>' + T(cx, cy + 4, t, { s: 12, w: 800 });
    }
    function bar3v(x, y, w, h, d, gid, side, topc) {
      var b = y + h;
      return '<ellipse cx="' + (x + w / 2 + d / 2) + '" cy="' + (b + 2) + '" rx="' + (w * .64) + '" ry="4" fill="rgba(0,0,0,.28)"/>' +
        '<path d="M' + (x + w) + ',' + y + ' L' + (x + w + d) + ',' + (y - d) + ' L' + (x + w + d) + ',' + (b - d) + ' L' + (x + w) + ',' + b + 'Z" fill="' + side + '"/>' +
        '<path d="M' + x + ',' + y + ' L' + (x + d) + ',' + (y - d) + ' L' + (x + w + d) + ',' + (y - d) + ' L' + (x + w) + ',' + y + 'Z" fill="' + topc + '"/>' +
        '<rect x="' + x + '" y="' + y + '" width="' + w + '" height="' + h + '" fill="url(#' + gid + ')"/>' +
        '<rect x="' + (x + w * .1) + '" y="' + (y + 3) + '" width="' + (w * .17) + '" height="' + Math.max(h - 6, 0) + '" rx="' + (w * .085) + '" fill="rgba(255,255,255,.22)"/>';
    }
    function bar3h(x, y, w, h, d, gid, side, topc) {
      return '<path d="M' + x + ',' + y + ' L' + (x + d) + ',' + (y - d) + ' L' + (x + w + d) + ',' + (y - d) + ' L' + (x + w) + ',' + y + 'Z" fill="' + topc + '"/>' +
        '<path d="M' + (x + w) + ',' + y + ' L' + (x + w + d) + ',' + (y - d) + ' L' + (x + w + d) + ',' + (y + h - d) + ' L' + (x + w) + ',' + (y + h) + 'Z" fill="' + side + '"/>' +
        '<rect x="' + x + '" y="' + y + '" width="' + w + '" height="' + h + '" fill="url(#' + gid + ')"/>' +
        '<rect x="' + x + '" y="' + (y + 1.5) + '" width="' + w + '" height="' + (h * .38) + '" fill="rgba(255,255,255,.22)"/>';
    }
    function wrap(t, n) {
      t = String(t); if (t.length <= n) return [t];
      var a = '', b = ''; t.split(' ').forEach(function (x) { if (!b && (a + ' ' + x).trim().length <= n) a = (a + ' ' + x).trim(); else b = (b + ' ' + x).trim(); });
      if (b.length > n) b = b.slice(0, n - 1) + '\u2026';
      return [a || t.slice(0, n), b].filter(Boolean);
    }
    function lines(x, y, t, n, o) { var ls = wrap(t, n), s = ''; ls.forEach(function (l, k) { s += T(x, y + (ls.length > 1 ? (k ? 7 : -6) : 4), l, o); }); return s; }
    function legend(items) { return '<div class="pdf-leg">' + items.map(function (s) { return '<span><i style="background:' + s.color + '"></i>' + esc(s.name) + '</span>'; }).join('') + '</div>'; }

    function vbar(labels, vals, o) {
      o = o || {}; var W = o.w || 560, H = o.h || 210, n = labels.length || 1, pl = 14, pr = 24, pt = 40, pb = 44, id = 'g' + (++uid), ph = H - pt - pb;
      var max = Math.max.apply(null, vals.concat([1])), slot = (W - pl - pr) / n, bw = Math.min(o.bw || slot * .5, 76), d = Math.min(11, bw * .22), inner = '', defs = '';
      [0, .25, .5, .75, 1].forEach(function (g) { var y = H - pb - ph * g; inner += '<line x1="' + pl + '" x2="' + (W - pr) + '" y1="' + y + '" y2="' + y + '" stroke="rgba(148,163,184,' + (g ? .18 : .55) + ')"' + (g ? ' stroke-dasharray="3 5"' : '') + '/>'; });
      labels.forEach(function (l, i) {
        var h = Math.max(vals[i] / max * ph, 3), x = pl + slot * i + (slot - bw) / 2 - d / 2, y = H - pb - h, c = o.colors ? o.colors[i] : (o.color || '#4472c4'), gid = id + 'b' + i;
        var lo = o.grad ? o.grad[0] : shade(c, -.22), hi = o.grad ? o.grad[1] : shade(c, .32), side = shade(o.grad ? o.grad[0] : c, -.42), topc = shade(o.grad ? o.grad[1] : c, .5);
        defs += grad(gid, lo, hi, true);
        inner += bar3v(x, y, bw, h, d, gid, side, topc) + pill(x + bw / 2 + d / 2, y - d - 14, o.fmt ? o.fmt(vals[i], i) : vals[i]);
        wrap(l, slot < 120 ? 13 : 24).forEach(function (t, k) { inner += T(pl + slot * i + slot / 2, H - pb + 19 + k * 13, t, { s: 11 }); });
      });
      return svg(W, H, '<defs>' + defs + '</defs>' + inner);
    }
    function hbar(labels, vals, o) {
      o = o || {}; var n = labels.length || 1, rowH = o.rowH || 40, lw = o.lw || 150, W = o.w || 620, pr = o.pr || 56, aw = W - lw - pr, H = n * rowH + (o.axis ? 28 : 8);
      var max = o.max || Math.max.apply(null, vals.concat([1])), id = 'g' + (++uid), defs = '', inner = '';
      labels.forEach(function (l, i) {
        var yc = i * rowH + rowH / 2 + 3, bh = Math.min(o.bh || 26, rowH - 10), d = Math.min(7, bh * .26), w = Math.max(6, vals[i] / max * aw), y = yc - bh / 2 + d / 2, gid = id + 'h' + i;
        var c = o.colors ? o.colors[i] : (o.color || '#4472c4');
        var lo = o.grad ? o.grad[0] : shade(c, -.2), hi = o.grad ? o.grad[1] : shade(c, .34), side = shade(o.grad ? o.grad[0] : c, -.42), topc = shade(o.grad ? o.grad[1] : c, .5);
        defs += grad(gid, lo, hi);
        var txt = o.fmt ? o.fmt(vals[i], i) : vals[i];
        inner += lines(lw - 12, yc, l, o.wrap || 26, { a: 'end', s: 11 }) +
          '<rect x="' + lw + '" y="' + (y - d) + '" width="' + (aw + d) + '" height="' + (bh + d) + '" rx="4" fill="rgba(148,163,184,.12)"/>' +
          bar3h(lw, y, w, bh, d, gid, side, topc) +
          (o.inside && w > String(txt).length * 8 + 18 ? T(lw + w / 2, yc + 5, txt, { c: '#fff', s: 13, w: 800 }) : T(lw + w + d + 9, yc + 5, txt, { a: 'start', s: 13, w: 800 }));
      });
      if (o.axis) { inner += '<line x1="' + lw + '" x2="' + lw + '" y1="0" y2="' + (n * rowH) + '" stroke="rgba(160,174,192,.6)"/>'; [0, 20, 40, 60, 80, 100].forEach(function (t) { inner += T(lw + aw * t / 100, H - 6, t + '%', { s: 10, w: 600 }); }); }
      return svg(W, H, '<defs>' + defs + '</defs>' + inner);
    }
    function clustered(labels, series, o) {
      o = o || {}; var W = o.w || 560, H = o.h || 230, n = labels.length || 1, k = series.length || 1, pl = 12, pt = 36, pb = 44, slot = (W - 2 * pl) / n, gw = slot * .84, bw = Math.min(gw / k - 3, 34), d = Math.min(6, bw * .22), inner = '', defs = '', id = 'g' + (++uid), ph = H - pt - pb;
      var max = Math.max.apply(null, [1].concat.apply([], series.map(function (s) { return s.data || []; })));
      [0, .25, .5, .75, 1].forEach(function (g) { var y = H - pb - ph * g; inner += '<line x1="' + pl + '" x2="' + (W - pl) + '" y1="' + y + '" y2="' + y + '" stroke="rgba(148,163,184,' + (g ? .18 : .55) + ')"' + (g ? ' stroke-dasharray="3 5"' : '') + '/>'; });
      series.forEach(function (s, j) { defs += grad(id + 's' + j, shade(s.color, -.22), shade(s.color, .32), true); });
      labels.forEach(function (l, i) {
        var x0 = pl + slot * i + (slot - (bw + 3) * k) / 2;
        series.forEach(function (s, j) {
          var v = (s.data || [])[i] || 0, h = v ? Math.max(v / max * ph, 3) : 0, x = x0 + j * (bw + 3);
          if (v) inner += bar3v(x, H - pb - h, bw, h, d, id + 's' + j, shade(s.color, -.42), shade(s.color, .5)) + T(x + bw / 2 + d / 2, H - pb - h - d - 6, v, { s: 11, w: 800 });
        });
        wrap(l, slot < 110 ? 12 : 22).forEach(function (t, q) { inner += T(pl + slot * i + slot / 2, H - pb + 19 + q * 13, t, { s: 11 }); });
      });
      return svg(W, H, '<defs>' + defs + '</defs>' + inner) + legend(series);
    }
    function stacked(rows, series, o) {
      o = o || {}; var W = o.w || 680, rowH = 48, lw = 210, pr = 24, aw = W - lw - pr, H = rows.length * rowH + 8, inner = '', defs = '', id = 'g' + (++uid);
      var tots = rows.map(function (_, i) { return sum(series.map(function (s) { return s.data[i] || 0; })); }), max = Math.max.apply(null, tots.concat([1]));
      series.forEach(function (s, j) { defs += grad(id + 's' + j, shade(s.color, -.28), shade(s.color, .3), true); });
      rows.forEach(function (r, i) {
        var yc = i * rowH + rowH / 2 + 3, y0 = yc - 15, x = lw, tw = tots[i] / max * aw, cp = id + 'c' + i, seg = '', txt = '';
        inner += lines(lw - 12, yc, r, 34, { a: 'end', s: 11 }) + '<rect x="' + lw + '" y="' + (y0 - 1) + '" width="' + aw + '" height="32" rx="9" fill="rgba(148,163,184,.12)"/>';
        defs += '<clipPath id="' + cp + '"><rect x="' + lw + '" y="' + y0 + '" width="' + Math.max(tw, 1) + '" height="30" rx="8"/></clipPath>';
        series.forEach(function (s, j) {
          var v = s.data[i] || 0; if (!v) return; var w = v / max * aw;
          seg += '<rect x="' + x + '" y="' + y0 + '" width="' + w + '" height="30" fill="url(#' + id + 's' + j + ')"/><rect x="' + (x + w - 1.5) + '" y="' + y0 + '" width="1.5" height="30" fill="rgba(0,0,0,.28)"/>';
          if (w > 16) txt += T(x + w / 2, yc + 5, v, { c: '#fff', s: 12, w: 800 });
          x += w;
        });
        inner += '<g clip-path="url(#' + cp + ')">' + seg + '<rect x="' + lw + '" y="' + y0 + '" width="' + Math.max(tw, 1) + '" height="13" fill="rgba(255,255,255,.2)"/></g>' + txt;
      });
      return svg(W, H, '<defs>' + defs + '</defs>' + inner) + legend(series);
    }
    function donut(p, color, text, cap) {
      var r = 64, sw = 30, cx = 130, cy = 104, D = 12, c = 2 * Math.PI * r, d = Math.max(0, Math.min(100, p)) / 100 * c, id = 'd' + (++uid), depth = '';
      var ring = function (y, col, dash) { return '<circle cx="' + cx + '" cy="' + (cy + y) + '" r="' + r + '" fill="none" stroke="' + col + '" stroke-width="' + sw + '"' + (dash ? ' stroke-dasharray="' + d + ' ' + (c - d) + '" transform="rotate(-90 ' + cx + ' ' + (cy + y) + ')"' : '') + '/>'; };
      for (var k = D; k >= 1; k--) depth += ring(k, '#2b3647', false) + (d > 0 ? ring(k, shade(color, -.5), true) : '');
      return svg(260, 250,
        '<defs><linearGradient id="' + id + 'v" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="' + shade(color, .45) + '"/><stop offset="1" stop-color="' + color + '"/></linearGradient>' +
        '<linearGradient id="' + id + 't" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="#7b889b"/><stop offset="1" stop-color="#566377"/></linearGradient>' +
        '<radialGradient id="' + id + 's" cx=".5" cy=".5" r=".5"><stop offset="0" stop-color="#000" stop-opacity=".5"/><stop offset="1" stop-color="#000" stop-opacity="0"/></radialGradient></defs>' +
        '<ellipse cx="' + cx + '" cy="' + (cy + r + sw / 2 + D + 6) + '" rx="' + (r + 34) + '" ry="12" fill="url(#' + id + 's)"/>' + depth +
        ring(0, 'url(#' + id + 't)', false) + (d > 0 ? ring(0, 'url(#' + id + 'v)', true) : '') +
        '<circle cx="' + cx + '" cy="' + cy + '" r="' + (r + sw / 2 - 2) + '" fill="none" stroke="rgba(255,255,255,.3)" stroke-width="1.2"/>' +
        '<circle cx="' + cx + '" cy="' + cy + '" r="' + (r - sw / 2 + 2) + '" fill="none" stroke="rgba(0,0,0,.25)" stroke-width="1.2"/>' +
        T(cx, cy + 6, text, { s: 24, w: 800 }) + (cap ? T(cx, cy + 24, cap, { s: 9, w: 700 }) : ''));
    }
    function pie(labels, vals, colors) {
      /* Real 3D pie: lit top face, extruded side wall, floor shadow, gloss, and leader-line labels that are
         spread apart so neighbouring small slices (e.g. Netherlands / Canada) can never overlap. */
      var W = 800, H = 372, cx = 400, cy = 170, r = 124, ry = .54, R = r * ry, dp = 30, tot = sum(vals) || 1, id = 'p' + (++uid);
      var f1 = function (n) { return +n.toFixed(2); };
      var P2 = function (a, dy) { return [cx + r * Math.cos(a), cy + R * Math.sin(a) + (dy || 0)]; };
      var pt = function (p) { return f1(p[0]) + ',' + f1(p[1]); };
      var defs = '<radialGradient id="' + id + 's" cx=".5" cy=".5" r=".5"><stop offset="0" stop-color="#000" stop-opacity=".6"/><stop offset=".65" stop-color="#000" stop-opacity=".22"/><stop offset="1" stop-color="#000" stop-opacity="0"/></radialGradient>' +
        '<linearGradient id="' + id + 'w" gradientUnits="userSpaceOnUse" x1="' + (cx - r) + '" y1="0" x2="' + (cx + r) + '" y2="0"><stop offset="0" stop-color="#000" stop-opacity=".55"/><stop offset=".28" stop-color="#fff" stop-opacity=".14"/><stop offset=".55" stop-color="#000" stop-opacity="0"/><stop offset="1" stop-color="#000" stop-opacity=".6"/></linearGradient>' +
        '<radialGradient id="' + id + 'g" cx=".3" cy=".3" r=".75"><stop offset="0" stop-color="#fff" stop-opacity=".2"/><stop offset=".5" stop-color="#fff" stop-opacity=".04"/><stop offset="1" stop-color="#fff" stop-opacity="0"/></radialGradient>';
      var a0 = -Math.PI / 2, walls = '', tops = '', items = [];
      vals.forEach(function (v, i) {
        var col = colors[i % colors.length], frac = v / tot, a1 = a0 + frac * 2 * Math.PI;
        if (v > 0) {
          defs += '<linearGradient id="' + id + 't' + i + '" gradientUnits="userSpaceOnUse" x1="0" y1="' + (cy - R) + '" x2="0" y2="' + (cy + R) + '"><stop offset="0" stop-color="' + shade(col, .38) + '"/><stop offset="1" stop-color="' + shade(col, -.14) + '"/></linearGradient>';
          var p0 = P2(a0), p1 = P2(a1), big = (a1 - a0) > Math.PI ? 1 : 0, top;
          if (frac >= .9999) top = 'M' + (cx - r) + ',' + cy + ' A' + r + ',' + f1(R) + ' 0 1 1 ' + (cx + r) + ',' + cy + ' A' + r + ',' + f1(R) + ' 0 1 1 ' + (cx - r) + ',' + cy + 'Z';
          else top = 'M' + cx + ',' + cy + ' L' + pt(p0) + ' A' + r + ',' + f1(R) + ' 0 ' + big + ' 1 ' + pt(p1) + 'Z';
          var s = Math.max(a0, 0), e = Math.min(a1, Math.PI);
          if (e > s) {
            var wd = 'M' + pt(P2(s)) + ' A' + r + ',' + f1(R) + ' 0 0 1 ' + pt(P2(e)) + ' L' + pt(P2(e, dp)) + ' A' + r + ',' + f1(R) + ' 0 0 0 ' + pt(P2(s, dp)) + 'Z';
            walls += '<path d="' + wd + '" fill="' + shade(col, -.45) + '"/><path d="' + wd + '" fill="url(#' + id + 'w)"/>';
          }
          tops += '<path d="' + top + '" fill="url(#' + id + 't' + i + ')" stroke="rgba(255,255,255,.8)" stroke-width="1.6" stroke-linejoin="round"/>';
          items.push({ i: i, m: (a0 + a1) / 2, col: col, v: v });
        }
        a0 = a1;
      });
      var sides = { r: [], l: [] };
      items.forEach(function (it) {
        var c = Math.cos(it.m), sn = Math.sin(it.m);
        it.ax = cx + r * c; it.ay = cy + R * sn + (sn > 0 ? dp * .55 : 0);
        it.ex = cx + (r + 22) * c; it.ey = it.ay + sn * 16; it.ly = it.ey; it.dir = c >= 0 ? 1 : -1;
        sides[c >= 0 ? 'r' : 'l'].push(it);
      });
      var gap = 42, lo = 28, hi = H - 28;
      ['r', 'l'].forEach(function (k) {
        var a = sides[k]; a.sort(function (p, q) { return p.ly - q.ly; });
        /* relax: push neighbours apart symmetrically so label stacks stay centred on their slices */
        for (var it2 = 0; it2 < 60; it2++) {
          var moved = false;
          for (var n = 1; n < a.length; n++) {
            var d = a[n].ly - a[n - 1].ly;
            if (d < gap - .5) { var h = (gap - d) / 2; a[n - 1].ly -= h; a[n].ly += h; moved = true; }
          }
          a.forEach(function (x) { x.ly = Math.min(hi, Math.max(lo, x.ly)); });
          if (!moved) break;
        }
      });
      var lab = '';
      items.forEach(function (it) {
        var lx = cx + it.dir * (r + 96), an = it.dir > 0 ? 'start' : 'end', tx = lx + it.dir * 13, nm = String(labels[it.i]);
        if (nm.length > 20) nm = nm.slice(0, 19) + '\u2026';
        lab += '<polyline points="' + f1(it.ax) + ',' + f1(it.ay) + ' ' + f1(it.ex) + ',' + f1(it.ey) + ' ' + f1(lx - it.dir * 8) + ',' + f1(it.ly + 3) + '" fill="none" stroke="' + it.col + '" stroke-width="1.7" stroke-linejoin="round"/>' +
          '<circle cx="' + f1(it.ax) + '" cy="' + f1(it.ay) + '" r="3.6" fill="#fff" stroke="' + it.col + '" stroke-width="2"/>' +
          '<circle cx="' + f1(lx) + '" cy="' + f1(it.ly + 3) + '" r="5.5" fill="' + it.col + '" stroke="rgba(255,255,255,.85)" stroke-width="1.4"/>' +
          T(tx, it.ly - 1, nm, { a: an, s: 14, w: 800 }) +
          '<text x="' + f1(tx) + '" y="' + f1(it.ly + 16) + '" text-anchor="' + an + '" font-size="13" fill="currentColor"><tspan font-weight="800" fill="' + it.col + '">' + pc(it.v, tot, 1) + '%</tspan><tspan font-weight="600" fill-opacity=".7"> \u00b7 ' + Number(it.v).toLocaleString() + '</tspan></text>';
      });
      return svg(W, H,
        '<defs>' + defs + '</defs>' +
        '<ellipse cx="' + cx + '" cy="' + (cy + dp + 8) + '" rx="' + f1(r * 1.1) + '" ry="' + f1(R * 1.3) + '" fill="url(#' + id + 's)"/>' +
        walls + tops +
        '<ellipse cx="' + cx + '" cy="' + cy + '" rx="' + r + '" ry="' + f1(R) + '" fill="url(#' + id + 'g)"/>' +
        '<ellipse cx="' + cx + '" cy="' + cy + '" rx="' + r + '" ry="' + f1(R) + '" fill="none" stroke="rgba(255,255,255,.35)" stroke-width="1"/>' + lab);
    }

    /* ---------- page helpers ---------- */
    var panel = function (h, body, st) { return '<div class="panel pdf-panel"' + (st ? ' style="' + st + '"' : '') + '>' + (h ? '<h4>' + h + '</h4>' : '') + body + '</div>'; };
    var kpi = function (t, ic, v) { return '<div class="panel pdf-kpi"><h4>' + t + '</h4><div class="ic">' + badge(ic, KC[ic], 88) + '</div><div class="num">' + v + '</div></div>'; };
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
      grid('1fr 2fr 2fr', kpi('Leads Generated', 'target', leads) + panel('Job Level', vbar(jl, jv, { w: 520, h: 200, bw: 46, grad: ['#1d4ed8', '#38bdf8'] })) + panel('Job Level Split', hrows(split, [['#a21caf', '#f0f'], ['#0284c7', '#22e5ff'], ['#f59e0b', '#fde047'], ['#ea580c', '#fbbf24']]))) +
      grid('2fr 1.3fr 1.3fr', panel('Job Functions', hbar(P.func_labels, P.func_values.map(function (v) { return v / ft * 100; }), { w: 520, rowH: 70, bh: 56, lw: 130, pr: 20, max: 100, axis: true, inside: true, wrap: 16, color: '#4472c4', fmt: function (v) { return v.toFixed(2) + '%'; } })) + panel('Decision Makers', donut(dmP, '#12e7d4', dmP.toFixed(1) + '%', 'DECISION MAKERS')) + panel('Recommender', donut(100 - dmP, '#ff9a1f', (100 - dmP).toFixed(1) + '%', 'INFLUENCERS'))));

    /* ---- Slide 4 ---- */
    var ind = P.industry_labels.slice(0, 5), indv = P.industry_values.slice(0, 5);
    var si = order(P.size_labels, function (a, b) { return sizeKey(b) - sizeKey(a); });
    var sl = si.map(function (i) { return P.size_labels[i]; }), sv = si.map(function (i) { return P.size_values[i]; }), stt = sum(sv) || 1;
    var dt = sum(P.device_values) || 1;
    var DEV_COL = ['#12e7d4', '#ff9a1f', '#7c5ce5', '#ff3d8b'];
    function devKind(l, i) {
      l = String(l);
      if (/desk|laptop|\bpc\b|computer|windows|\bmac\b/i.test(l)) return 'desktop';
      if (/mobile|phone|android|\bios\b|iphone/i.test(l)) return 'phone';
      if (/tablet|ipad/i.test(l)) return 'tablet';
      if (/influ|recommend|non[\s-]*decision|research|evaluat/i.test(l)) return 'group';
      if (/decision|maker|executive|c[\s-]*level|budget/i.test(l)) return 'decision';
      return ['decision', 'group', 'desktop', 'phone'][i % 4];
    }
    function devIcon(kind) {
      var W = '#fff', D = 'rgba(8,16,30,.34)', b = '';
      if (kind === 'desktop') b = '<rect x="7" y="9" width="50" height="33" rx="4.5" fill="' + W + '"/><rect x="11.5" y="13.5" width="41" height="24" rx="2" fill="' + D + '"/><path d="M14 34L23 26L29 31L38 21L45 28L50 27V35H14Z" fill="#fff" fill-opacity=".34"/><path d="M24 43h16l2.4 8H21.6z" fill="' + W + '" fill-opacity=".92"/><rect x="16" y="50" width="32" height="4.5" rx="2.25" fill="' + W + '"/>';
      else if (kind === 'phone') b = '<rect x="18" y="4" width="28" height="56" rx="6.5" fill="' + W + '"/><rect x="22.5" y="12" width="19" height="35" rx="2" fill="' + D + '"/><path d="M24 42L29 35L33 39L38 31V44H24Z" fill="#fff" fill-opacity=".34"/><rect x="28" y="7" width="8" height="2" rx="1" fill="' + D + '"/><circle cx="32" cy="53.5" r="2.4" fill="' + D + '"/>';
      else if (kind === 'tablet') b = '<rect x="8" y="7" width="48" height="50" rx="6.5" fill="' + W + '"/><rect x="13" y="12" width="38" height="36" rx="2" fill="' + D + '"/><path d="M15 43L24 33L30 39L38 27L49 41V46H15Z" fill="#fff" fill-opacity=".34"/><circle cx="32" cy="52.5" r="2.2" fill="' + D + '"/>';
      else if (kind === 'group') b = '<circle cx="13" cy="25" r="6.5" fill="' + W + '" fill-opacity=".72"/><path d="M1.5 50c0-8.5 5-13.5 11.5-13.5 2.3 0 4.3.6 6 1.6C16.6 42 15.5 45.8 15.5 50z" fill="' + W + '" fill-opacity=".72"/><circle cx="51" cy="25" r="6.5" fill="' + W + '" fill-opacity=".72"/><path d="M62.5 50c0-8.5-5-13.5-11.5-13.5-2.3 0-4.3.6-6 1.6C47.4 42 48.5 45.8 48.5 50z" fill="' + W + '" fill-opacity=".72"/><circle cx="32" cy="20" r="9.5" fill="' + W + '"/><path d="M14.5 56c0-11.5 7-19 17.5-19s17.5 7.5 17.5 19z" fill="' + W + '"/>';
      else if (kind === 'decision') b = '<circle cx="30" cy="19" r="10.5" fill="' + W + '"/><path d="M9 57c0-12.5 8.5-20.5 21-20.5S51 44.5 51 57z" fill="' + W + '"/><path d="M30 38l-4.2 5 3 3-2.4 9.5h7.2L31 46l3-3z" fill="' + D + '"/><path d="M52 4.5l2.7 5.5 6 .8-4.4 4.2 1.1 6-5.4-2.9-5.4 2.9 1.1-6-4.4-4.2 6-.8z" fill="#ffd54a"/>';
      else b = '<circle cx="32" cy="32" r="24" fill="' + W + '"/><path d="M32 32V8a24 24 0 0 1 22.8 16.6z" fill="' + D + '"/>';
      return '<svg viewBox="0 0 64 64" xmlns="http://www.w3.org/2000/svg">' + b + '</svg>';
    }
    var devHtml;
    if (!P.device_values || !P.device_values.length || !sum(P.device_values)) devHtml = '<div class="pdf-dev"><div class="tiles"><div class="tile"><div class="nm">No data</div></div></div></div>';
    else {
      var shown = P.device_labels.slice(0, 4), tiles = '', bar = '', used = 0;
      shown.forEach(function (l, i) {
        var c = DEV_COL[i % DEV_COL.length], v = P.device_values[i] || 0, p = pc(v, dt, 1); used += v;
        tiles += '<div class="tile"><div class="badge" style="background:radial-gradient(circle at 30% 22%,' + shade(c, .55) + ',' + c + ' 52%,' + shade(c, -.4) + ');box-shadow:0 0 0 6px ' + rgba(c, .14) + ',0 18px 26px -12px ' + c + ',inset 0 -10px 18px rgba(0,0,0,.32),inset 0 8px 14px rgba(255,255,255,.42)">' + devIcon(devKind(l, i)) + '</div>' +
          '<div class="big" style="color:' + c + ';text-shadow:0 0 16px ' + rgba(c, .55) + '">' + p + '%</div><div class="nm">' + esc(l) + '</div><div class="ct">' + (P.device_is_pct ? 'of audience' : Number(v).toLocaleString() + ' leads') + '</div></div>';
        bar += '<span style="width:' + p + '%;background:linear-gradient(180deg,' + shade(c, .42) + ' 0%,' + c + ' 48%,' + shade(c, -.32) + ' 100%)">' + (+p >= 11 ? p + '%' : '') + '</span>';
      });
      var restV = dt - used;
      if (restV > 0.0001) bar += '<span style="width:' + pc(restV, dt, 1) + '%;background:linear-gradient(180deg,#cbd5e1,#94a3b8 50%,#64748b)"></span>';
      devHtml = '<div class="pdf-dev"><div class="tiles">' + tiles + '</div><div class="bar">' + bar + '</div></div>';
    }
    slide('s4', 'Campaign Dashboard',
      grid('1fr 3fr', kpi('Unique Industries', 'factory', P.unique_industry_count || P.industry_labels.length) + panel('Top 5 Industries', vbar(ind.map(function (l) { return l.toUpperCase(); }), indv, { w: 760, h: 210, bw: 34, color: '#4472c4', fmt: function (v) { return pc(v, leads) + '%'; } }))) +
      grid('1.4fr 1fr', panel('Employee Size', hbar(sl, sv, { w: 560, rowH: 34, bh: 24, lw: 80, pr: 50, grad: ['#7c3aed', '#ff3d8b'], fmt: function (v) { return pc(v, stt) + '%'; } })) + panel(P.device_title || 'Devices', devHtml)));

    /* ---- Slide 5 (pie on top, SVG location map below) ---- */
    var s5 = $('s5');
    if (s5) {
      var t5 = s5.querySelector('.title'); if (t5) t5.textContent = 'Campaign Dashboard';
      var cl = P.country_labels.slice(0, 7), cvv = P.country_values.slice(0, 7), rest = (P.geo_total || sum(P.country_values)) - sum(cvv);
      if (rest > 0) { cl.push('Others'); cvv.push(rest); }
      var row = document.createElement('div');
      row.innerHTML = grid('1fr 3fr', kpi('Unique Geo Locations', 'pin', P.unique_country_count || P.country_labels.length) + panel('Location Split', pie(cl, cvv, (window.PRA_PIE_COLORS || ['#70ad47', '#ffc000', '#ed7d31', '#4472c4', '#5b9bd5', '#a5a5a5', '#9e480e', '#7c5ce5']))));
      if (t5) t5.after(row.firstChild);
    }

    /* ---- Slide 6 ---- */
    var ai = order(P.asset_labels, function (a, b, i, j) { return P.asset_values[i] - P.asset_values[j]; }), at = sum(P.asset_values) || 1;
    var seriesBy = function (mat) { return P.asset_names.map(function (a, i) { return { name: a, color: COL[i % COL.length], data: (mat && mat[i]) || [] }; }); };
    slide('s6', 'Asset Dashboard',
      grid('1fr 3fr', kpi('Unique Assets', 'doc', P.unique_asset_count) + panel('Asset Split', hbar(ai.map(function (i) { return P.asset_labels[i]; }).reverse(), ai.map(function (i) { return P.asset_values[i] / at * 100; }).reverse(), { w: 700, rowH: 44, lw: 260, wrap: 38, bh: 24, pr: 56, max: 100, colors: ai.map(function (i) { return acol(P.asset_labels[i]); }).reverse(), fmt: function (v) { return v.toFixed(1) + '%'; } }))) +
      grid('1fr 1fr', panel('Asset Engagement by Top Industries', clustered(P.industry_names.slice(0, 5), seriesBy(P.asset_industry_matrix))) + panel('Asset Engagement by Country', clustered(P.country_names || [], seriesBy(P.asset_country_matrix)))));

    /* ---- Slide 7 ---- */
    var szI = order(P.size_names, function (a, b) { return sizeKey(a) - sizeKey(b); });
    var sizeSeries = szI.map(function (j, k) { return { name: P.size_names[j], color: COL[k % COL.length], data: P.asset_names.map(function (_, a) { return (P.asset_size_matrix && P.asset_size_matrix[a]) ? P.asset_size_matrix[a][j] : 0; }) }; });
    var jobSeries = P.job_names.map(function (n, j) { return { name: n, color: COL[j % COL.length], data: P.asset_names.map(function (_, a) { return (P.asset_job_matrix && P.asset_job_matrix[a]) ? P.asset_job_matrix[a][j] : 0; }) }; });
    slide('s7', 'Asset Dashboard', panel('Asset Engagement by Employee Size', stacked(P.asset_names, sizeSeries)) + '<div style="height:14px"></div>' + panel('Asset Engagement by Job Level', stacked(P.asset_names, jobSeries)));

    /* ---- Slides 8-10 ---- */
    function splitSlide(id, title, word, labels, values) {
      var t = sum(values) || 1;
      slide(id, title, grid('1fr 3fr', kpi('Unique Assets', 'doc', P.unique_asset_count) + panel(word + ' Split', hbar(labels, values, { w: 700, rowH: 40, lw: 270, wrap: 40, bh: 22, pr: 50, outline: true }))) +
        panel(word + ' Percentage', vbar(labels, values, { w: 900, h: 230, bw: 120, colors: labels.map(acol), fmt: function (v) { return pc(v, t) + '%'; } })));
    }
    splitSlide('s8', 'Asset Wise Open Split', 'Open', P.asset_open_labels, P.asset_open_values);
    splitSlide('s9', 'Asset Wise Click Split', 'Click', P.asset_click_labels, P.asset_click_values);
    splitSlide('s10', 'Asset Wise Conversion Split', 'Conversion', P.asset_conv_labels, P.asset_conv_values);

    /* ---- Slide 11 ---- */
    var st = [['plane', P.sent, 'SENT'], ['inbox', P.delivered, 'DELIVERED'], ['mail', P.opens, 'OPENS'], ['cursor', P.clicks, 'CLICKS'], ['check', P.conversion, 'CONVERSION'], ['ban', P.bounced, 'BOUNCED']];
    var rates = [['Bounce', P.bounce_rate], ['Conversion', P.conversion_rate], ['Clicks', P.click_rate], ['Open', P.open_rate], ['Delivered', P.delivery_rate]];
    slide('s11', 'Campaign Statistics',
      '<div class="panel"><div class="pdf-stats">' + st.map(function (x) { return '<div class="pdf-stat"><div class="ic">' + badge(x[0], KC[x[0]], 64) + '</div><b>' + Number(x[1] || 0).toLocaleString() + '</b><small>' + x[2] + '</small></div>'; }).join('') + '</div></div><div style="height:14px"></div>' +
      panel('Statistics Split', rates.map(function (r) {
        var RC = { Bounce: ['ban', '#ef4444', '#fb923c'], Conversion: ['check', '#f59e0b', '#fde047'], Clicks: ['cursor', '#8b5cf6', '#c4b5fd'], Open: ['mail', '#06b6d4', '#67e8f9'], Delivered: ['inbox', '#10b981', '#6ee7b7'] }, k = RC[r[0]] || ['check', '#3b82f6', '#93c5fd'];
        return '<div class="pdf-hbar" style="grid-template-columns:170px 1fr 66px"><span class="lb">' + badge(k[0], k[1], 32) + r[0] + '</span><div class="pdf-track"><div class="pdf-fill" style="width:' + Math.max(3, Math.min(100, r[1])) + '%;background:linear-gradient(90deg,' + k[1] + ',' + k[2] + ');box-shadow:inset 0 2px 0 rgba(255,255,255,.42),inset 0 -6px 9px rgba(0,0,0,.24),0 0 14px ' + rgba(k[1], .55) + '"></div></div><span class="vl">' + Number(r[1]).toFixed(1) + '%</span></div>';
      }).join('')));

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