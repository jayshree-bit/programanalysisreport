from __future__ import annotations

import pandas as pd
from typing import Any

from pra.helpers import find_column, series_clean, top_counts


try:
    from countryinfo import CountryInfo
except ImportError:
    CountryInfo = None

try:
    import plotly  # noqa: F401
except ImportError:
    plotly = None

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