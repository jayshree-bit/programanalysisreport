from __future__ import annotations

import pandas as pd
from typing import Any

from pra.config import DATE_FORMATS, HEX_COLOR
from pra.geo import location_points
from pra.helpers import business_days, find_column, ordered_job_levels, pct, safe_json, series_clean, top_counts
from pra.metrics import asset_event_metrics


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