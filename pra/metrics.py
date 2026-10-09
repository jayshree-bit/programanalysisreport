from __future__ import annotations

import math
import pandas as pd

from pra.helpers import numeric_column, series_clean, top_counts


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