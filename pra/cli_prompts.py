from __future__ import annotations

import pandas as pd
from typing import Any


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