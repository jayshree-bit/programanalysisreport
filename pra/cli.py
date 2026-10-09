from __future__ import annotations

import argparse
import os
import pandas as pd
from pathlib import Path

from pra.cli_prompts import collect_inputs
from pra.config import DATE_FORMATS, DEFAULT_EXCEL, DEFAULT_TEMPLATE
from pra.layers.design_layer import DESIGNS
from pra.report import build_report


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
    parser.add_argument("--design", choices=sorted(DESIGNS), default="classic",
                        help="Report design (can also be changed later inside the report)")
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
    inputs.setdefault("design", args.design)
    inputs.setdefault("palette", [c.strip() for c in args.palette.split(",") if c.strip()])

    build_report(template_path, excel_path, inputs)