"""
Web front-end for the Program Analysis Report generator.

Run:
    python app.py
Then open http://127.0.0.1:5000 in the browser, fill in the campaign inputs,
upload the raw leads Excel (.xlsx / .xls / .ods / .csv) and click Generate.

The report itself is produced by New_generate_dynamic_pra.build_report(), so
the output is identical to the command-line version.
"""

from __future__ import annotations

import base64
import tempfile
from pathlib import Path

import pandas as pd
from flask import Flask, abort, jsonify, request, send_from_directory

from New_generate_dynamic_pra import BASE_DIR, DATE_FORMATS, DEFAULT_TEMPLATE, DESIGNS, HEX_COLOR, build_report

REPORTS_DIR = BASE_DIR / "generated_reports"
ALLOWED_EXT = {".xlsx", ".xls", ".ods", ".csv"}
INT_FIELDS = ["sent", "delivered", "opens", "clicks", "conversion"]
LOGO_TYPES = {".png": "image/png", ".jpg": "image/jpeg", ".jpeg": "image/jpeg",
              ".svg": "image/svg+xml", ".webp": "image/webp", ".gif": "image/gif"}
MAX_LOGO_BYTES = 2 * 1024 * 1024

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 25 * 1024 * 1024  # 25 MB upload limit


def parse_date(value: str, label: str) -> str:
    dt = pd.to_datetime(value, errors="coerce")
    if not value or pd.isna(dt):
        raise ValueError(f"{label} is not a valid date.")
    return dt.strftime("%d-%b-%Y")


@app.get("/")
def index():
    return send_from_directory(BASE_DIR, "pra_input_form.html")


@app.post("/api/generate")
def generate():
    form = request.form
    try:
        campaign = form.get("campaign_name", "").strip()
        if not campaign:
            raise ValueError("Campaign Name is required.")

        inputs = {
            "campaign_name": campaign,
            "report_date": parse_date(form.get("report_date", ""), "Report Date"),
            "start_date": parse_date(form.get("start_date", ""), "Start Date"),
            "end_date": parse_date(form.get("end_date", ""), "End Date"),
            "prepared_by": form.get("prepared_by", "").strip() or "Quality Department",
            "maps_key": form.get("maps_key", "").strip(),
            "date_format": form.get("date_format", "month_year"),
            "palette": [c.strip() for c in form.get("palette", "").split(",") if c.strip()],
        }
        if inputs["date_format"] not in DATE_FORMATS:
            raise ValueError("Unknown date format.")
        inputs["design"] = form.get("design", "classic")
        if inputs["design"] not in DESIGNS:
            raise ValueError("Unknown report design.")
        if any(not HEX_COLOR.match(c) for c in inputs["palette"]):
            raise ValueError("Chart colours must be hex values like #3f5bd8.")

        logo = request.files.get("logo")
        if logo and logo.filename:
            logo_type = LOGO_TYPES.get(Path(logo.filename).suffix.lower())
            if not logo_type:
                raise ValueError("Logo must be a PNG, JPG, SVG, WEBP or GIF image.")
            raw = logo.read()
            if len(raw) > MAX_LOGO_BYTES:
                raise ValueError("Logo image must be 2 MB or smaller.")
            inputs["logo_data_uri"] = f"data:{logo_type};base64,{base64.b64encode(raw).decode()}"
        for name in INT_FIELDS:
            raw = form.get(name, "").replace(",", "").strip()
            if not raw.isdigit():
                raise ValueError(f"{name.title()} must be a whole number >= 0.")
            inputs[name] = int(raw)

        dev = {}
        for key, label in (("desktop_pct", "Desktop %"), ("mobile_pct", "Mobile %")):
            raw = form.get(key, "").replace("%", "").strip()
            if raw:
                if not raw.isdigit() or int(raw) > 100:
                    raise ValueError(f"{label} must be a whole number from 0 to 100.")
                dev[key] = int(raw)
        if len(dev) == 1:  # one value given -> the other makes up the rest
            only = next(iter(dev))
            dev["mobile_pct" if only == "desktop_pct" else "desktop_pct"] = 100 - dev[only]
        if dev and dev["desktop_pct"] + dev["mobile_pct"] != 100:
            raise ValueError("Desktop % and Mobile % must add up to 100.")
        inputs.update(dev)

        if inputs["delivered"] > inputs["sent"]:
            raise ValueError("Delivered cannot be greater than Sent.")
        if pd.to_datetime(inputs["end_date"]) < pd.to_datetime(inputs["start_date"]):
            raise ValueError("End Date cannot be before Start Date.")

        upload = request.files.get("excel")
        if not upload or not upload.filename:
            raise ValueError("Please upload the raw leads Excel file.")
        suffix = Path(upload.filename).suffix.lower()
        if suffix not in ALLOWED_EXT:
            raise ValueError("Upload an .xlsx, .xls, .ods or .csv file.")
    except ValueError as exc:
        return jsonify(ok=False, error=str(exc)), 400

    REPORTS_DIR.mkdir(exist_ok=True)
    with tempfile.TemporaryDirectory() as tmp:
        excel_path = Path(tmp) / f"leads{suffix}"
        upload.save(excel_path)
        try:
            output = build_report(Path(DEFAULT_TEMPLATE), excel_path, inputs, output_dir=REPORTS_DIR)
        except Exception as exc:  # surface parsing/template problems to the page
            return jsonify(ok=False, error=f"Could not build the report: {exc}"), 500

    return jsonify(ok=True, name=output.name, url=f"/reports/{output.name}")


@app.get("/reports/<path:name>")
def report(name: str):
    if not (REPORTS_DIR / name).is_file():
        abort(404)
    return send_from_directory(REPORTS_DIR, name, as_attachment=request.args.get("download") == "1")


if __name__ == "__main__":
    app.run(debug=True)