from __future__ import annotations

import re
from pathlib import Path
from typing import Any

from pra.data import apply_display_options, prepare_data
from pra.geo import build_plotly_geo_spec
from pra.helpers import esc, fmt, pct_text, read_raw_leads, slugify
from pra.js_layer import js_dynamic_layer
from pra.layers.design_layer import DESIGN_LAYER_END, DESIGN_LAYER_START, design_layer
from pra.layers.export_layer import EXPORT_LAYER_END, EXPORT_LAYER_START, export_layer
from pra.layers.pdf_layer import PDF_LAYER_END, PDF_LAYER_START, pdf_layout_layer
from pra.layers.theme_layer import THEME_LAYER_END, THEME_LAYER_START, theme_layer
from pra.slides import remove_single_asset_slides, remove_summary_slide


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
        re.escape(DESIGN_LAYER_START) + r".*?" + re.escape(DESIGN_LAYER_END) + r"\s*",
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
        dynamic_layer + "\n" + theme_layer() + design_layer(inputs.get("design", "classic"))
        + pdf_layout_layer() + export_layer() + marker,
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