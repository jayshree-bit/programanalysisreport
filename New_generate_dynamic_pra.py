"""
Dynamic Program Analysis Report generator  (entry point / compatibility shim).

The old single 3,000-line script now lives in the `pra/` package:

    pra/config.py              paths, DATE_FORMATS, HEX_COLOR
    pra/helpers.py             small utilities + read_raw_leads()
    pra/cli_prompts.py         interactive question prompts (collect_inputs)
    pra/geo.py                 country / US-state coordinates, map spec
    pra/metrics.py             asset event metrics
    pra/slides.py              summary slide + removing slides from the HTML
    pra/data.py                prepare_data(), apply_display_options()
    pra/js_layer.py            the dynamic JavaScript injected in the report
    pra/layers/pdf_layer.py    PDF layout layer (CSS + JS)
    pra/layers/theme_layer.py  theme switcher layer
    pra/layers/design_layer.py design styles layer (DESIGNS)
    pra/layers/export_layer.py download PDF / PPT layer
    pra/report.py              build_report()
    pra/cli.py                 command line main()

Everything is re-exported here so `app.py` and old commands keep working:
    python New_generate_dynamic_pra.py ...
"""
from pra.config import BASE_DIR, DATE_FORMATS, DEFAULT_EXCEL, DEFAULT_TEMPLATE, HEX_COLOR  # noqa: F401
from pra.layers.design_layer import DESIGNS  # noqa: F401
from pra.report import build_report  # noqa: F401
from pra.cli import main

if __name__ == "__main__":
    main()