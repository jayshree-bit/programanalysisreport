from __future__ import annotations

import pandas as pd
import re
from typing import Any

from pra.helpers import esc, pct, pct_text, series_clean


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