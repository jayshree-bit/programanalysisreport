from __future__ import annotations

import json


DESIGN_LAYER_START = "<!-- PRA-DESIGN-LAYER:START -->"

DESIGN_LAYER_END = "<!-- PRA-DESIGN-LAYER:END -->"

# Report designs. "classic" is the template as-is; the others restyle the same slides (layout, data and
# charts are unchanged), each with a light and a dark variant so the Light/Dark button keeps working.
DESIGNS = {
    "classic": "Classic",
    "corporate": "Corporate",
    "aurora": "Aurora",
    "minimal": "Minimal",
    "executive": "Executive",
}

DESIGN_LAYER_CSS = r"""
.pra-design-sel{appearance:none;-webkit-appearance:none;min-width:128px;padding-right:30px!important;background-image:linear-gradient(45deg,transparent 50%,currentColor 50%),linear-gradient(135deg,currentColor 50%,transparent 50%)!important;background-position:calc(100% - 13px) 16px,calc(100% - 8px) 16px!important;background-size:5px 5px!important;background-repeat:no-repeat!important}
@media print{.pra-design-sel{display:none!important}}

/* ---------- Corporate: navy + gold, serif headings, square cards ---------- */
html[data-design="corporate"] body:before{background:#e8ecf2!important}
html[data-design="corporate"] body:after{opacity:0!important}
html[data-design="corporate"] .toolbar{background:#0f2742!important;border-bottom:3px solid #c8a24a!important}
html[data-design="corporate"] .brand-copy h1{color:#fff!important}
html[data-design="corporate"] .brand-copy p,html[data-design="corporate"] .meta-block small{color:#b9c6d6!important}
html[data-design="corporate"] .meta-block b{color:#fff!important}
html[data-design="corporate"] .slide{background:#ffffff!important;border:0!important;border-top:6px solid #0f2742!important;border-radius:0!important;box-shadow:0 10px 30px rgba(15,39,66,.14)!important}
html[data-design="corporate"] .slide:before{display:none}
html[data-design="corporate"] .title{font-family:Georgia,"Times New Roman",serif!important;font-weight:700!important;color:#0f2742!important;align-self:flex-start;border-bottom:3px solid #c8a24a;padding-bottom:8px}
html[data-design="corporate"] .panel{background:#f7f9fc!important;border:1px solid #d9e1ea!important;border-top:3px solid #c8a24a!important;border-radius:0!important;box-shadow:none!important}
html[data-design="corporate"] .section,html[data-design="corporate"] .pdf-panel h4,html[data-design="corporate"] .pdf-kpi h4{color:#0f2742!important;font-family:Georgia,"Times New Roman",serif}
html[data-design="corporate"] .footer{color:#c8a24a!important}
html[data-design="corporate"] .slideNo{background:#0f2742!important;color:#fff!important}
html[data-design="corporate"]{--ink:#1f2f44}
html[data-design="corporate"] .pdf-kpi .num{color:#0f2742!important;text-shadow:none!important}
html[data-design="corporate"][data-theme="dark"] body:before{background:#08131f!important}
html[data-design="corporate"][data-theme="dark"] .slide{background:#10233a!important;box-shadow:0 14px 40px rgba(0,0,0,.5)!important}
html[data-design="corporate"][data-theme="dark"] .title{color:#f3f6fa!important}
html[data-design="corporate"][data-theme="dark"] .panel{background:#15304f!important;border-color:#24456b!important;border-top-color:#c8a24a!important}
html[data-design="corporate"][data-theme="dark"] .section,html[data-design="corporate"][data-theme="dark"] .pdf-panel h4,html[data-design="corporate"][data-theme="dark"] .pdf-kpi h4{color:#e9d9a8!important}
html[data-design="corporate"][data-theme="dark"]{--ink:#e6edf5}
html[data-design="corporate"][data-theme="dark"] .pdf-kpi .num{color:#e9c46a!important}

/* ---------- Aurora: violet-to-teal gradients, rounded cards ---------- */
html[data-design="aurora"] body:before{background:linear-gradient(135deg,#ede9fe 0%,#e0f2fe 55%,#ccfbf1 100%)!important}
html[data-design="aurora"] body:after{opacity:0!important}
html[data-design="aurora"] .toolbar{background:linear-gradient(90deg,#4c1d95,#1e40af 60%,#0e7490)!important;border-bottom:0!important}
html[data-design="aurora"] .brand-copy h1,html[data-design="aurora"] .meta-block b{color:#fff!important}
html[data-design="aurora"] .brand-copy p,html[data-design="aurora"] .meta-block small{color:#c7d2fe!important}
html[data-design="aurora"] .slide{background:linear-gradient(160deg,#ffffff 0%,#f5f3ff 55%,#ecfeff 100%)!important;border:1px solid #ddd6fe!important;border-radius:24px!important;box-shadow:0 24px 60px rgba(76,29,149,.18)!important}
html[data-design="aurora"] .slide:before{border-color:rgba(124,58,237,.18)!important}
html[data-design="aurora"] .title{color:#4c1d95!important;font-family:"Trebuchet MS","Segoe UI",sans-serif!important;font-weight:800!important}
html[data-design="aurora"] .panel{background:rgba(255,255,255,.9)!important;border:1px solid #e9d5ff!important;border-radius:18px!important;box-shadow:0 10px 26px rgba(76,29,149,.10)!important}
html[data-design="aurora"] .section,html[data-design="aurora"] .pdf-panel h4,html[data-design="aurora"] .pdf-kpi h4{color:#6d28d9!important}
html[data-design="aurora"] .footer{color:#7c3aed!important}
html[data-design="aurora"] .slideNo{background:linear-gradient(135deg,#7c3aed,#0891b2)!important;color:#fff!important;border-radius:999px!important}
html[data-design="aurora"]{--ink:#312e81}
html[data-design="aurora"] .pdf-kpi .num{color:#7c3aed!important;text-shadow:0 0 14px rgba(124,58,237,.35)!important}
html[data-design="aurora"][data-theme="dark"] body:before{background:linear-gradient(135deg,#1e1b4b 0%,#172554 55%,#083344 100%)!important}
html[data-design="aurora"][data-theme="dark"] .slide{background:linear-gradient(160deg,#241f5c 0%,#1a1f4d 50%,#0b3442 100%)!important;border-color:rgba(167,139,250,.35)!important;box-shadow:0 24px 60px rgba(0,0,0,.45)!important}
html[data-design="aurora"][data-theme="dark"] .title{color:#e0e7ff!important;text-shadow:0 0 18px rgba(167,139,250,.45)}
html[data-design="aurora"][data-theme="dark"] .panel{background:rgba(255,255,255,.06)!important;border-color:rgba(196,181,253,.22)!important;box-shadow:0 10px 26px rgba(0,0,0,.3)!important}
html[data-design="aurora"][data-theme="dark"] .section,html[data-design="aurora"][data-theme="dark"] .pdf-panel h4,html[data-design="aurora"][data-theme="dark"] .pdf-kpi h4{color:#c4b5fd!important}
html[data-design="aurora"][data-theme="dark"]{--ink:#e0e7ff}
html[data-design="aurora"][data-theme="dark"] .pdf-kpi .num{color:#5eead4!important;text-shadow:0 0 14px rgba(94,234,212,.45)!important}

/* ---------- Minimal: white space, hairlines, no shadows ---------- */
html[data-design="minimal"] body:before{background:#f4f4f2!important}
html[data-design="minimal"] body:after{opacity:0!important}
html[data-design="minimal"] .toolbar{background:#ffffff!important;border-bottom:1px solid #e7e7e4!important;box-shadow:none!important}
html[data-design="minimal"] .slide{background:#ffffff!important;border:1px solid #e7e7e4!important;border-radius:6px!important;box-shadow:none!important}
html[data-design="minimal"] .slide:before{display:none}
html[data-design="minimal"] .title{font-family:"Segoe UI",Helvetica,Arial,sans-serif!important;font-weight:300!important;font-size:34px!important;letter-spacing:-.01em!important;color:#111!important}
html[data-design="minimal"] .panel{background:#ffffff!important;border:1px solid #ececea!important;border-radius:10px!important;box-shadow:none!important}
html[data-design="minimal"] .section,html[data-design="minimal"] .pdf-panel h4,html[data-design="minimal"] .pdf-kpi h4{color:#555!important;font-weight:600!important;letter-spacing:.14em!important}
html[data-design="minimal"] .footer{color:#999!important}
html[data-design="minimal"] .slideNo{background:transparent!important;color:#999!important;box-shadow:none!important}
html[data-design="minimal"]{--ink:#2b2b2b}
html[data-design="minimal"] .pdf-kpi .num{color:#111!important;text-shadow:none!important;font-family:"Segoe UI",Helvetica,Arial,sans-serif!important;font-weight:300!important}
html[data-design="minimal"] .ibadge{filter:grayscale(.35)}
html[data-design="minimal"][data-theme="dark"] body:before{background:#0d0d0d!important}
html[data-design="minimal"][data-theme="dark"] .toolbar{background:#141414!important;border-bottom-color:#262626!important}
html[data-design="minimal"][data-theme="dark"] .slide{background:#171717!important;border-color:#262626!important}
html[data-design="minimal"][data-theme="dark"] .title{color:#f5f5f5!important}
html[data-design="minimal"][data-theme="dark"] .panel{background:#1c1c1c!important;border-color:#2a2a2a!important}
html[data-design="minimal"][data-theme="dark"] .section,html[data-design="minimal"][data-theme="dark"] .pdf-panel h4,html[data-design="minimal"][data-theme="dark"] .pdf-kpi h4{color:#a3a3a3!important}
html[data-design="minimal"][data-theme="dark"]{--ink:#e5e5e5}
html[data-design="minimal"][data-theme="dark"] .pdf-kpi .num{color:#fafafa!important}

/* ---------- Executive: charcoal + gold, accent bar on titles ---------- */
html[data-design="executive"] body:before{background:#efe9dd!important}
html[data-design="executive"] body:after{opacity:0!important}
html[data-design="executive"] .toolbar{background:#1b1a17!important;border-bottom:2px solid #b8913a!important}
html[data-design="executive"] .brand-copy h1,html[data-design="executive"] .meta-block b{color:#f5e6c4!important}
html[data-design="executive"] .brand-copy p,html[data-design="executive"] .meta-block small{color:#a89f8c!important}
html[data-design="executive"] .slide{background:#fffdf8!important;border:1px solid #e6dcc6!important;border-radius:4px!important;box-shadow:0 18px 44px rgba(60,45,15,.16)!important}
html[data-design="executive"] .slide:before{border-color:rgba(184,145,58,.18)!important}
html[data-design="executive"] .title{color:#1b1a17!important;border-left:6px solid #b8913a;padding-left:14px;font-family:Georgia,"Times New Roman",serif!important}
html[data-design="executive"] .panel{background:#ffffff!important;border:1px solid #eadfc6!important;border-radius:6px!important;box-shadow:0 4px 14px rgba(60,45,15,.06)!important}
html[data-design="executive"] .section,html[data-design="executive"] .pdf-panel h4,html[data-design="executive"] .pdf-kpi h4{color:#8a6a24!important}
html[data-design="executive"] .footer{color:#b8913a!important}
html[data-design="executive"] .slideNo{background:#b8913a!important;color:#1b1a17!important}
html[data-design="executive"]{--ink:#2d2a24}
html[data-design="executive"] .pdf-kpi .num{color:#b8913a!important;text-shadow:none!important}
html[data-design="executive"][data-theme="dark"] body:before{background:#0b0b0c!important}
html[data-design="executive"][data-theme="dark"] .slide{background:linear-gradient(160deg,#1c1b18,#141413)!important;border-color:#3a3226!important;box-shadow:0 18px 44px rgba(0,0,0,.55)!important}
html[data-design="executive"][data-theme="dark"] .title{color:#f5e6c4!important}
html[data-design="executive"][data-theme="dark"] .panel{background:#211f1b!important;border-color:#3a3226!important}
html[data-design="executive"][data-theme="dark"] .section,html[data-design="executive"][data-theme="dark"] .pdf-panel h4,html[data-design="executive"][data-theme="dark"] .pdf-kpi h4{color:#d9b864!important}
html[data-design="executive"][data-theme="dark"]{--ink:#ece3cf}
html[data-design="executive"][data-theme="dark"] .pdf-kpi .num{color:#e3c06b!important;text-shadow:0 0 12px rgba(227,192,107,.35)!important}

/* ---------- Per-design accent colours (side bar, progress, lists, callouts, big numbers, icons) ---------- */
html[data-design="corporate"]{--acc:#c8a24a;--acc2:#0f2742;--acc-soft:rgba(200,162,74,.12)}
html[data-design="aurora"]{--acc:#7c3aed;--acc2:#06b6d4;--acc-soft:rgba(124,58,237,.10)}
html[data-design="minimal"]{--acc:#111111;--acc2:#111111;--acc-soft:rgba(0,0,0,.04)}
html[data-design="minimal"][data-theme="dark"]{--acc:#e5e5e5;--acc2:#e5e5e5;--acc-soft:rgba(255,255,255,.05)}
html[data-design="executive"]{--acc:#b8913a;--acc2:#1b1a17;--acc-soft:rgba(184,145,58,.10)}
html[data-design="executive"][data-theme="dark"]{--acc2:#e3c06b}
html[data-design]:not([data-design="classic"]) .slide:after{background:linear-gradient(180deg,var(--acc),var(--acc2))!important}
html[data-design]:not([data-design="classic"]) .progress>div{background:linear-gradient(90deg,var(--acc),var(--acc2))!important}
html[data-design]:not([data-design="classic"]) .geo-item{border-left-color:var(--acc)!important}
html[data-design]:not([data-design="classic"]) .callout{border-left-color:var(--acc)!important;background:var(--acc-soft)!important}
html[data-design]:not([data-design="classic"]) .metric .value{color:var(--acc)!important}
html[data-design]:not([data-design="classic"]) .section .section-icon,html[data-design]:not([data-design="classic"]) .stats-card .stat-icon{color:var(--acc)!important;background:var(--acc-soft)!important;box-shadow:inset 0 0 0 1px var(--acc-soft)!important}
html[data-design="corporate"] .slide-indicator,html[data-design="aurora"] .slide-indicator,html[data-design="executive"] .slide-indicator{color:#fff!important;opacity:.85}
html[data-design="minimal"] .slide:after{display:none}
html[data-design="minimal"] .pdf-obs li:before,html[data-design="corporate"] .pdf-obs li:before,html[data-design="executive"] .pdf-obs li:before,html[data-design="aurora"] .pdf-obs li:before{color:var(--acc)!important}
html[data-design="aurora"] .panel{border-top:4px solid transparent!important;background-image:linear-gradient(rgba(255,255,255,.92),rgba(255,255,255,.92)),linear-gradient(90deg,#7c3aed,#06b6d4)!important;background-origin:border-box!important;background-clip:padding-box,border-box!important}
html[data-design="aurora"][data-theme="dark"] .panel{background-image:linear-gradient(#1f2150,#1f2150),linear-gradient(90deg,#7c3aed,#06b6d4)!important}
html[data-design="executive"] .panel{border-left:3px solid #b8913a!important}
html[data-design="corporate"] .pdf-track{border-radius:0!important;height:24px!important;box-shadow:none!important}
html[data-design="corporate"] .pdf-fill{border-radius:0!important;box-shadow:none!important}
html[data-design="minimal"] .pdf-track{height:8px!important;border-radius:4px!important;box-shadow:none!important}
html[data-design="minimal"] .pdf-fill{border-radius:4px!important;box-shadow:none!important;min-width:6px!important}
html[data-design="executive"] .pdf-track{height:12px!important;border-radius:0!important;box-shadow:none!important}
html[data-design="executive"] .pdf-fill{border-radius:0!important;box-shadow:none!important}
html[data-design="corporate"] .pdf-fill:after,html[data-design="minimal"] .pdf-fill:after,html[data-design="executive"] .pdf-fill:after{display:none!important}
html[data-design="aurora"] .pdf-track{height:18px!important;box-shadow:none!important}
html[data-design="aurora"] .pdf-fill{box-shadow:0 4px 12px rgba(124,58,237,.35)!important}
html[data-design="aurora"] .pdf-fill:after{display:none!important}
html[data-design="minimal"] .pdf-dev .badge,html[data-design="corporate"] .pdf-dev .badge,html[data-design="executive"] .pdf-dev .badge{background:var(--acc)!important;box-shadow:none!important}
html[data-design="minimal"] .pdf-dev .big,html[data-design="corporate"] .pdf-dev .big,html[data-design="executive"] .pdf-dev .big,html[data-design="aurora"] .pdf-dev .big{text-shadow:none!important}
html[data-design="corporate"][data-theme="dark"],html[data-design="corporate"][data-theme="dark"] body{background:#08131f!important}
html[data-design="aurora"][data-theme="dark"],html[data-design="aurora"][data-theme="dark"] body{background:#151339!important}
html[data-design="minimal"][data-theme="dark"],html[data-design="minimal"][data-theme="dark"] body{background:#0d0d0d!important}
html[data-design="executive"][data-theme="dark"],html[data-design="executive"][data-theme="dark"] body{background:#0b0b0c!important}
html[data-design="corporate"]:not([data-theme="dark"]),html[data-design="corporate"]:not([data-theme="dark"]) body{background:#e8ecf2!important}
html[data-design="aurora"]:not([data-theme="dark"]),html[data-design="aurora"]:not([data-theme="dark"]) body{background:#eef2ff!important}
html[data-design="minimal"]:not([data-theme="dark"]),html[data-design="minimal"]:not([data-theme="dark"]) body{background:#f4f4f2!important}
html[data-design="executive"]:not([data-theme="dark"]),html[data-design="executive"]:not([data-theme="dark"]) body{background:#efe9dd!important}
"""

DESIGN_LAYER_JS = r"""(function () {
  'use strict';
  var NAMES = __DESIGN_NAMES__, DEF = '__DESIGN_DEFAULT__', root = document.documentElement;
  var KEY = 'praDesign:' + location.pathname;
  function saved() { try { return localStorage.getItem(KEY); } catch (e) { return null; } }
  function apply(d, persist) {
    if (!NAMES[d]) d = DEF;
    root.setAttribute('data-design', d);
    if (persist) { try { localStorage.setItem(KEY, d); } catch (e) {} }
    var sel = document.getElementById('praDesignSel'); if (sel) sel.value = d;
  }
  apply(saved() || DEF, false);
  function addPicker() {
    var host = document.querySelector('.report-meta');
    if (!host || document.getElementById('praDesignSel')) return;
    var sel = document.createElement('select');
    sel.id = 'praDesignSel'; sel.className = 'pra-theme-btn pra-design-sel'; sel.title = 'Change the report design';
    sel.innerHTML = Object.keys(NAMES).map(function (k) { return '<option value="' + k + '">\uD83C\uDFA8 ' + NAMES[k] + '</option>'; }).join('');
    sel.value = root.getAttribute('data-design');
    sel.addEventListener('change', function () { apply(sel.value, true); if (window.praRebuildLayout) window.praRebuildLayout(); });
    host.insertBefore(sel, document.getElementById('praThemeBtn') || host.querySelector('.icon-btn') || null);
  }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', addPicker); else addPicker();
})();
"""

def design_layer(default: str = "classic") -> str:
    """Selectable report designs plus a design picker in the report toolbar."""
    if default not in DESIGNS:
        default = "classic"
    js = DESIGN_LAYER_JS.replace("__DESIGN_NAMES__", json.dumps(DESIGNS)).replace("__DESIGN_DEFAULT__", default)
    return (
        DESIGN_LAYER_START + "\n<style>" + DESIGN_LAYER_CSS + "</style>\n<script>" + js + "</script>\n"
        + DESIGN_LAYER_END + "\n"
    )