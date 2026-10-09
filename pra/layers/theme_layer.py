from __future__ import annotations


THEME_LAYER_START = "<!-- PRA-THEME-LAYER:START -->"

THEME_LAYER_END = "<!-- PRA-THEME-LAYER:END -->"

THEME_LAYER_CSS = r""".pra-theme-btn{border:1px solid #dfe5ec;background:#fff;color:#344054;border-radius:10px;height:38px;padding:0 13px;font:800 11px Arial,sans-serif;letter-spacing:.04em;cursor:pointer;box-shadow:0 2px 7px rgba(16,24,40,.04);white-space:nowrap}
.pra-theme-btn:hover{border-color:#f47b20;color:#f47b20}
html[data-theme="dark"]{color-scheme:dark;--line:#2a3850;--white:#e6edf5;--muted:#9fb0c3}
html[data-theme="dark"],html[data-theme="dark"] body{background:#0e1520!important;color:#e6edf5!important}
html[data-theme="dark"] .toolbar{background:rgba(15,23,35,.98)!important;border-bottom:1px solid #243044!important;box-shadow:0 3px 16px rgba(0,0,0,.4)!important}
html[data-theme="dark"] .brand-divider{background:#2a3850}
html[data-theme="dark"] .meta-block{border-right-color:#2a3850}
html[data-theme="dark"] .meta-icon{background:rgba(244,81,50,.16)}
html[data-theme="dark"] .brand-copy h1,html[data-theme="dark"] .title,html[data-theme="dark"] .cover .title,html[data-theme="dark"] .cover .campaign,html[data-theme="dark"] .audience-stat b,html[data-theme="dark"] .section,html[data-theme="dark"] .meta-block b,html[data-theme="dark"] .brand h1{color:#f1f5f9!important}
html[data-theme="dark"] .brand-copy p,html[data-theme="dark"] .subtitle,html[data-theme="dark"] .metric .label,html[data-theme="dark"] .metric .note,html[data-theme="dark"] .meta-block small,html[data-theme="dark"] .slide-indicator,html[data-theme="dark"] .cover .prepared,html[data-theme="dark"] .slideNo,html[data-theme="dark"] .toc-row span:last-child{color:#9fb0c3!important}
html[data-theme="dark"] .slideNo{color:#fff!important}
html[data-theme="dark"] .slide{background:linear-gradient(145deg,#182233 0%,#131c2a 76%,#1b1d29 100%)!important;border-color:#26334a!important;box-shadow:0 16px 48px rgba(0,0,0,.45)!important}
html[data-theme="dark"] .cover{background:linear-gradient(135deg,#182233 0%,#131c2a 58%,#241d22 100%)!important}
html[data-theme="dark"] .slide:before{border-color:rgba(244,123,32,.14)}
html[data-theme="dark"] .panel{background:rgba(26,37,54,.92)!important;border-color:#2a3850!important;box-shadow:0 8px 22px rgba(0,0,0,.30)!important}
html[data-theme="dark"] .callout{background:rgba(244,123,32,.10)!important;color:#d7dee8!important}
html[data-theme="dark"] .table th{color:#9fb0c3!important}
html[data-theme="dark"] .table td,html[data-theme="dark"] .observation-list{color:#d0d9e4!important}
html[data-theme="dark"] .table th,html[data-theme="dark"] .table td{border-bottom-color:#263349!important}
html[data-theme="dark"] .pill,html[data-theme="dark"] .geo-item{background:#1d2a3d!important;border-color:#2a3850!important;color:#c5d0de!important}
html[data-theme="dark"] .toc-row{border-bottom-color:#263349!important}
html[data-theme="dark"] .legend,html[data-theme="dark"] .legend.left-legend .legend-item{color:#c5d0de!important}
html[data-theme="dark"] .stats-card .label{color:#c5d0de!important}
html[data-theme="dark"] .stats-card .note{color:#9fb0c3!important}
html[data-theme="dark"] .audience-stat small{color:#e6edf5!important}
html[data-theme="dark"] .progress{background:#26334a}
html[data-theme="dark"] .section .section-icon,html[data-theme="dark"] .stats-card .stat-icon{background:rgba(244,123,32,.16)!important;box-shadow:inset 0 0 0 1px rgba(244,123,32,.30)!important}
html[data-theme="dark"] .icon-btn,html[data-theme="dark"] .nav-btn,html[data-theme="dark"] .btn,html[data-theme="dark"] .pra-theme-btn{background:#1d2a3d;border-color:#2a3850;color:#d0d9e4}
html[data-theme="dark"] .icon-btn:hover,html[data-theme="dark"] .nav-btn:hover,html[data-theme="dark"] .pra-theme-btn:hover{border-color:#f47b20;color:#ff9a5a;background:#243349}
html[data-theme="dark"] #geoMap,html[data-theme="dark"] .leaflet-container{background:#1b2a39!important;border-color:#2a3850!important}
html[data-theme="dark"] [style*="color:#64748b"]{color:#9fb0c3!important}
html[data-theme="dark"] [style*="color:#344054"],html[data-theme="dark"] [style*="color:#172033"]{color:#e6edf5!important}
@media print{.pra-theme-btn{display:none!important}}
"""

THEME_LAYER_JS = r"""(function () {
  'use strict';
  var KEY = 'praTheme', root = document.documentElement;

  function saved() { try { return localStorage.getItem(KEY); } catch (e) { return null; } }
  function isDark() { return root.getAttribute('data-theme') === 'dark'; }

  function themeCharts() {
    if (!window.Chart) return;
    var dark = isDark();
    var tick = dark ? '#aab6c5' : '#64748b';
    var grid = dark ? 'rgba(255,255,255,.09)' : '#edf0f4';
    try { Chart.defaults.color = tick; } catch (e) {}
    var list = [];
    try { list = Chart.instances ? Object.values(Chart.instances) : []; } catch (e) {}
    list.forEach(function (c) {
      try {
        var o = c.options || {};
        Object.keys(o.scales || {}).forEach(function (k) {
          var s = o.scales[k];
          if (!s) return;
          s.ticks = s.ticks || {}; s.ticks.color = tick;
          s.grid = s.grid || {};
          if (s.grid.display !== false) s.grid.color = grid;
          if (s.title && s.title.display) s.title.color = tick;
        });
        if (o.plugins && o.plugins.legend && o.plugins.legend.labels) o.plugins.legend.labels.color = tick;
        c.update('none');
      } catch (e) {}
    });
  }

  function label() { return isDark() ? '\u2600 Light' : '\u263E Dark'; }

  function apply(theme, persist) {
    if (theme === 'dark') root.setAttribute('data-theme', 'dark'); else root.setAttribute('data-theme', 'light');
    if (persist) { try { localStorage.setItem(KEY, theme); } catch (e) {} }
    var b = document.getElementById('praThemeBtn');
    if (b) { b.textContent = label(); b.setAttribute('aria-pressed', String(isDark())); }
    themeCharts();
  }

  function addButton() {
    var host = document.querySelector('.report-meta');
    if (!host || document.getElementById('praThemeBtn')) return;
    var b = document.createElement('button');
    b.id = 'praThemeBtn'; b.type = 'button'; b.className = 'pra-theme-btn';
    b.title = 'Switch between light and dark mode';
    b.textContent = label();
    b.addEventListener('click', function () { apply(isDark() ? 'light' : 'dark', true); });
    host.insertBefore(b, host.querySelector('.icon-btn') || null);
  }

  // Start in the saved theme (light by default).
  apply(saved() === 'light' ? 'light' : 'dark', false);

  function init() {
    addButton();
    apply(isDark() ? 'dark' : 'light', false);
    // Charts are created at different moments by the template and the data layer, so re-theme a few times.
    [400, 1200, 2800].forEach(function (ms) { setTimeout(themeCharts, ms); });
  }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init); else init();
  window.addEventListener('load', function () { setTimeout(themeCharts, 300); });
  // Slide changes can reveal charts that were rebuilt after load.
  document.addEventListener('click', function () { setTimeout(themeCharts, 150); });
  document.addEventListener('keydown', function () { setTimeout(themeCharts, 150); });
})();
"""

def theme_layer() -> str:
    """Light / dark mode toggle. Light is the default; the choice is remembered in the browser."""
    return (
        THEME_LAYER_START + "\n"
        "<style>" + THEME_LAYER_CSS + "</style>\n"
        "<script>" + THEME_LAYER_JS + "</script>\n"
        + THEME_LAYER_END + "\n"
    )