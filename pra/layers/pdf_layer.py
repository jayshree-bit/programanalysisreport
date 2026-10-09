from __future__ import annotations


# ---------------------------------------------------------------------------
# Download-as-PDF / Download-as-PPT layer (runs in the browser, no server needed)
# ---------------------------------------------------------------------------
PDF_LAYER_START = "<!-- PRA-PDFLAYOUT-LAYER:START -->"

PDF_LAYER_END = "<!-- PRA-PDFLAYOUT-LAYER:END -->"

PDF_LAYER_CSS = r"""#praModeSwitch{display:none!important}
.pdfx .slide .kicker{display:none}
.pdfx .slide{padding-bottom:82px!important}
.pdfx #s1.slide .title{text-transform:uppercase;font-size:68px!important;letter-spacing:.01em;margin-bottom:14px!important}
@media(max-width:760px){.pdfx #s1.slide .title{font-size:52px!important}}
html[data-theme="dark"] .slide{background:linear-gradient(rgba(33,41,54,.93),rgba(27,35,47,.96)),repeating-linear-gradient(90deg,rgba(255,255,255,.04) 0 38px,transparent 38px 70px,rgba(255,255,255,.022) 70px 96px,transparent 96px 140px)!important;background-color:#222a36!important}
html[data-theme="dark"] .cover{background:#222833!important}
html[data-theme="dark"]{--trk:rgba(0,0,0,.32);--ink:#e6edf5}
html[data-theme="light"]{--trk:rgba(23,32,51,.09);--ink:#334155}
.pdf-grid{display:grid;gap:14px;margin-bottom:14px}
.pdf-panel h4,.pdf-kpi h4{margin:0 0 8px;text-align:center;font-size:14px;letter-spacing:.09em;text-transform:uppercase;color:var(--ink)}
.pdf-kpi{display:flex;flex-direction:column;align-items:center;justify-content:center;gap:10px;text-align:center;min-height:200px}
.pdf-kpi .num{font:900 54px/1 "Courier New",monospace;color:#12e7d4;text-shadow:0 0 12px rgba(18,231,212,.5)}
html[data-theme="light"] .pdf-kpi .num{color:#0e9aa7;text-shadow:none}
.pdf-cv{position:relative;width:100%}
.pdf-hbar{display:grid;grid-template-columns:96px 1fr 56px;align-items:center;gap:12px;margin:10px 0;font-size:13px;color:var(--ink);font-weight:700}
.pdf-hbar .lb,.pdf-hbar span:first-child{display:flex;align-items:center;justify-content:flex-end;gap:9px;text-align:right}
.pdf-hbar .vl,.pdf-hbar span:last-child{font:800 14px Arial,sans-serif}
.pdf-track{height:30px;background:var(--trk);border-radius:16px;overflow:hidden;box-shadow:inset 0 3px 7px rgba(0,0,0,.42),0 1px 0 rgba(255,255,255,.06)}
.pdf-fill{position:relative;height:100%;min-width:18px;border-radius:16px;box-shadow:inset 0 2px 0 rgba(255,255,255,.42),inset 0 -6px 9px rgba(0,0,0,.24)}
.pdf-fill:after{content:"";position:absolute;left:10px;right:14px;top:4px;height:9px;border-radius:9px;background:linear-gradient(rgba(255,255,255,.42),rgba(255,255,255,0))}
.ibadge{border-radius:50%;display:flex;align-items:center;justify-content:center;flex:none}
.ibadge svg{width:100%;height:100%;display:block;filter:drop-shadow(0 2px 2px rgba(0,0,0,.4))}
.pdf-kpi .ic{display:flex;justify-content:center;font-size:inherit}
.pdf-stat .ic{display:flex;justify-content:center;margin-bottom:8px;font-size:inherit}
.pdf-dev{display:flex;flex-direction:column;justify-content:center;gap:22px;height:100%;min-height:250px;padding:6px 10px 2px;color:var(--ink)}
.pdf-dev .tiles{display:flex;justify-content:space-around;align-items:flex-start;gap:14px}
.pdf-dev .tile{flex:1;min-width:0;display:flex;flex-direction:column;align-items:center;gap:7px;text-align:center}
.pdf-dev .badge{width:104px;height:104px;border-radius:50%;display:flex;align-items:center;justify-content:center;margin-bottom:4px}
.pdf-dev .badge svg{width:60px;height:60px;filter:drop-shadow(0 3px 3px rgba(0,0,0,.4))}
.pdf-dev .big{font:900 40px/1.05 "Courier New",monospace;letter-spacing:.01em}
html[data-theme="light"] .pdf-dev .big{filter:brightness(.72) saturate(1.25);text-shadow:none!important}
.pdf-dev .nm{font-size:16px;font-weight:800;letter-spacing:.03em}
.pdf-dev .ct{font-size:11px;font-weight:700;opacity:.65;letter-spacing:.05em;text-transform:uppercase}
.pdf-dev .bar{display:flex;height:30px;border-radius:16px;overflow:hidden;background:var(--trk);margin:0 6px;box-shadow:inset 0 3px 7px rgba(0,0,0,.4),0 10px 16px -10px rgba(0,0,0,.65)}
.pdf-dev .bar span{display:flex;align-items:center;justify-content:center;font:900 12px Arial,sans-serif;color:#fff;text-shadow:0 1px 3px rgba(0,0,0,.6);border-right:2px solid rgba(0,0,0,.28);box-shadow:inset 0 2px 0 rgba(255,255,255,.35)}
.pdf-dev .bar span:last-child{border-right:0}
@media(max-width:760px){.pdf-dev .badge{width:76px;height:76px}.pdf-dev .badge svg{width:44px;height:44px}.pdf-dev .big{font-size:28px}.pdf-dev .nm{font-size:13px}}
.pdf-stats{display:grid;grid-template-columns:repeat(6,1fr);gap:10px;margin-bottom:14px}
.pdf-stat{text-align:center;color:var(--ink)}.pdf-stat b{display:block;font-size:30px}.pdf-stat small{font-size:12px;font-weight:800;letter-spacing:.05em}
.pdf-obs{margin:0;padding:6px 10px;list-style:none;color:var(--ink);font-size:14px;line-height:1.75}
.pdf-obs li{padding-left:22px;position:relative;margin-bottom:6px}.pdf-obs li:before{content:"\27A2";position:absolute;left:0;color:#f47b20}
.pdf-thanks{flex:1;display:flex;align-items:center;justify-content:center;font-size:56px;font-weight:900;color:var(--ink)}
.pdf-offices{display:grid;grid-template-columns:repeat(3,1fr);gap:20px;color:var(--ink);font-size:12px;line-height:1.6;padding:0 40px 40px}
.pdf-offices b{display:block;font-size:13px}
#s14 .pdf-thanks{flex:none;margin:14px 0 4px;font-size:52px}
#praOfficeMap{flex:1;display:flex;align-items:center;justify-content:center;padding:0 40px 14px;min-height:0}
#praOfficeMap svg{width:100%;max-height:380px;display:block}
#geoMap{height:300px!important}
.pdf-svg{color:var(--ink)}
.pdf-leg{display:flex;flex-wrap:wrap;gap:6px 16px;justify-content:center;margin-top:8px;font-size:11px;font-weight:700;color:var(--ink)}
.pdf-leg i{display:inline-block;width:12px;height:12px;margin-right:6px;vertical-align:-2px;border-radius:4px;box-shadow:inset 0 2px 0 rgba(255,255,255,.4),0 2px 4px rgba(0,0,0,.35)}
.pdf-panel{display:flex;flex-direction:column}.pdf-panel>.pdf-svg{margin:auto 0}
.pdf-hbar .lb{justify-content:flex-start!important;text-align:left!important;padding-left:14px}
"""

PDF_LAYER_JS = r"""(function () {
  'use strict';
  function build() {
    var P = (typeof PRA !== 'undefined') ? PRA : window.PRA; if (!P) return;
    /* Each report design gets its own chart style: bar shape, pie/donut form and colours. */
    var DESIGN = document.documentElement.getAttribute('data-design') || 'classic';
    var THEMES = {
      classic: { bar: '3d', pie: '3d', donut: '3d', prim: ['#1d4ed8', '#38bdf8'], sec: ['#7c3aed', '#ff3d8b'], one: '#4472c4',
        split: [['#a21caf', '#f0f'], ['#0284c7', '#22e5ff'], ['#f59e0b', '#fde047'], ['#ea580c', '#fbbf24']], donutC: ['#12e7d4', '#ff9a1f'],
        dev: ['#12e7d4', '#ff9a1f', '#7c5ce5', '#ff3d8b'], pieC: ['#70ad47', '#ffc000', '#ed7d31', '#4472c4', '#5b9bd5', '#a5a5a5', '#9e480e', '#7c5ce5'],
        series: ['#4472c4', '#ed7d31', '#a5a5a5', '#ffc000', '#5b9bd5', '#70ad47', '#9e480e', '#7c5ce5'] },
      corporate: { bar: 'flat', pie: 'flat', donut: 'flat', prim: ['#2c5282', '#2c5282'], sec: ['#a07a2c', '#c8a24a'], one: '#2c5282',
        split: [['#2c5282', '#4a6fa5'], ['#a07a2c', '#c8a24a'], ['#5b7a99', '#8a9bb0'], ['#8a9bb0', '#b3c0cf']], donutC: ['#2c5282', '#c8a24a'],
        dev: ['#2c5282', '#c8a24a', '#5b7a99', '#8a9bb0'], pieC: ['#2c5282', '#c8a24a', '#5b7a99', '#8a9bb0', '#4a6fa5', '#d9c08a', '#3d5a80', '#a3843f'],
        series: ['#2c5282', '#c8a24a', '#5b7a99', '#8a9bb0', '#4a6fa5', '#d9c08a', '#3d5a80', '#a3843f'], badge: '#2c5282', flatBadge: true },
      aurora: { bar: 'pill', pie: 'ring', donut: 'round', prim: ['#7c3aed', '#06b6d4'], sec: ['#ec4899', '#8b5cf6'], one: '#7c3aed',
        split: [['#7c3aed', '#a78bfa'], ['#06b6d4', '#67e8f9'], ['#ec4899', '#f9a8d4'], ['#22c55e', '#86efac']], donutC: ['#7c3aed', '#06b6d4'],
        dev: ['#7c3aed', '#06b6d4', '#ec4899', '#22c55e'], pieC: ['#7c3aed', '#06b6d4', '#ec4899', '#22c55e', '#f59e0b', '#3b82f6', '#a855f7', '#14b8a6'],
        series: ['#7c3aed', '#06b6d4', '#ec4899', '#22c55e', '#f59e0b', '#3b82f6', '#a855f7', '#14b8a6'], badge: '#7c3aed' },
      minimal: { bar: 'thin', pie: 'ring', donut: 'flat', prim: ['#8a8a8a', '#8a8a8a'], sec: ['#8a8a8a', '#8a8a8a'], one: '#8a8a8a', hl: '#2563eb',
        split: [['#2563eb', '#2563eb'], ['#8a8a8a', '#8a8a8a'], ['#a3a3a3', '#a3a3a3'], ['#bdbdbd', '#bdbdbd']], donutC: ['#2563eb', '#8a8a8a'],
        dev: ['#2563eb', '#8a8a8a', '#a3a3a3', '#5f5f5f'], pieC: ['#2563eb', '#525252', '#737373', '#a3a3a3', '#c4c4c4', '#8a8a8a', '#5f5f5f', '#d4d4d4'],
        series: ['#2563eb', '#737373', '#a3a3a3', '#525252', '#8a8a8a', '#c4c4c4', '#5f5f5f', '#d4d4d4'], badge: '#525252', flatBadge: true },
      executive: { bar: 'lollipop', pie: 'flat', donut: 'flat', prim: ['#8a6a24', '#e3c06b'], sec: ['#8a6a24', '#e3c06b'], one: '#b8913a',
        split: [['#8a6a24', '#e3c06b'], ['#6b5a3a', '#a08a5c'], ['#b8913a', '#d9b864'], ['#8a7a5c', '#c9b68a']], donutC: ['#b8913a', '#8a7a5c'],
        dev: ['#b8913a', '#8a7a5c', '#d9b864', '#5c4a2a'], pieC: ['#b8913a', '#6b5a3a', '#d9b864', '#8a7a5c', '#e3c06b', '#4f4636', '#c9b68a', '#a07a2c'],
        series: ['#b8913a', '#6b5a3a', '#d9b864', '#8a7a5c', '#e3c06b', '#4f4636', '#c9b68a', '#a07a2c'], badge: '#b8913a', flatBadge: true }
    };
    var TH = THEMES[DESIGN] || THEMES.classic, S3D = TH.bar === '3d';
    var COL = (P.palette && P.palette.length) ? P.palette : TH.series;
    var $ = function (id) { return document.getElementById(id); };
    var sum = function (a) { return a.reduce(function (x, y) { return x + (+y || 0); }, 0); };
    var pc = function (v, t, d) { return t ? (v / t * 100).toFixed(d == null ? 0 : d) : '0'; };
    var esc = function (s) { return String(s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); };
    var FONT = 'Calibri, Segoe UI, Arial, sans-serif', uid = 0;
    try { if (window.Chart && Chart.instances) Object.values(Chart.instances).forEach(function (c) { try { c.destroy(); } catch (e) {} }); } catch (e) {}
    document.documentElement.classList.add('pdfx');

    /* ---------- tiny SVG chart kit (no external libraries) ---------- */
    function svg(w, h, inner) { return '<svg class="pdf-svg" viewBox="0 0 ' + w + ' ' + h + '" width="100%" style="display:block;max-height:' + h + 'px" xmlns="http://www.w3.org/2000/svg" font-family="' + FONT + '">' + inner + '</svg>'; }
    function T(x, y, t, o) { o = o || {}; return '<text x="' + x + '" y="' + y + '" text-anchor="' + (o.a || 'middle') + '" font-size="' + (o.s || 12) + '" font-weight="' + (o.w || 700) + '" fill="' + (o.c || 'currentColor') + '">' + esc(t) + '</text>'; }
    function grad(id, c1, c2, vert) { return '<linearGradient id="' + id + '" x1="0" y1="' + (vert ? 1 : 0) + '" x2="' + (vert ? 0 : 1) + '" y2="0"><stop offset="0" stop-color="' + c1 + '"/><stop offset="1" stop-color="' + c2 + '"/></linearGradient>'; }
    function shade(c, f) {
      var m = /^#([0-9a-f]{6})$/i.exec(c || ''); if (!m) return c;
      var n = parseInt(m[1], 16), r = n >> 16, g = (n >> 8) & 255, b = n & 255;
      var t = function (v) { return Math.round(f < 0 ? v * (1 + f) : v + (255 - v) * f); };
      return 'rgb(' + t(r) + ',' + t(g) + ',' + t(b) + ')';
    }
    function rgba(c, a) {
      var m = /^#([0-9a-f]{6})$/i.exec(c || ''); if (!m) return c;
      var n = parseInt(m[1], 16); return 'rgba(' + (n >> 16) + ',' + ((n >> 8) & 255) + ',' + (n & 255) + ',' + a + ')';
    }
    var KC = { target: '#12e7d4', factory: '#ff9a1f', pin: '#ff5a7a', doc: '#7c5ce5', plane: '#3b82f6', inbox: '#10b981', mail: '#06b6d4', cursor: '#8b5cf6', check: '#f59e0b', ban: '#ef4444' };
    function ico(k) {
      var W = '#fff', D = 'rgba(8,16,30,.36)', b = '';
      if (k === 'target') b = '<circle cx="30" cy="34" r="22" fill="none" stroke="' + W + '" stroke-width="5"/><circle cx="30" cy="34" r="12" fill="none" stroke="' + W + '" stroke-width="5"/><circle cx="30" cy="34" r="4.5" fill="' + W + '"/><path d="M32 32L56 8" stroke="' + W + '" stroke-width="4.5" stroke-linecap="round"/><path d="M46 6v12h12" fill="none" stroke="' + W + '" stroke-width="3.5" stroke-linecap="round" stroke-linejoin="round"/>';
      else if (k === 'factory') b = '<path d="M5 57V29l15 9v-9l15 9V9h11v28h13v20z" fill="' + W + '"/><rect x="13" y="45" width="7" height="7" rx="1" fill="' + D + '"/><rect x="28" y="45" width="7" height="7" rx="1" fill="' + D + '"/><rect x="43" y="45" width="7" height="7" rx="1" fill="' + D + '"/>';
      else if (k === 'pin') b = '<path d="M32 3C19.5 3 11 12.5 11 24c0 15.5 21 37 21 37s21-21.5 21-37C53 12.5 44.5 3 32 3z" fill="' + W + '"/><circle cx="32" cy="24" r="8.5" fill="' + D + '"/>';
      else if (k === 'doc') b = '<path d="M13 3h27l15 15v43H13z" fill="' + W + '"/><path d="M40 3v15h15z" fill="' + D + '"/><rect x="21" y="29" width="26" height="4" rx="2" fill="' + D + '"/><rect x="21" y="38" width="26" height="4" rx="2" fill="' + D + '"/><rect x="21" y="47" width="17" height="4" rx="2" fill="' + D + '"/>';
      else if (k === 'plane') b = '<path d="M59 5L6 26l18 7 8 19z" fill="' + W + '"/><path d="M24 33L59 5 31 41z" fill="' + D + '"/>';
      else if (k === 'inbox') b = '<rect x="4" y="12" width="48" height="34" rx="5" fill="' + W + '"/><path d="M7 17l21 16L49 17" fill="none" stroke="' + D + '" stroke-width="3.5" stroke-linecap="round" stroke-linejoin="round"/><circle cx="48" cy="46" r="14" fill="#22c55e" stroke="' + W + '" stroke-width="3"/><path d="M41 46l5 5 9-10" fill="none" stroke="' + W + '" stroke-width="4" stroke-linecap="round" stroke-linejoin="round"/>';
      else if (k === 'mail') b = '<path d="M6 27L32 6l26 21z" fill="' + W + '" fill-opacity=".78"/><rect x="14" y="16" width="36" height="26" rx="2" fill="' + D + '"/><path d="M6 27l26 18 26-18v26a5 5 0 0 1-5 5H11a5 5 0 0 1-5-5z" fill="' + W + '"/>';
      else if (k === 'cursor') b = '<path d="M14 4l36 28-16 3 10 19-9 4-10-19-11 11z" fill="' + W + '" stroke="' + D + '" stroke-width="2" stroke-linejoin="round"/>';
      else if (k === 'check') b = '<circle cx="32" cy="32" r="27" fill="' + W + '"/><path d="M19 33l9.5 9.5L46 22" fill="none" stroke="' + D + '" stroke-width="6.5" stroke-linecap="round" stroke-linejoin="round"/>';
      else if (k === 'ban') b = '<circle cx="32" cy="32" r="25" fill="none" stroke="' + W + '" stroke-width="7"/><path d="M14.5 49.5l35-35" stroke="' + W + '" stroke-width="7" stroke-linecap="round"/>';
      return '<svg viewBox="0 0 64 64" xmlns="http://www.w3.org/2000/svg">' + b + '</svg>';
    }
    function badge(k, c, sz) {
      c = TH.badge || c || KC[k] || '#12e7d4'; var q = function (d) { return Math.round(sz / d); }, g = Math.round(sz * .58);
      if (TH.flatBadge) return '<div class="ibadge" style="width:' + sz + 'px;height:' + sz + 'px;background:' + c + ';box-shadow:0 0 0 ' + q(16) + 'px ' + rgba(c, .14) + '"><div style="width:' + g + 'px;height:' + g + 'px">' + ico(k) + '</div></div>';
      if (DESIGN === 'aurora') return '<div class="ibadge" style="width:' + sz + 'px;height:' + sz + 'px;background:linear-gradient(135deg,#7c3aed,#06b6d4);box-shadow:0 ' + q(6) + 'px ' + q(3) + 'px -' + q(8) + 'px #7c3aed"><div style="width:' + g + 'px;height:' + g + 'px">' + ico(k) + '</div></div>';
      return '<div class="ibadge" style="width:' + sz + 'px;height:' + sz + 'px;background:radial-gradient(circle at 30% 22%,' + shade(c, .55) + ',' + c + ' 52%,' + shade(c, -.4) + ');box-shadow:0 0 0 ' + q(16) + 'px ' + rgba(c, .14) + ',0 ' + q(5) + 'px ' + q(3.2) + 'px -' + q(6) + 'px ' + c + ',inset 0 -' + q(10) + 'px ' + q(5.5) + 'px rgba(0,0,0,.3),inset 0 ' + q(12) + 'px ' + q(7) + 'px rgba(255,255,255,.42)"><div style="width:' + g + 'px;height:' + g + 'px">' + ico(k) + '</div></div>';
    }
    function pill(cx, cy, t) {
      t = String(t); var w = t.length * 7.6 + 18;
      if (DESIGN !== 'classic' && DESIGN !== 'aurora') return T(cx, cy + 4, t, { s: 12, w: 800 });
      return '<rect x="' + (cx - w / 2) + '" y="' + (cy - 11) + '" width="' + w + '" height="21" rx="10.5" fill="rgba(148,163,184,.17)" stroke="rgba(148,163,184,.4)"/>' + T(cx, cy + 4, t, { s: 12, w: 800 });
    }
    function bar3v(x, y, w, h, d, gid, side, topc) {
      var b = y + h;
      return '<ellipse cx="' + (x + w / 2 + d / 2) + '" cy="' + (b + 2) + '" rx="' + (w * .64) + '" ry="4" fill="rgba(0,0,0,.28)"/>' +
        '<path d="M' + (x + w) + ',' + y + ' L' + (x + w + d) + ',' + (y - d) + ' L' + (x + w + d) + ',' + (b - d) + ' L' + (x + w) + ',' + b + 'Z" fill="' + side + '"/>' +
        '<path d="M' + x + ',' + y + ' L' + (x + d) + ',' + (y - d) + ' L' + (x + w + d) + ',' + (y - d) + ' L' + (x + w) + ',' + y + 'Z" fill="' + topc + '"/>' +
        '<rect x="' + x + '" y="' + y + '" width="' + w + '" height="' + h + '" fill="url(#' + gid + ')"/>' +
        '<rect x="' + (x + w * .1) + '" y="' + (y + 3) + '" width="' + (w * .17) + '" height="' + Math.max(h - 6, 0) + '" rx="' + (w * .085) + '" fill="rgba(255,255,255,.22)"/>';
    }
    function bar3h(x, y, w, h, d, gid, side, topc) {
      return '<path d="M' + x + ',' + y + ' L' + (x + d) + ',' + (y - d) + ' L' + (x + w + d) + ',' + (y - d) + ' L' + (x + w) + ',' + y + 'Z" fill="' + topc + '"/>' +
        '<path d="M' + (x + w) + ',' + y + ' L' + (x + w + d) + ',' + (y - d) + ' L' + (x + w + d) + ',' + (y + h - d) + ' L' + (x + w) + ',' + (y + h) + 'Z" fill="' + side + '"/>' +
        '<rect x="' + x + '" y="' + y + '" width="' + w + '" height="' + h + '" fill="url(#' + gid + ')"/>' +
        '<rect x="' + x + '" y="' + (y + 1.5) + '" width="' + w + '" height="' + (h * .38) + '" fill="rgba(255,255,255,.22)"/>';
    }
    /* bar shapes per design: 3d (classic), flat, pill (rounded), thin, lollipop */
    function barV(x, y, w, h, d, gid, side, topc, c) {
      var b = y + h;
      if (TH.bar === '3d') return bar3v(x, y, w, h, d, gid, side, topc);
      if (TH.bar === 'flat') return '<rect x="' + x + '" y="' + y + '" width="' + w + '" height="' + h + '" fill="' + c + '"/>';
      if (TH.bar === 'pill') return '<rect x="' + x + '" y="' + y + '" width="' + w + '" height="' + h + '" rx="' + Math.min(w / 2, h / 2) + '" fill="url(#' + gid + ')"/>';
      if (TH.bar === 'thin') { var tw = Math.max(6, w * .42); return '<rect x="' + (x + (w - tw) / 2) + '" y="' + y + '" width="' + tw + '" height="' + h + '" fill="' + c + '"/>'; }
      var cx = x + w / 2, r = Math.max(6, Math.min(13, w * .26));
      return '<line x1="' + cx + '" y1="' + b + '" x2="' + cx + '" y2="' + (y + r) + '" stroke="' + c + '" stroke-width="3"/><circle cx="' + cx + '" cy="' + (y + r) + '" r="' + r + '" fill="' + c + '" stroke="#fff" stroke-width="2"/>';
    }
    function barH(x, y, w, h, d, gid, side, topc, c) {
      if (TH.bar === '3d') return bar3h(x, y, w, h, d, gid, side, topc);
      if (TH.bar === 'flat') return '<rect x="' + x + '" y="' + y + '" width="' + w + '" height="' + h + '" fill="' + c + '"/>';
      if (TH.bar === 'pill') return '<rect x="' + x + '" y="' + y + '" width="' + w + '" height="' + h + '" rx="' + (h / 2) + '" fill="url(#' + gid + ')"/>';
      if (TH.bar === 'thin') { var th = Math.max(5, h * .4); return '<rect x="' + x + '" y="' + (y + (h - th) / 2) + '" width="' + w + '" height="' + th + '" fill="' + c + '"/>'; }
      var cy = y + h / 2, r = Math.max(6, Math.min(11, h * .42));
      return '<line x1="' + x + '" y1="' + cy + '" x2="' + (x + w - r) + '" y2="' + cy + '" stroke="' + c + '" stroke-width="3"/><circle cx="' + (x + w - r) + '" cy="' + cy + '" r="' + r + '" fill="' + c + '" stroke="#fff" stroke-width="2"/>';
    }
    function wrap(t, n) {
      t = String(t); if (t.length <= n) return [t];
      var a = '', b = ''; t.split(' ').forEach(function (x) { if (!b && (a + ' ' + x).trim().length <= n) a = (a + ' ' + x).trim(); else b = (b + ' ' + x).trim(); });
      if (b.length > n) b = b.slice(0, n - 1) + '\u2026';
      return [a || t.slice(0, n), b].filter(Boolean);
    }
    function lines(x, y, t, n, o) { var ls = wrap(t, n), s = ''; ls.forEach(function (l, k) { s += T(x, y + (ls.length > 1 ? (k ? 7 : -6) : 4), l, o); }); return s; }
    function legend(items) { return '<div class="pdf-leg">' + items.map(function (s) { return '<span><i style="background:' + s.color + '"></i>' + esc(s.name) + '</span>'; }).join('') + '</div>'; }

    function vbar(labels, vals, o) {
      o = o || {}; var W = o.w || 560, H = o.h || 210, n = labels.length || 1, pl = 14, pr = 24, pt = 40, pb = 44, id = 'g' + (++uid), ph = H - pt - pb;
      var max = Math.max.apply(null, vals.concat([1])), slot = (W - pl - pr) / n, bw = Math.min(o.bw || slot * .5, 76), d = S3D ? Math.min(11, bw * .22) : 0, inner = '', defs = '';
      [0, .25, .5, .75, 1].forEach(function (g) { var y = H - pb - ph * g; inner += '<line x1="' + pl + '" x2="' + (W - pr) + '" y1="' + y + '" y2="' + y + '" stroke="rgba(148,163,184,' + (g ? .18 : .55) + ')"' + (g ? ' stroke-dasharray="3 5"' : '') + '/>'; });
      labels.forEach(function (l, i) {
        var h = Math.max(vals[i] / max * ph, 3), x = pl + slot * i + (slot - bw) / 2 - d / 2, y = H - pb - h, c = o.colors ? o.colors[i] : (o.color || '#4472c4'), gid = id + 'b' + i;
        if (TH.hl && !o.colors) c = vals[i] === max ? TH.hl : c;
        var lo = o.grad ? o.grad[0] : shade(c, -.22), hi = o.grad ? o.grad[1] : shade(c, .32), side = shade(o.grad ? o.grad[0] : c, -.42), topc = shade(o.grad ? o.grad[1] : c, .5);
        defs += grad(gid, lo, hi, true);
        inner += barV(x, y, bw, h, d, gid, side, topc, o.grad && !TH.hl ? o.grad[0] : c) + pill(x + bw / 2 + d / 2, y - d - 14, o.fmt ? o.fmt(vals[i], i) : vals[i]);
        wrap(l, slot < 120 ? 13 : 24).forEach(function (t, k) { inner += T(pl + slot * i + slot / 2, H - pb + 19 + k * 13, t, { s: 11 }); });
      });
      return svg(W, H, '<defs>' + defs + '</defs>' + inner);
    }
    function hbar(labels, vals, o) {
      o = o || {}; var n = labels.length || 1, rowH = o.rowH || 40, lw = o.lw || 150, W = o.w || 620, pr = o.pr || 56, aw = W - lw - pr, H = n * rowH + (o.axis ? 28 : 8);
      var max = o.max || Math.max.apply(null, vals.concat([1])), id = 'g' + (++uid), defs = '', inner = '';
      labels.forEach(function (l, i) {
        var yc = i * rowH + rowH / 2 + 3, bh = Math.min(o.bh || 26, rowH - 10), d = S3D ? Math.min(7, bh * .26) : 0, w = Math.max(6, vals[i] / max * aw), y = yc - bh / 2 + d / 2, gid = id + 'h' + i;
        var c = o.colors ? o.colors[i] : (o.color || '#4472c4');
        if (TH.hl && !o.colors) c = vals[i] === max ? TH.hl : c;
        var lo = o.grad ? o.grad[0] : shade(c, -.2), hi = o.grad ? o.grad[1] : shade(c, .34), side = shade(o.grad ? o.grad[0] : c, -.42), topc = shade(o.grad ? o.grad[1] : c, .5);
        defs += grad(gid, lo, hi);
        var txt = o.fmt ? o.fmt(vals[i], i) : vals[i];
        inner += lines(lw - 12, yc, l, o.wrap || 26, { a: 'end', s: 11 }) +
          (TH.bar === 'thin' || TH.bar === 'lollipop' ? '' : '<rect x="' + lw + '" y="' + (y - d) + '" width="' + (aw + d) + '" height="' + (bh + d) + '" rx="' + (TH.bar === 'pill' ? (bh + d) / 2 : TH.bar === 'flat' ? 0 : 4) + '" fill="rgba(148,163,184,.12)"/>') +
          barH(lw, y, w, bh, d, gid, side, topc, o.grad && !TH.hl ? o.grad[0] : c) +
          (o.inside && TH.bar !== 'thin' && TH.bar !== 'lollipop' && w > String(txt).length * 8 + 18 ? T(lw + w / 2, yc + 5, txt, { c: '#fff', s: 13, w: 800 }) : T(lw + w + d + 9, yc + 5, txt, { a: 'start', s: 13, w: 800 }));
      });
      if (o.axis) { inner += '<line x1="' + lw + '" x2="' + lw + '" y1="0" y2="' + (n * rowH) + '" stroke="rgba(160,174,192,.6)"/>'; [0, 20, 40, 60, 80, 100].forEach(function (t) { inner += T(lw + aw * t / 100, H - 6, t + '%', { s: 10, w: 600 }); }); }
      return svg(W, H, '<defs>' + defs + '</defs>' + inner);
    }
    function clustered(labels, series, o) {
      o = o || {}; var W = o.w || 560, H = o.h || 230, n = labels.length || 1, k = series.length || 1, pl = 12, pt = 36, pb = 44, slot = (W - 2 * pl) / n, gw = slot * .84, bw = Math.min(gw / k - 3, 34), d = S3D ? Math.min(6, bw * .22) : 0, inner = '', defs = '', id = 'g' + (++uid), ph = H - pt - pb;
      var max = Math.max.apply(null, [1].concat.apply([], series.map(function (s) { return s.data || []; })));
      [0, .25, .5, .75, 1].forEach(function (g) { var y = H - pb - ph * g; inner += '<line x1="' + pl + '" x2="' + (W - pl) + '" y1="' + y + '" y2="' + y + '" stroke="rgba(148,163,184,' + (g ? .18 : .55) + ')"' + (g ? ' stroke-dasharray="3 5"' : '') + '/>'; });
      series.forEach(function (s, j) { defs += grad(id + 's' + j, shade(s.color, -.22), shade(s.color, .32), true); });
      labels.forEach(function (l, i) {
        var x0 = pl + slot * i + (slot - (bw + 3) * k) / 2;
        series.forEach(function (s, j) {
          var v = (s.data || [])[i] || 0, h = v ? Math.max(v / max * ph, 3) : 0, x = x0 + j * (bw + 3);
          if (v) inner += barV(x, H - pb - h, bw, h, d, id + 's' + j, shade(s.color, -.42), shade(s.color, .5), s.color) + T(x + bw / 2 + d / 2, H - pb - h - d - 6, v, { s: 11, w: 800 });
        });
        wrap(l, slot < 110 ? 12 : 22).forEach(function (t, q) { inner += T(pl + slot * i + slot / 2, H - pb + 19 + q * 13, t, { s: 11 }); });
      });
      return svg(W, H, '<defs>' + defs + '</defs>' + inner) + legend(series);
    }
    function stacked(rows, series, o) {
      o = o || {}; var W = o.w || 680, rowH = 48, lw = 210, pr = 24, aw = W - lw - pr, H = rows.length * rowH + 8, inner = '', defs = '', id = 'g' + (++uid);
      var tots = rows.map(function (_, i) { return sum(series.map(function (s) { return s.data[i] || 0; })); }), max = Math.max.apply(null, tots.concat([1]));
      series.forEach(function (s, j) { defs += grad(id + 's' + j, shade(s.color, -.28), shade(s.color, .3), true); });
      rows.forEach(function (r, i) {
        var yc = i * rowH + rowH / 2 + 3, y0 = yc - 15, x = lw, tw = tots[i] / max * aw, cp = id + 'c' + i, seg = '', txt = '';
        inner += lines(lw - 12, yc, r, 34, { a: 'end', s: 11 }) + '<rect x="' + lw + '" y="' + (y0 - 1) + '" width="' + aw + '" height="32" rx="9" fill="rgba(148,163,184,.12)"/>';
        defs += '<clipPath id="' + cp + '"><rect x="' + lw + '" y="' + y0 + '" width="' + Math.max(tw, 1) + '" height="30" rx="8"/></clipPath>';
        series.forEach(function (s, j) {
          var v = s.data[i] || 0; if (!v) return; var w = v / max * aw;
          seg += '<rect x="' + x + '" y="' + y0 + '" width="' + w + '" height="30" fill="url(#' + id + 's' + j + ')"/><rect x="' + (x + w - 1.5) + '" y="' + y0 + '" width="1.5" height="30" fill="rgba(0,0,0,.28)"/>';
          if (w > 16) txt += T(x + w / 2, yc + 5, v, { c: '#fff', s: 12, w: 800 });
          x += w;
        });
        inner += '<g clip-path="url(#' + cp + ')">' + seg + '<rect x="' + lw + '" y="' + y0 + '" width="' + Math.max(tw, 1) + '" height="13" fill="rgba(255,255,255,.2)"/></g>' + txt;
      });
      return svg(W, H, '<defs>' + defs + '</defs>' + inner) + legend(series);
    }
    function donut(p, color, text, cap) {
      var r = 64, sw = 30, cx = 130, cy = 104, D = 12, c = 2 * Math.PI * r, d = Math.max(0, Math.min(100, p)) / 100 * c, id = 'd' + (++uid), depth = '';
      var ring = function (y, col, dash) { return '<circle cx="' + cx + '" cy="' + (cy + y) + '" r="' + r + '" fill="none" stroke="' + col + '" stroke-width="' + sw + '"' + (dash ? ' stroke-dasharray="' + d + ' ' + (c - d) + '" transform="rotate(-90 ' + cx + ' ' + (cy + y) + ')"' : '') + '/>'; };
      if (TH.donut !== '3d') {
        var sw2 = TH.donut === 'round' ? 20 : 26, lcap = TH.donut === 'round' ? ' stroke-linecap="round"' : '';
        var ring2 = function (col, dash, op) { return '<circle cx="' + cx + '" cy="' + cy + '" r="' + r + '" fill="none" stroke="' + col + '" stroke-opacity="' + (op || 1) + '" stroke-width="' + sw2 + '"' + lcap + (dash ? ' stroke-dasharray="' + d + ' ' + (c - d) + '" transform="rotate(-90 ' + cx + ' ' + cy + ')"' : '') + '/>'; };
        return svg(260, 250, '<defs><linearGradient id="' + id + 'v" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="' + (TH.donut === 'round' ? TH.prim[0] : color) + '"/><stop offset="1" stop-color="' + (TH.donut === 'round' ? TH.prim[1] : color) + '"/></linearGradient></defs>' +
          ring2('rgba(148,163,184,.25)', false) + (d > 0 ? ring2(TH.donut === 'round' ? 'url(#' + id + 'v)' : color, true) : '') +
          T(cx, cy + 6, text, { s: 24, w: 800 }) + (cap ? T(cx, cy + 24, cap, { s: 9, w: 700 }) : ''));
      }
      for (var k = D; k >= 1; k--) depth += ring(k, '#2b3647', false) + (d > 0 ? ring(k, shade(color, -.5), true) : '');
      return svg(260, 250,
        '<defs><linearGradient id="' + id + 'v" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="' + shade(color, .45) + '"/><stop offset="1" stop-color="' + color + '"/></linearGradient>' +
        '<linearGradient id="' + id + 't" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="#7b889b"/><stop offset="1" stop-color="#566377"/></linearGradient>' +
        '<radialGradient id="' + id + 's" cx=".5" cy=".5" r=".5"><stop offset="0" stop-color="#000" stop-opacity=".5"/><stop offset="1" stop-color="#000" stop-opacity="0"/></radialGradient></defs>' +
        '<ellipse cx="' + cx + '" cy="' + (cy + r + sw / 2 + D + 6) + '" rx="' + (r + 34) + '" ry="12" fill="url(#' + id + 's)"/>' + depth +
        ring(0, 'url(#' + id + 't)', false) + (d > 0 ? ring(0, 'url(#' + id + 'v)', true) : '') +
        '<circle cx="' + cx + '" cy="' + cy + '" r="' + (r + sw / 2 - 2) + '" fill="none" stroke="rgba(255,255,255,.3)" stroke-width="1.2"/>' +
        '<circle cx="' + cx + '" cy="' + cy + '" r="' + (r - sw / 2 + 2) + '" fill="none" stroke="rgba(0,0,0,.25)" stroke-width="1.2"/>' +
        T(cx, cy + 6, text, { s: 24, w: 800 }) + (cap ? T(cx, cy + 24, cap, { s: 9, w: 700 }) : ''));
    }
    function pie(labels, vals, colors) {
      /* Real 3D pie: lit top face, extruded side wall, floor shadow, gloss, and leader-line labels that are
         spread apart so neighbouring small slices (e.g. Netherlands / Canada) can never overlap. */
      var FLAT = TH.pie !== '3d';
      var W = 800, H = 372, cx = 400, cy = FLAT ? 180 : 170, r = FLAT ? 140 : 124, ry = FLAT ? 1 : .54, R = r * ry, dp = FLAT ? 0 : 30, tot = sum(vals) || 1, id = 'p' + (++uid);
      var ri = TH.pie === 'ring' ? r * .58 : 0;
      var f1 = function (n) { return +n.toFixed(2); };
      var P2 = function (a, dy) { return [cx + r * Math.cos(a), cy + R * Math.sin(a) + (dy || 0)]; };
      var pt = function (p) { return f1(p[0]) + ',' + f1(p[1]); };
      var defs = '<radialGradient id="' + id + 's" cx=".5" cy=".5" r=".5"><stop offset="0" stop-color="#000" stop-opacity=".6"/><stop offset=".65" stop-color="#000" stop-opacity=".22"/><stop offset="1" stop-color="#000" stop-opacity="0"/></radialGradient>' +
        '<linearGradient id="' + id + 'w" gradientUnits="userSpaceOnUse" x1="' + (cx - r) + '" y1="0" x2="' + (cx + r) + '" y2="0"><stop offset="0" stop-color="#000" stop-opacity=".55"/><stop offset=".28" stop-color="#fff" stop-opacity=".14"/><stop offset=".55" stop-color="#000" stop-opacity="0"/><stop offset="1" stop-color="#000" stop-opacity=".6"/></linearGradient>' +
        '<radialGradient id="' + id + 'g" cx=".3" cy=".3" r=".75"><stop offset="0" stop-color="#fff" stop-opacity=".2"/><stop offset=".5" stop-color="#fff" stop-opacity=".04"/><stop offset="1" stop-color="#fff" stop-opacity="0"/></radialGradient>';
      var a0 = -Math.PI / 2, walls = '', tops = '', items = [];
      vals.forEach(function (v, i) {
        var col = colors[i % colors.length], frac = v / tot, a1 = a0 + frac * 2 * Math.PI;
        if (v > 0) {
          defs += '<linearGradient id="' + id + 't' + i + '" gradientUnits="userSpaceOnUse" x1="0" y1="' + (cy - R) + '" x2="0" y2="' + (cy + R) + '"><stop offset="0" stop-color="' + shade(col, .38) + '"/><stop offset="1" stop-color="' + shade(col, -.14) + '"/></linearGradient>';
          var p0 = P2(a0), p1 = P2(a1), big = (a1 - a0) > Math.PI ? 1 : 0, top;
          if (FLAT && ri) {
            var q0 = [cx + ri * Math.cos(a0), cy + ri * Math.sin(a0)], q1 = [cx + ri * Math.cos(a1), cy + ri * Math.sin(a1)];
            if (frac >= .9999) top = 'M' + (cx - r) + ',' + cy + ' A' + r + ',' + r + ' 0 1 1 ' + (cx + r) + ',' + cy + ' A' + r + ',' + r + ' 0 1 1 ' + (cx - r) + ',' + cy + 'Z M' + (cx - ri) + ',' + cy + ' A' + ri + ',' + ri + ' 0 1 0 ' + (cx + ri) + ',' + cy + ' A' + ri + ',' + ri + ' 0 1 0 ' + (cx - ri) + ',' + cy + 'Z';
            else top = 'M' + pt(p0) + ' A' + r + ',' + r + ' 0 ' + big + ' 1 ' + pt(p1) + ' L' + pt(q1) + ' A' + ri + ',' + ri + ' 0 ' + big + ' 0 ' + pt(q0) + 'Z';
          } else if (frac >= .9999) top = 'M' + (cx - r) + ',' + cy + ' A' + r + ',' + f1(R) + ' 0 1 1 ' + (cx + r) + ',' + cy + ' A' + r + ',' + f1(R) + ' 0 1 1 ' + (cx - r) + ',' + cy + 'Z';
          else top = 'M' + cx + ',' + cy + ' L' + pt(p0) + ' A' + r + ',' + f1(R) + ' 0 ' + big + ' 1 ' + pt(p1) + 'Z';
          var s = Math.max(a0, 0), e = Math.min(a1, Math.PI);
          if (e > s && !FLAT) {
            var wd = 'M' + pt(P2(s)) + ' A' + r + ',' + f1(R) + ' 0 0 1 ' + pt(P2(e)) + ' L' + pt(P2(e, dp)) + ' A' + r + ',' + f1(R) + ' 0 0 0 ' + pt(P2(s, dp)) + 'Z';
            walls += '<path d="' + wd + '" fill="' + shade(col, -.45) + '"/><path d="' + wd + '" fill="url(#' + id + 'w)"/>';
          }
          tops += '<path d="' + top + '" fill-rule="evenodd" fill="' + (FLAT ? col : 'url(#' + id + 't' + i + ')') + '" stroke="rgba(255,255,255,.8)" stroke-width="' + (FLAT ? 2 : 1.6) + '" stroke-linejoin="round"/>';
          items.push({ i: i, m: (a0 + a1) / 2, col: col, v: v });
        }
        a0 = a1;
      });
      var sides = { r: [], l: [] };
      items.forEach(function (it) {
        var c = Math.cos(it.m), sn = Math.sin(it.m);
        it.ax = cx + r * c; it.ay = cy + R * sn + (sn > 0 ? dp * .55 : 0);
        it.ex = cx + (r + 22) * c; it.ey = it.ay + sn * 16; it.ly = it.ey; it.dir = c >= 0 ? 1 : -1;
        sides[c >= 0 ? 'r' : 'l'].push(it);
      });
      var gap = 42, lo = 28, hi = H - 28;
      ['r', 'l'].forEach(function (k) {
        var a = sides[k]; a.sort(function (p, q) { return p.ly - q.ly; });
        /* relax: push neighbours apart symmetrically so label stacks stay centred on their slices */
        for (var it2 = 0; it2 < 60; it2++) {
          var moved = false;
          for (var n = 1; n < a.length; n++) {
            var d = a[n].ly - a[n - 1].ly;
            if (d < gap - .5) { var h = (gap - d) / 2; a[n - 1].ly -= h; a[n].ly += h; moved = true; }
          }
          a.forEach(function (x) { x.ly = Math.min(hi, Math.max(lo, x.ly)); });
          if (!moved) break;
        }
      });
      var lab = '';
      items.forEach(function (it) {
        var lx = cx + it.dir * (r + 96), an = it.dir > 0 ? 'start' : 'end', tx = lx + it.dir * 13, nm = String(labels[it.i]);
        if (nm.length > 20) nm = nm.slice(0, 19) + '\u2026';
        lab += '<polyline points="' + f1(it.ax) + ',' + f1(it.ay) + ' ' + f1(it.ex) + ',' + f1(it.ey) + ' ' + f1(lx - it.dir * 8) + ',' + f1(it.ly + 3) + '" fill="none" stroke="' + it.col + '" stroke-width="1.7" stroke-linejoin="round"/>' +
          '<circle cx="' + f1(it.ax) + '" cy="' + f1(it.ay) + '" r="3.6" fill="#fff" stroke="' + it.col + '" stroke-width="2"/>' +
          '<circle cx="' + f1(lx) + '" cy="' + f1(it.ly + 3) + '" r="5.5" fill="' + it.col + '" stroke="rgba(255,255,255,.85)" stroke-width="1.4"/>' +
          T(tx, it.ly - 1, nm, { a: an, s: 14, w: 800 }) +
          '<text x="' + f1(tx) + '" y="' + f1(it.ly + 16) + '" text-anchor="' + an + '" font-size="13" fill="currentColor"><tspan font-weight="800" fill="' + it.col + '">' + pc(it.v, tot, 1) + '%</tspan><tspan font-weight="600" fill-opacity=".7"> \u00b7 ' + Number(it.v).toLocaleString() + '</tspan></text>';
      });
      if (FLAT) return svg(W, H, '<defs>' + defs + '</defs>' + tops + lab +
        (ri ? T(cx, cy + 2, Number(tot).toLocaleString(), { s: 30, w: 800 }) + T(cx, cy + 22, 'LEADS', { s: 11, w: 700 }) : ''));
      return svg(W, H,
        '<defs>' + defs + '</defs>' +
        '<ellipse cx="' + cx + '" cy="' + (cy + dp + 8) + '" rx="' + f1(r * 1.1) + '" ry="' + f1(R * 1.3) + '" fill="url(#' + id + 's)"/>' +
        walls + tops +
        '<ellipse cx="' + cx + '" cy="' + cy + '" rx="' + r + '" ry="' + f1(R) + '" fill="url(#' + id + 'g)"/>' +
        '<ellipse cx="' + cx + '" cy="' + cy + '" rx="' + r + '" ry="' + f1(R) + '" fill="none" stroke="rgba(255,255,255,.35)" stroke-width="1"/>' + lab);
    }

    /* ---------- page helpers ---------- */
    var panel = function (h, body, st) { return '<div class="panel pdf-panel"' + (st ? ' style="' + st + '"' : '') + '>' + (h ? '<h4>' + h + '</h4>' : '') + body + '</div>'; };
    var kpi = function (t, ic, v) { return '<div class="panel pdf-kpi"><h4>' + t + '</h4><div class="ic">' + badge(ic, KC[ic], 88) + '</div><div class="num">' + v + '</div></div>'; };
    var grid = function (cols, inner) { return '<div class="pdf-grid" style="grid-template-columns:' + cols + '">' + inner + '</div>'; };
    function slide(id, title, body) {
      var s = $(id); if (!s) return;
      var n = (s.querySelector('.slideNo') || {}).textContent || '';
      s.innerHTML = '<h2 class="title">' + title + '</h2>' + body + '<div class="footer">VALASYS MEDIA\u2122</div><div class="slideNo">' + n + '</div>';
    }
    function hrows(rows, grads) {
      return rows.map(function (r, i) { var g = grads[i % grads.length];
        return '<div class="pdf-hbar"><span>' + esc(r[0]) + '</span><div class="pdf-track"><div class="pdf-fill" style="width:' + Math.max(2, Math.min(100, r[1])) + '%;background:linear-gradient(90deg,' + g[0] + ',' + g[1] + ')"></div></div><span>' + r[2] + '</span></div>'; }).join('');
    }
    var sizeKey = function (l) { var m = String(l).replace(/,/g, '').match(/\d+/); return m ? +m[0] : 1e9; };
    var order = function (names, fn) { return names.map(function (_, i) { return i; }).sort(function (a, b) { return fn(names[a], names[b], a, b); }); };
    var acol = function (name) { var i = P.asset_names.indexOf(name); return COL[(i < 0 ? 0 : i) % COL.length]; };
    var leads = P.lead_count || sum(P.job_values || []) || 1;

    /* ---- Slide 3 ---- */
    var ji = order(P.job_labels, function (a, b, i, j) { return P.job_values[i] - P.job_values[j]; });
    var jl = ji.map(function (i) { return P.job_labels[i]; }), jv = ji.map(function (i) { return P.job_values[i]; }), jt = sum(jv) || 1;
    var split = jl.map(function (l, i) { return [l, jv[i] / jt * 100, pc(jv[i], jt) + '%']; }).reverse();
    var dm = 0;
    if (P.decision_labels && P.decision_labels.length) { P.decision_labels.forEach(function (l, i) { if (/decision/i.test(l) && !/non|not/i.test(l)) dm += P.decision_values[i]; }); jt = sum(P.decision_values) || jt; }
    else jl.forEach(function (l, i) { if (/c-?level|chief|\bc[a-z]o\b|vice|\bvp\b|president|director|head|owner|founder|decision/i.test(l)) dm += jv[i]; });
    var dmP = jt ? dm / jt * 100 : 0, ft = sum(P.func_values) || 1;
    slide('s3', 'Campaign Dashboard',
      grid('1fr 2fr 2fr', kpi('Leads Generated', 'target', leads) + panel('Job Level', vbar(jl, jv, { w: 520, h: 200, bw: 46, grad: TH.prim })) + panel('Job Level Split', hrows(split, TH.split))) +
      grid('2fr 1.3fr 1.3fr', panel('Job Functions', hbar(P.func_labels, P.func_values.map(function (v) { return v / ft * 100; }), { w: 520, rowH: 70, bh: 56, lw: 130, pr: 20, max: 100, axis: true, inside: true, wrap: 16, color: TH.one, fmt: function (v) { return v.toFixed(2) + '%'; } })) + panel('Decision Makers', donut(dmP, TH.donutC[0], dmP.toFixed(1) + '%', 'DECISION MAKERS')) + panel('Recommender', donut(100 - dmP, TH.donutC[1], (100 - dmP).toFixed(1) + '%', 'INFLUENCERS'))));

    /* ---- Slide 4 ---- */
    var ind = P.industry_labels.slice(0, 5), indv = P.industry_values.slice(0, 5);
    var si = order(P.size_labels, function (a, b) { return sizeKey(b) - sizeKey(a); });
    var sl = si.map(function (i) { return P.size_labels[i]; }), sv = si.map(function (i) { return P.size_values[i]; }), stt = sum(sv) || 1;
    var dt = sum(P.device_values) || 1;
    var DEV_COL = TH.dev;
    function devKind(l, i) {
      l = String(l);
      if (/desk|laptop|\bpc\b|computer|windows|\bmac\b/i.test(l)) return 'desktop';
      if (/mobile|phone|android|\bios\b|iphone/i.test(l)) return 'phone';
      if (/tablet|ipad/i.test(l)) return 'tablet';
      if (/influ|recommend|non[\s-]*decision|research|evaluat/i.test(l)) return 'group';
      if (/decision|maker|executive|c[\s-]*level|budget/i.test(l)) return 'decision';
      return ['decision', 'group', 'desktop', 'phone'][i % 4];
    }
    function devIcon(kind) {
      var W = '#fff', D = 'rgba(8,16,30,.34)', b = '';
      if (kind === 'desktop') b = '<rect x="7" y="9" width="50" height="33" rx="4.5" fill="' + W + '"/><rect x="11.5" y="13.5" width="41" height="24" rx="2" fill="' + D + '"/><path d="M14 34L23 26L29 31L38 21L45 28L50 27V35H14Z" fill="#fff" fill-opacity=".34"/><path d="M24 43h16l2.4 8H21.6z" fill="' + W + '" fill-opacity=".92"/><rect x="16" y="50" width="32" height="4.5" rx="2.25" fill="' + W + '"/>';
      else if (kind === 'phone') b = '<rect x="18" y="4" width="28" height="56" rx="6.5" fill="' + W + '"/><rect x="22.5" y="12" width="19" height="35" rx="2" fill="' + D + '"/><path d="M24 42L29 35L33 39L38 31V44H24Z" fill="#fff" fill-opacity=".34"/><rect x="28" y="7" width="8" height="2" rx="1" fill="' + D + '"/><circle cx="32" cy="53.5" r="2.4" fill="' + D + '"/>';
      else if (kind === 'tablet') b = '<rect x="8" y="7" width="48" height="50" rx="6.5" fill="' + W + '"/><rect x="13" y="12" width="38" height="36" rx="2" fill="' + D + '"/><path d="M15 43L24 33L30 39L38 27L49 41V46H15Z" fill="#fff" fill-opacity=".34"/><circle cx="32" cy="52.5" r="2.2" fill="' + D + '"/>';
      else if (kind === 'group') b = '<circle cx="13" cy="25" r="6.5" fill="' + W + '" fill-opacity=".72"/><path d="M1.5 50c0-8.5 5-13.5 11.5-13.5 2.3 0 4.3.6 6 1.6C16.6 42 15.5 45.8 15.5 50z" fill="' + W + '" fill-opacity=".72"/><circle cx="51" cy="25" r="6.5" fill="' + W + '" fill-opacity=".72"/><path d="M62.5 50c0-8.5-5-13.5-11.5-13.5-2.3 0-4.3.6-6 1.6C47.4 42 48.5 45.8 48.5 50z" fill="' + W + '" fill-opacity=".72"/><circle cx="32" cy="20" r="9.5" fill="' + W + '"/><path d="M14.5 56c0-11.5 7-19 17.5-19s17.5 7.5 17.5 19z" fill="' + W + '"/>';
      else if (kind === 'decision') b = '<circle cx="30" cy="19" r="10.5" fill="' + W + '"/><path d="M9 57c0-12.5 8.5-20.5 21-20.5S51 44.5 51 57z" fill="' + W + '"/><path d="M30 38l-4.2 5 3 3-2.4 9.5h7.2L31 46l3-3z" fill="' + D + '"/><path d="M52 4.5l2.7 5.5 6 .8-4.4 4.2 1.1 6-5.4-2.9-5.4 2.9 1.1-6-4.4-4.2 6-.8z" fill="#ffd54a"/>';
      else b = '<circle cx="32" cy="32" r="24" fill="' + W + '"/><path d="M32 32V8a24 24 0 0 1 22.8 16.6z" fill="' + D + '"/>';
      return '<svg viewBox="0 0 64 64" xmlns="http://www.w3.org/2000/svg">' + b + '</svg>';
    }
    var devHtml;
    if (!P.device_values || !P.device_values.length || !sum(P.device_values)) devHtml = '<div class="pdf-dev"><div class="tiles"><div class="tile"><div class="nm">No data</div></div></div></div>';
    else {
      var shown = P.device_labels.slice(0, 4), tiles = '', bar = '', used = 0;
      shown.forEach(function (l, i) {
        var c = DEV_COL[i % DEV_COL.length], v = P.device_values[i] || 0, p = pc(v, dt, 1); used += v;
        tiles += '<div class="tile"><div class="badge" style="background:radial-gradient(circle at 30% 22%,' + shade(c, .55) + ',' + c + ' 52%,' + shade(c, -.4) + ');box-shadow:0 0 0 6px ' + rgba(c, .14) + ',0 18px 26px -12px ' + c + ',inset 0 -10px 18px rgba(0,0,0,.32),inset 0 8px 14px rgba(255,255,255,.42)">' + devIcon(devKind(l, i)) + '</div>' +
          '<div class="big" style="color:' + c + ';text-shadow:0 0 16px ' + rgba(c, .55) + '">' + p + '%</div><div class="nm">' + esc(l) + '</div><div class="ct">' + (P.device_is_pct ? 'of audience' : Number(v).toLocaleString() + ' leads') + '</div></div>';
        bar += '<span style="width:' + p + '%;background:linear-gradient(180deg,' + shade(c, .42) + ' 0%,' + c + ' 48%,' + shade(c, -.32) + ' 100%)">' + (+p >= 11 ? p + '%' : '') + '</span>';
      });
      var restV = dt - used;
      if (restV > 0.0001) bar += '<span style="width:' + pc(restV, dt, 1) + '%;background:linear-gradient(180deg,#cbd5e1,#94a3b8 50%,#64748b)"></span>';
      devHtml = '<div class="pdf-dev"><div class="tiles">' + tiles + '</div><div class="bar">' + bar + '</div></div>';
    }
    slide('s4', 'Campaign Dashboard',
      grid('1fr 3fr', kpi('Unique Industries', 'factory', P.unique_industry_count || P.industry_labels.length) + panel('Top 5 Industries', vbar(ind.map(function (l) { return l.toUpperCase(); }), indv, { w: 760, h: 210, bw: 34, color: TH.one, fmt: function (v) { return pc(v, leads) + '%'; } }))) +
      grid('1.4fr 1fr', panel('Employee Size', hbar(sl, sv, { w: 560, rowH: 34, bh: 24, lw: 80, pr: 50, grad: TH.sec, fmt: function (v) { return pc(v, stt) + '%'; } })) + panel(P.device_title || 'Devices', devHtml)));

    /* ---- Slide 5: KPI + location map on top, Location Split pie + Geographic Reach below ---- */
    var s5 = $('s5');
    if (s5) {
      var t5 = s5.querySelector('.title'); if (t5) t5.textContent = 'Campaign Dashboard';
      var cl = P.country_labels.slice(0, 7), cvv = P.country_values.slice(0, 7), rest = (P.geo_total || sum(P.country_values)) - sum(cvv);
      if (rest > 0) { cl.push('Others'); cvv.push(rest); }
      window.PRA_PIE_COLORS = TH.pieC;
      if (!$('praS5Top')) {
        var gm = $('geoMap'), mapPanel = gm && gm.closest('.panel'), oldGrid = mapPanel && mapPanel.parentElement;
        var reach = oldGrid ? [].filter.call(oldGrid.children, function (c) { return c !== mapPanel; })[0] : null;
        var topRow = document.createElement('div'); topRow.id = 'praS5Top'; topRow.className = 'pdf-grid'; topRow.style.gridTemplateColumns = '1fr 3fr';
        var botRow = document.createElement('div'); botRow.id = 'praS5Bot'; botRow.className = 'pdf-grid'; botRow.style.gridTemplateColumns = '1.7fr 1fr';
        var kp = document.createElement('div'); kp.id = 'praS5Kpi'; topRow.appendChild(kp);
        if (mapPanel) topRow.appendChild(mapPanel);
        var pp = document.createElement('div'); pp.id = 'praS5Pie'; pp.className = 'panel pdf-panel'; botRow.appendChild(pp);
        if (reach) { botRow.appendChild(reach); var rm = reach.querySelector('.metric'); if (rm) rm.style.display = 'none'; }
        if (oldGrid && oldGrid.parentElement === s5) oldGrid.replaceWith(topRow); else if (t5) t5.after(topRow); else s5.prepend(topRow);
        topRow.after(botRow);
      }
      var kTmp = document.createElement('div');
      kTmp.innerHTML = kpi('Unique Geo Locations', 'pin', P.unique_country_count || P.country_labels.length);
      kTmp.firstChild.id = 'praS5Kpi'; $('praS5Kpi').replaceWith(kTmp.firstChild);
      $('praS5Pie').innerHTML = '<h4>Location Split</h4>' + pie(cl, cvv, TH.pieC);
      try { if (typeof praRenderGeoMap === 'function') praRenderGeoMap(); } catch (e) {}
    }

    /* ---- Slide 6 ---- */
    var ai = order(P.asset_labels, function (a, b, i, j) { return P.asset_values[i] - P.asset_values[j]; }), at = sum(P.asset_values) || 1;
    var seriesBy = function (mat) { return P.asset_names.map(function (a, i) { return { name: a, color: COL[i % COL.length], data: (mat && mat[i]) || [] }; }); };
    slide('s6', 'Asset Dashboard',
      grid('1fr 3fr', kpi('Unique Assets', 'doc', P.unique_asset_count) + panel('Asset Split', hbar(ai.map(function (i) { return P.asset_labels[i]; }).reverse(), ai.map(function (i) { return P.asset_values[i] / at * 100; }).reverse(), { w: 700, rowH: 44, lw: 260, wrap: 38, bh: 24, pr: 56, max: 100, colors: ai.map(function (i) { return acol(P.asset_labels[i]); }).reverse(), fmt: function (v) { return v.toFixed(1) + '%'; } }))) +
      grid('1fr 1fr', panel('Asset Engagement by Top Industries', clustered(P.industry_names.slice(0, 5), seriesBy(P.asset_industry_matrix))) + panel('Asset Engagement by Country', clustered(P.country_names || [], seriesBy(P.asset_country_matrix)))));

    /* ---- Slide 7 ---- */
    var szI = order(P.size_names, function (a, b) { return sizeKey(a) - sizeKey(b); });
    var sizeSeries = szI.map(function (j, k) { return { name: P.size_names[j], color: COL[k % COL.length], data: P.asset_names.map(function (_, a) { return (P.asset_size_matrix && P.asset_size_matrix[a]) ? P.asset_size_matrix[a][j] : 0; }) }; });
    var jobSeries = P.job_names.map(function (n, j) { return { name: n, color: COL[j % COL.length], data: P.asset_names.map(function (_, a) { return (P.asset_job_matrix && P.asset_job_matrix[a]) ? P.asset_job_matrix[a][j] : 0; }) }; });
    slide('s7', 'Asset Dashboard', panel('Asset Engagement by Employee Size', stacked(P.asset_names, sizeSeries)) + '<div style="height:14px"></div>' + panel('Asset Engagement by Job Level', stacked(P.asset_names, jobSeries)));

    /* ---- Slides 8-10 ---- */
    function splitSlide(id, title, word, labels, values) {
      var t = sum(values) || 1;
      slide(id, title, grid('1fr 3fr', kpi('Unique Assets', 'doc', P.unique_asset_count) + panel(word + ' Split', hbar(labels, values, { w: 700, rowH: 40, lw: 270, wrap: 40, bh: 22, pr: 50, outline: true }))) +
        panel(word + ' Percentage', vbar(labels, values, { w: 900, h: 230, bw: 120, colors: labels.map(acol), fmt: function (v) { return pc(v, t) + '%'; } })));
    }
    splitSlide('s8', 'Asset Wise Open Split', 'Open', P.asset_open_labels, P.asset_open_values);
    splitSlide('s9', 'Asset Wise Click Split', 'Click', P.asset_click_labels, P.asset_click_values);
    splitSlide('s10', 'Asset Wise Conversion Split', 'Conversion', P.asset_conv_labels, P.asset_conv_values);

    /* ---- Slide 11 ---- */
    var st = [['plane', P.sent, 'SENT'], ['inbox', P.delivered, 'DELIVERED'], ['mail', P.opens, 'OPENS'], ['cursor', P.clicks, 'CLICKS'], ['check', P.conversion, 'CONVERSION'], ['ban', P.bounced, 'BOUNCED']];
    var rates = [['Bounce', P.bounce_rate], ['Conversion', P.conversion_rate], ['Clicks', P.click_rate], ['Open', P.open_rate], ['Delivered', P.delivery_rate]];
    slide('s11', 'Campaign Statistics',
      '<div class="panel"><div class="pdf-stats">' + st.map(function (x) { return '<div class="pdf-stat"><div class="ic">' + badge(x[0], KC[x[0]], 64) + '</div><b>' + Number(x[1] || 0).toLocaleString() + '</b><small>' + x[2] + '</small></div>'; }).join('') + '</div></div><div style="height:14px"></div>' +
      panel('Statistics Split', rates.map(function (r) {
        var RC = { Bounce: ['ban', '#ef4444', '#fb923c'], Conversion: ['check', '#f59e0b', '#fde047'], Clicks: ['cursor', '#8b5cf6', '#c4b5fd'], Open: ['mail', '#06b6d4', '#67e8f9'], Delivered: ['inbox', '#10b981', '#6ee7b7'] }, k = RC[r[0]] || ['check', '#3b82f6', '#93c5fd'];
        if (DESIGN !== 'classic') { var ci = ['Bounce', 'Conversion', 'Clicks', 'Open', 'Delivered'].indexOf(r[0]), cc = TH.series[(ci < 0 ? 0 : ci) % TH.series.length]; k = [k[0], cc, TH.bar === 'flat' || TH.bar === 'thin' || TH.bar === 'lollipop' ? cc : shade(cc, .4)]; }
        return '<div class="pdf-hbar" style="grid-template-columns:170px 1fr 66px"><span class="lb">' + badge(k[0], k[1], 32) + r[0] + '</span><div class="pdf-track"><div class="pdf-fill" style="width:' + Math.max(3, Math.min(100, r[1])) + '%;background:linear-gradient(90deg,' + k[1] + ',' + k[2] + ');box-shadow:inset 0 2px 0 rgba(255,255,255,.42),inset 0 -6px 9px rgba(0,0,0,.24),0 0 14px ' + rgba(k[1], .55) + '"></div></div><span class="vl">' + Number(r[1]).toFixed(1) + '%</span></div>';
      }).join('')));

    /* ---- Slide 12 + Thank you ---- */
    slide('s12', 'Observations &amp; Recommendations', panel('', '<ul class="pdf-obs">' + (P.observations || []).concat(P.recommendations || []).map(function (o) { return '<li>' + o + '</li>'; }).join('') + '</ul>'));
    var deck = $('deckView');
    if (deck && !$('s14')) {
      var ty = document.createElement('section'); ty.className = 'slide hidden'; ty.id = 's14';
      ty.innerHTML = '<div class="pdf-thanks">Thank You!</div><div id="praOfficeMap"></div><div class="pdf-offices">' +
        '<div><b><span class="office-flag">🇺🇸</span> USA Office</b>111 Town Square Place, Suite 1203, Jersey City, NJ 07310<br><b>Ph.: +1 303-960-0264</b><br>255 S Orange Avenue, Suite 104 #2185, Orlando, FL 32801</div>' +
        '<div><b><span class="office-flag">🇦🇪</span> Dubai Office</b>Unit No: 492, DMCC Business Centre, Level No 1, Jewellery &amp; Gemplex 3, Dubai - United Arab Emirates<br><b>Ph.: +971-544570526</b></div>' +
        '<div><b><span class="office-flag">🇮🇳</span> India Office</b>801, 8th Floor, Cerebrum IT Park, B-3 Building, Kalyani Nagar, Pune - 411014</div></div>' +
        '<div class="footer">VALASYS MEDIA\u2122</div>';
      deck.appendChild(ty);
      try { slides.push(ty); } catch (e) {}
    }
    /* ---- Thank-you slide: office locations on a world map ---- */
    var om = $('praOfficeMap');
    if (om && typeof PRA_LAND !== 'undefined') {
      var ac = TH.badge || '#f45132', x0 = -112, x1 = 98, top = 74, bot = 0, MW = 1000, SC = MW / (x1 - x0), MH = (top - bot) * SC;
      var X = function (lon) { return (lon - x0) * SC; }, Y = function (lat) { return (top - lat) * SC; };
      var land = Object.keys(PRA_LAND).map(function (k) { return '<path d="M' + PRA_LAND[k].map(function (q) { return X(q[0]).toFixed(1) + ',' + Y(q[1]).toFixed(1); }).join('L') + 'Z"/>'; }).join('');
      var OFF = [
        { n: 'New Jersey, USA', lat: 40.72, lon: -74.07, dx: -24, a: 'end', dy: -10 },
        { n: 'Orlando, USA', lat: 28.54, lon: -81.38, dx: -24, a: 'end', dy: -10 },
        { n: 'Dubai, UAE', lat: 25.2, lon: 55.27, dx: -22, a: 'end', dy: 22 },
        { n: 'Pune, India', lat: 18.52, lon: 73.86, dx: 18, a: 'start', dy: -10 }
      ];
      var arc = function (a, b, lift) {
        var ax = X(a.lon), ay = Y(a.lat), bx = X(b.lon), by = Y(b.lat), mx = (ax + bx) / 2, my = Math.min(ay, by) - lift;
        return '<path d="M' + ax.toFixed(1) + ',' + ay.toFixed(1) + ' Q' + mx.toFixed(1) + ',' + my.toFixed(1) + ' ' + bx.toFixed(1) + ',' + by.toFixed(1) + '" fill="none" stroke="' + ac + '" stroke-width="2.2" stroke-dasharray="7 6" stroke-opacity=".75"/>';
      };
      var pins = OFF.map(function (o) {
        var x = X(o.lon), y = Y(o.lat);
        return '<circle cx="' + x.toFixed(1) + '" cy="' + y.toFixed(1) + '" r="17" fill="' + ac + '" fill-opacity=".18"/>' +
          '<path d="M' + x.toFixed(1) + ',' + y.toFixed(1) + ' c-7,-9 -11,-14 -11,-20 a11,11 0 1 1 22,0 c0,6 -4,11 -11,20z" fill="' + ac + '" stroke="#fff" stroke-width="2"/>' +
          '<circle cx="' + x.toFixed(1) + '" cy="' + (y - 20).toFixed(1) + '" r="4" fill="#fff"/>' +
          T(x + o.dx, y + (o.dy || 0) + 5, o.n, { a: o.a, s: 15, w: 800 });
      }).join('');
      om.innerHTML = svg(MW, Math.round(MH),
        '<g fill="currentColor" fill-opacity=".16" stroke="currentColor" stroke-opacity=".25" stroke-width="1" stroke-linejoin="round">' + land + '</g>' +
        arc(OFF[0], OFF[1], 30) + arc(OFF[0], OFF[2], 90) + arc(OFF[2], OFF[3], 40) + pins);
    }
    try { update(); } catch (e) {}
  }
  window.praRebuildLayout = function () { try { build(); } catch (e) { console.error('PDF layout failed', e); } };
  function start() { setTimeout(window.praRebuildLayout, 150); }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', start); else start();
})();
"""

def pdf_layout_layer() -> str:
    """Re-lays the dashboard slides out like the reference PDF (KPI tiles, bar/donut/pie/stacked charts)."""
    return (
        PDF_LAYER_START + "\n<style>" + PDF_LAYER_CSS + "</style>\n<script>" + PDF_LAYER_JS + "</script>\n"
        + PDF_LAYER_END + "\n"
    )