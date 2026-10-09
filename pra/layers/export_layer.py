from __future__ import annotations


EXPORT_LAYER_START = "<!-- PRA-EXPORT-LAYER:START -->"

EXPORT_LAYER_END = "<!-- PRA-EXPORT-LAYER:END -->"

EXPORT_LAYER_JS = r"""(function () {
  'use strict';
  var PAGE_W = 1280, MIN_H = 720, SCALE = 2, busy = false;

  function fileBase() {
    var parts = (document.title || '').split('|');
    var name = (parts[1] || 'Report').trim().replace(/[^A-Za-z0-9]+/g, '_').replace(/^_+|_+$/g, '');
    return 'PRA_' + (name || 'Report');
  }

  function addButtons() {
    var host = document.querySelector('.report-meta');
    if (!host || document.getElementById('praExportBtns')) return;
    var wrap = document.createElement('div');
    wrap.id = 'praExportBtns';
    wrap.className = 'pra-export';
    wrap.innerHTML =
      '<button type="button" data-fmt="pdf" title="Download the whole report as a PDF">\u2B07 PDF</button>' +
      '<button type="button" data-fmt="pptx" title="Download the whole report as a PowerPoint">\u2B07 PPT</button>';
    var anchor = host.querySelector('.icon-btn');
    host.insertBefore(wrap, anchor || null);
    wrap.addEventListener('click', function (e) {
      var b = e.target.closest('button[data-fmt]');
      if (b) exportReport(b.getAttribute('data-fmt'));
    });
  }

  function overlay() {
    var o = document.createElement('div');
    o.id = 'praExportOverlay';
    o.innerHTML = '<div class="pra-box"><div class="pra-spin"></div><b id="praExportMsg">Preparing\u2026</b>' +
      '<small>Please keep this tab open. It takes about 20\u201340 seconds.</small></div>';
    document.body.appendChild(o);
    return o;
  }

  function wait(ms) { return new Promise(function (r) { setTimeout(r, ms); }); }

  async function settle(el) {
    try {
      if (window.Chart && Chart.instances) {
        Object.values(Chart.instances).forEach(function (c) {
          if (c && c.canvas && el.contains(c.canvas)) {
            c.options.animation = false; c.resize(); c.update('none');
          }
        });
      }
    } catch (e) {}
    try {
      if (window.Plotly) el.querySelectorAll('.js-plotly-plot').forEach(function (p) { Plotly.Plots.resize(p); });
    } catch (e) {}
    await wait(700);
  }

  function onClone(doc) {
    // Cross-origin iframes (embedded Google map) cannot be rasterised, so show a clear placeholder instead of a blank box.
    var ink = (getComputedStyle(document.documentElement).getPropertyValue('--ink') || '').trim() || '#e6edf5';
    doc.querySelectorAll('svg.pdf-svg').forEach(function (s) { s.style.color = ink; });
    doc.querySelectorAll('iframe').forEach(function (f) {
      var d = doc.createElement('div');
      d.style.cssText = 'display:flex;align-items:center;justify-content:center;width:100%;min-height:320px;' +
        'background:#eef2f7;color:#64748b;font:600 13px Arial,sans-serif;text-align:center;border-radius:8px;padding:16px;';
      d.textContent = 'Interactive map \u2013 open the HTML report to explore locations';
      f.parentNode.replaceChild(d, f);
    });
  }

  async function captureSlides(setMsg) {
    var slides = Array.prototype.slice.call(document.querySelectorAll('#deckView .slide'));
    var visible = slides.map(function (s) { return !s.classList.contains('hidden'); });
    var scrollY = window.scrollY;
    var out = [];
    window.scrollTo(0, 0);
    try {
      for (var i = 0; i < slides.length; i++) {
        setMsg('Capturing slide ' + (i + 1) + ' of ' + slides.length + '\u2026');
        slides.forEach(function (s, k) { s.classList.toggle('hidden', k !== i); });
        var el = slides[i], saved = el.style.cssText;
        el.style.cssText += ';width:' + PAGE_W + 'px;max-width:none;margin:0;min-height:' + MIN_H + 'px;';
        await settle(el);
        var canvas = await html2canvas(el, {
          scale: SCALE, useCORS: true, backgroundColor: document.documentElement.getAttribute('data-theme') === 'dark' ? '#131c2a' : '#ffffff', logging: false,
          windowWidth: 1440, onclone: onClone
        });
        out.push({ data: canvas.toDataURL('image/jpeg', 0.92), w: canvas.width / SCALE, h: canvas.height / SCALE });
        el.style.cssText = saved;
      }
    } finally {
      slides.forEach(function (s, k) { s.classList.toggle('hidden', !visible[k]); s.style.removeProperty('width'); });
      window.dispatchEvent(new Event('resize'));
      window.scrollTo(0, scrollY);
    }
    return out;
  }

  function buildPdf(imgs, name) {
    var JsPDF = window.jspdf.jsPDF;
    var pdf = new JsPDF({ orientation: 'l', unit: 'px', format: [imgs[0].w, imgs[0].h], hotfixes: ['px_scaling'], compress: true });
    imgs.forEach(function (im, i) {
      if (i > 0) pdf.addPage([im.w, im.h], im.w >= im.h ? 'l' : 'p');
      pdf.addImage(im.data, 'JPEG', 0, 0, im.w, im.h, undefined, 'FAST');
    });
    pdf.save(name + '.pdf');
  }

  function buildPptx(imgs, name) {
    var pptx = new PptxGenJS();
    pptx.layout = 'LAYOUT_WIDE'; // 13.33 x 7.5 in
    pptx.title = (document.title || 'Program Analysis Report');
    var BW = 13.333, BH = 7.5;
    imgs.forEach(function (im) {
      var s = pptx.addSlide();
      s.background = { color: document.documentElement.getAttribute('data-theme') === 'dark' ? '131C2A' : 'FFFFFF' };
      var r = Math.min(BW / im.w, BH / im.h), w = im.w * r, h = im.h * r;
      s.addImage({ data: im.data, x: (BW - w) / 2, y: (BH - h) / 2, w: w, h: h });
    });
    return pptx.writeFile({ fileName: name + '.pptx' });
  }

  async function exportReport(fmt) {
    if (busy) return;
    if (!window.html2canvas || (fmt === 'pdf' && !window.jspdf) || (fmt === 'pptx' && !window.PptxGenJS)) {
      alert('Export libraries could not be loaded. Please check your internet connection and reload the report.');
      return;
    }
    busy = true;
    var btns = document.querySelectorAll('#praExportBtns button');
    btns.forEach(function (b) { b.disabled = true; });
    var o = overlay(), msgEl = o.querySelector('#praExportMsg');
    var setMsg = function (t) { msgEl.textContent = t; };
    try {
      var imgs = await captureSlides(setMsg);
      setMsg(fmt === 'pdf' ? 'Building PDF\u2026' : 'Building PowerPoint\u2026');
      await wait(50);
      if (fmt === 'pdf') buildPdf(imgs, fileBase()); else await buildPptx(imgs, fileBase());
    } catch (err) {
      console.error(err);
      alert('Could not create the ' + (fmt === 'pdf' ? 'PDF' : 'PowerPoint') + ': ' + (err && err.message ? err.message : err));
    } finally {
      o.remove();
      btns.forEach(function (b) { b.disabled = false; });
      busy = false;
    }
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', addButtons); else addButtons();
})();
"""

EXPORT_LAYER_CSS = """
.topbar{gap:12px!important}
.brand-large{flex:1 1 0!important;min-width:0!important}
.brand-large .brand-logo{flex:0 0 auto;width:clamp(180px,22vw,320px)!important}
.brand-copy{min-width:0!important;overflow:hidden}
.brand-copy h1,.brand-copy p{white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.report-meta{gap:10px!important}
@media(max-width:1500px){.report-meta .meta-block:first-of-type{display:none}.pra-export button,.pra-theme-btn{padding-left:10px!important;padding-right:10px!important}}
@media(max-width:1250px){.report-meta .meta-block{display:none}.brand-copy{display:none}}
.pra-export{display:flex;gap:8px;align-items:center}
.pra-export button{background:#f45132;color:#fff;border:0;border-radius:8px;padding:8px 13px;font:800 11px Arial,sans-serif;letter-spacing:.04em;cursor:pointer;box-shadow:0 4px 12px rgba(244,81,50,.25);white-space:nowrap}
.pra-export button:nth-child(2){background:#3f5bd8;box-shadow:0 4px 12px rgba(63,91,216,.25)}
.pra-export button:hover{filter:brightness(1.08)}
.pra-export button:disabled{opacity:.55;cursor:wait}
#praExportOverlay{position:fixed;inset:0;z-index:99999;background:rgba(23,32,51,.78);display:flex;align-items:center;justify-content:center;font-family:Arial,sans-serif}
#praExportOverlay .pra-box{background:#fff;border-radius:14px;padding:26px 34px;text-align:center;box-shadow:0 20px 60px rgba(0,0,0,.35);min-width:300px}
#praExportOverlay b{display:block;color:#172033;font-size:15px;margin:12px 0 6px}
#praExportOverlay small{color:#667085;font-size:11px}
.pra-spin{width:34px;height:34px;margin:0 auto;border:4px solid #e4e9ef;border-top-color:#f45132;border-radius:50%;animation:praSpin .8s linear infinite}
@keyframes praSpin{to{transform:rotate(360deg)}}
@media print{.pra-export,#praExportOverlay{display:none!important}}
"""

def export_layer() -> str:
    return (
        EXPORT_LAYER_START + "\n"
        "<style>" + EXPORT_LAYER_CSS + "</style>\n"
        '<script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script>\n'
        '<script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js"></script>\n'
        '<script src="https://cdnjs.cloudflare.com/ajax/libs/pptxgenjs/3.12.0/pptxgen.bundle.js"></script>\n'
        "<script>" + EXPORT_LAYER_JS + "</script>\n"
        + EXPORT_LAYER_END + "\n"
    )