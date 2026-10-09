from __future__ import annotations

import json
import pandas as pd
import re
from typing import Any
from urllib.parse import quote

from pra.data import build_matrix, chart_payload
from pra.helpers import esc, find_column, pct, pct_text, safe_json, top_counts


PRA_EXTRA_JS = r'''// PRA-EXTRA: count + percentage labels, SVG location map slide, decision-maker mix
var PRA_PAL=['#4e7fd8','#f28a3d','#12a8b8','#18a878','#7c5ce5','#e66aa4'];
if(typeof PRA!=='undefined'&&PRA.palette&&PRA.palette.length>=2)PRA_PAL=PRA.palette.slice();
function praC(i){return PRA_PAL[i%PRA_PAL.length];}
function praFmtN(n){return Number(n||0).toLocaleString('en-US');}
function praEsc(t){return String(t).replace(/[&<>"']/g,function(c){return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c];});}

/* ---------------- Chart style: 2D / 3D ---------------- */
var PRA_MODE=(function(){try{return localStorage.getItem('praChartMode')||'3d';}catch(e){return '3d';}})();
var PRA_DEPTH=9;
function praIs3D(){return PRA_MODE==='3d';}
function praShade(c,f){var m=/^#([0-9a-f]{6})$/i.exec(c||'');if(!m)return c;var n=parseInt(m[1],16),r=n>>16,g=(n>>8)&255,b=n&255;
  var t=function(v){return Math.round(f<0?v*(1+f):v+(255-v)*f);};return 'rgb('+t(r)+','+t(g)+','+t(b)+')';}
function praColorAt(ds,i){var c=ds.praColors||ds.hoverBackgroundColor||ds.backgroundColor;return Array.isArray(c)?c[i%c.length]:c;}
function praDepthOf(chart,el){var horiz=chart.options.indexAxis==='y';var s=horiz?el.height:el.width;return Math.max(3,Math.min(PRA_DEPTH,(s||20)*0.35));}

if (typeof Chart !== 'undefined') {
  // 3D extrusion: side/top faces for bars, a raised base for doughnuts. Drawn before the front faces.
  Chart.register({id:'praDepth', beforeDatasetsDraw:function(chart){
    if(!praIs3D())return; var ctx=chart.ctx, t=chart.config.type;
    ctx.save();
    chart.data.datasets.forEach(function(ds,di){
      var meta=chart.getDatasetMeta(di); if(meta.hidden)return;
      meta.data.forEach(function(el,i){
        var col=praColorAt(ds,i); if(typeof col!=='string')return;
        if(t==='doughnut'||t==='pie'){
          if(!chart.getDataVisibility(i))return;
          var D=PRA_DEPTH; ctx.fillStyle=praShade(col,-0.35);
          for(var k=D;k>=1;k--){ctx.beginPath();ctx.arc(el.x,el.y+k,el.outerRadius,el.startAngle,el.endAngle);
            ctx.arc(el.x,el.y+k,el.innerRadius,el.endAngle,el.startAngle,true);ctx.closePath();ctx.fill();}
          return;
        }
        if(t!=='bar')return;
        var horiz=chart.options.indexAxis==='y', d=praDepthOf(chart,el);
        var L,R,T,B;
        if(horiz){L=Math.min(el.x,el.base);R=Math.max(el.x,el.base);T=el.y-el.height/2;B=el.y+el.height/2;}
        else{L=el.x-el.width/2;R=el.x+el.width/2;T=Math.min(el.y,el.base);B=Math.max(el.y,el.base);}
        if(!(R-L>0.5)||!(B-T>0.5))return;
        ctx.fillStyle=praShade(col,0.35);
        ctx.beginPath();ctx.moveTo(L,T);ctx.lineTo(L+d,T-d);ctx.lineTo(R+d,T-d);ctx.lineTo(R,T);ctx.closePath();ctx.fill();
        ctx.fillStyle=praShade(col,-0.3);
        ctx.beginPath();ctx.moveTo(R,T);ctx.lineTo(R+d,T-d);ctx.lineTo(R+d,B-d);ctx.lineTo(R,B);ctx.closePath();ctx.fill();
      });
    });
    ctx.restore();
  }});

  // Count + percentage labels that never overlap: one line, two lines, count only, or hidden (tooltip still works).
  Chart.register({id:'praLabels', afterDatasetsDraw:function(chart){
    var o=chart.options.plugins&&chart.options.plugins.praLabels; if(!o)return;
    var ctx=chart.ctx, horiz=chart.options.indexAxis==='y', t=chart.config.type, d3=praIs3D();
    ctx.save(); ctx.textBaseline='middle';
    chart.data.datasets.forEach(function(ds,di){
      var meta=chart.getDatasetMeta(di); if(meta.hidden)return;
      var sum=(ds.data||[]).reduce(function(a,b){return a+Number(b||0);},0);
      meta.data.forEach(function(el,i){
        var c,p;
        if(o.grouped){c=Number(ds.data[i]||0); if(!c)return; p=sum?c/sum*100:0;}
        else{c=o.counts[i]; if(c==null)return; p=o.pcts?o.pcts[i]:(o.total?c/o.total*100:0);}
        if(t==='doughnut'){
          if(!chart.getDataVisibility(i)||p<7)return; var pos=el.tooltipPosition(); ctx.fillStyle='#fff'; ctx.textAlign='center';
          ctx.shadowColor='rgba(0,0,0,.35)'; ctx.shadowBlur=3;
          ctx.font='bold 11px Arial'; ctx.fillText(p.toFixed(0)+'%',pos.x,pos.y-6);
          ctx.font='10px Arial'; ctx.fillText(praFmtN(c),pos.x,pos.y+7); ctx.shadowBlur=0; return;
        }
        if(t==='line'){ctx.font='bold 10px Arial';ctx.fillStyle='#172033';ctx.textAlign='center';ctx.fillText(praFmtN(c)+' ('+p.toFixed(1)+'%)',el.x,el.y-12);return;}
        var dep=d3?praDepthOf(chart,el):0, full=praFmtN(c)+' ('+p.toFixed(1)+'%)', cnt=praFmtN(c), pc=p.toFixed(1)+'%';
        ctx.font='bold 10px Arial'; ctx.fillStyle='#172033';
        if(horiz){ctx.textAlign='left'; ctx.fillText(o.grouped?cnt:full,el.x+6+dep,el.y-dep/2); return;}
        ctx.textAlign='center';
        var slot=o.grouped?el.width+4:el.width/0.72, x=el.x+dep/2, y=el.y-10-dep;
        if(!o.grouped&&ctx.measureText(full).width<=slot-4){ctx.fillText(full,x,y);return;}
        if(!o.grouped&&Math.max(ctx.measureText(cnt).width,ctx.measureText(pc).width)<=slot-2){
          ctx.fillText(cnt,x,y-12); ctx.font='9px Arial'; ctx.fillStyle='#64748b'; ctx.fillText(pc,x,y); return;}
        if(ctx.measureText(cnt).width<=slot)ctx.fillText(cnt,x,y);
      });
    });
    ctx.restore();
  }});
}

function praAlpha(c,a){var m=/^#([0-9a-f]{6})$/i.exec(c);if(!m)return c;var n=parseInt(m[1],16);return 'rgba('+(n>>16)+','+((n>>8)&255)+','+(n&255)+','+a+')';}
function praGrad(colors,horiz){return function(c){var ch=c.chart,a=ch.chartArea;
  var col=Array.isArray(colors)?colors[c.dataIndex%colors.length]:colors; if(!a||typeof col!=='string')return col;
  if(praIs3D())return col;
  var g=horiz?ch.ctx.createLinearGradient(a.left,0,a.right,0):ch.ctx.createLinearGradient(0,a.bottom,0,a.top);
  g.addColorStop(0,praAlpha(col,0.45)); g.addColorStop(1,col); return g;};}
var PRA_TIP={backgroundColor:'rgba(23,32,51,.92)',padding:10,cornerRadius:8,titleFont:{size:11,weight:'bold'},bodyFont:{size:11},displayColors:true,boxPadding:4};
function praTrim(l,n){n=n||30;l=String(l);return l.length>n?l.slice(0,n-1)+'…':l;}
// Wrap an axis label onto at most two lines so it is never clipped at the canvas edge.
function praWrap(l,max){max=max||18;var words=String(l).split(/\s+/),lines=[''];
  words.forEach(function(w){var cur=lines[lines.length-1];if(cur&&(cur+' '+w).length>max)lines.push(w);else lines[lines.length-1]=cur?cur+' '+w:w;});
  if(lines.length>2){lines=[lines[0],lines.slice(1).join(' ')];}
  return lines.map(function(x){return praTrim(x,max+2);});}

// Chooses the chart form from the data: few parts of a whole -> doughnut, too many columns -> horizontal bars,
// a 1-2 point line -> column chart.
function praPickType(el,type,labels,options){
  var n=labels.length, horiz=options.indexAxis==='y';
  if(options.pie&&n>=2&&n<=options.pie)return {type:'doughnut',horiz:false};
  if(type==='line'&&n<3)return {type:'bar',horiz:false};
  if(type==='bar'&&!horiz){
    var w=(el.parentElement&&el.parentElement.clientWidth)||0;
    if(w?w/Math.max(n,1)<52:n>5)return {type:'bar',horiz:true};
  }
  return {type:type,horiz:horiz};
}

function praRebuildChart(canvasId,type,labels,values,options){
  options=options||{};
  var el=document.getElementById(canvasId); if(!el||typeof Chart==='undefined')return;
  try{var old=Chart.getChart(el); if(old)old.destroy();}catch(e){}
  var pick=praPickType(el,type,labels,options); type=pick.type;
  var bg=options.bg||PRA_PAL, counts=options.counts||null, total=options.total||0;
  var isDo=type==='doughnut', horiz=pick.horiz, d3=praIs3D();
  if(isDo&&options.pieValues)values=options.pieValues;
  if(isDo&&!Array.isArray(bg))bg=PRA_PAL;
  var pctOf=function(i){return options.pcts?options.pcts[i]:(total?counts[i]/total*100:0);};
  var lab=(isDo&&counts)?labels.map(function(l,i){return praTrim(l,20)+' · '+pctOf(i).toFixed(1)+'%';}):labels;
  var catAxis={grid:{display:false},afterFit:function(sc){if(horiz)sc.width+=8;},ticks:{color:'#334155',font:{size:9},autoSkip:false,maxRotation:horiz?0:40,
    callback:function(v){return horiz?praWrap(this.getLabelForValue(v),20):praWrap(this.getLabelForValue(v),labels.length<=3?24:14);}}};
  var valAxis={beginAtZero:true,grid:{color:'#edf0f4'},border:{display:false},ticks:{color:'#64748b',font:{size:9},
    callback:function(v){return options.percent?Number(v).toFixed(0)+'%':praFmtN(v);}}};
  if(options.percent&&!isDo){var mx=Math.max.apply(null,values.concat([1]));valAxis.suggestedMax=Math.min(100,mx*1.15);}
  var scales={}; if(!isDo){scales[horiz?'x':'y']=valAxis; scales[horiz?'y':'x']=catAxis;}
  var lc=options.border||praC(1);
  var lineFill=function(c){var a=c.chart.chartArea; if(!a)return praAlpha(lc,.14);
    var g=c.chart.ctx.createLinearGradient(0,a.top,0,a.bottom); g.addColorStop(0,praAlpha(lc,.35)); g.addColorStop(1,praAlpha(lc,0)); return g;};
  var dep=d3?PRA_DEPTH:0;
  new Chart(el,{type:type,
    data:{labels:lab,datasets:[{data:values,praColors:bg,
      backgroundColor:type==='bar'?praGrad(bg,horiz):(type==='line'?lineFill:bg),hoverBackgroundColor:type==='bar'?bg:undefined,
      borderColor:type==='line'?lc:(isDo?'#fff':bg),
      borderWidth:type==='line'?3:(isDo?2:0),borderRadius:isDo?0:(d3?0:6),borderSkipped:false,maxBarThickness:44,
      pointRadius:5,pointHoverRadius:7,pointBackgroundColor:'#fff',pointBorderColor:lc,pointBorderWidth:2,
      tension:0.35,fill:type==='line',hoverOffset:isDo?10:0}]},
    options:{responsive:true,maintainAspectRatio:false,indexAxis:horiz?'y':'x',cutout:isDo?(options.solidPie?'0%':'55%'):undefined,
      animation:{duration:900,easing:'easeOutQuart'},
      layout:{padding:isDo?{top:4,left:4,right:4,bottom:4+dep}:{right:horiz?96+dep:12+dep,top:horiz?4+dep:34+dep,left:10,bottom:4}},
      plugins:{
        legend:{display:isDo||!!options.legend,position:'bottom',
          labels:{color:'#475569',font:{size:9},boxWidth:10,usePointStyle:true,pointStyle:'circle',padding:8}},
        tooltip:Object.assign({},PRA_TIP,{callbacks:{label:function(c){var i=c.dataIndex;
          if(counts){return ' '+labels[i]+': '+praFmtN(counts[i])+' ('+pctOf(i).toFixed(1)+'%)';}
          return ' '+labels[i]+': '+c.raw+(options.percent?'%':'');}}}),
        praLabels:counts?{counts:counts,total:total,pcts:isDo?null:(options.pcts||null)}:false
      },
      scales:scales}});
  return {type:type,horiz:horiz};
}

function praRebuildGroupedChart(canvasId,labels,datasets,options){
  options=options||{};
  var el=document.getElementById(canvasId); if(!el||typeof Chart==='undefined')return;
  try{var old=Chart.getChart(el); if(old)old.destroy();}catch(e){}
  var horiz=options.indexAxis==='y', d3=praIs3D(), dep=d3?PRA_DEPTH:0;
  datasets.forEach(function(d,i){var col=d.praColors||d.backgroundColor||PRA_PAL[i%6];
    d.praColors=col; d.borderRadius=d3?0:5; d.borderSkipped=false; d.maxBarThickness=34;
    d.hoverBackgroundColor=col; d.backgroundColor=praGrad(col,horiz);});
  var scales={};
  scales[horiz?'x':'y']={beginAtZero:true,grid:{color:'#edf0f4'},border:{display:false},ticks:{color:'#64748b',font:{size:9},precision:0}};
  scales[horiz?'y':'x']={grid:{display:false},afterFit:function(sc){if(horiz)sc.width+=8;},ticks:{color:'#334155',font:{size:9},autoSkip:false,maxRotation:horiz?0:30,
    callback:function(v){return horiz?praWrap(this.getLabelForValue(v),20):praWrap(this.getLabelForValue(v),14);}}};
  new Chart(el,{type:'bar',data:{labels:labels,datasets:datasets},
    options:{responsive:true,maintainAspectRatio:false,indexAxis:horiz?'y':'x',
      animation:{duration:900,easing:'easeOutQuart'},
      layout:{padding:{right:(horiz?44:8)+dep,top:(horiz?4:20)+dep}},
      plugins:{legend:{display:!!options.legend,position:'bottom',
          labels:{color:'#475569',font:{size:9},boxWidth:10,usePointStyle:true,pointStyle:'circle',
            generateLabels:function(ch){return ch.data.datasets.map(function(ds,i){var c=Array.isArray(ds.praColors)?ds.praColors[0]:ds.praColors;
              return {text:praTrim(ds.label,32),fillStyle:c,strokeStyle:c,pointStyle:'circle',hidden:!ch.isDatasetVisible(i),datasetIndex:i};});}}},
        tooltip:Object.assign({},PRA_TIP,{callbacks:{label:function(c){
          var tot=(c.dataset.data||[]).reduce(function(a,b){return a+Number(b||0);},0);
          var p=tot?c.raw/tot*100:0;
          return ' '+c.dataset.label+': '+praFmtN(c.raw)+' ('+p.toFixed(1)+'% of this series)';}}}),
        praLabels:{grouped:true}},
      scales:scales}});
}

function praUpdateCharts(){
  var N=PRA.lead_count||1;
  var sum=function(a){return a.reduce(function(x,y){return x+Number(y||0);},0);};
  var pcts=function(a){return a.map(function(v){return +(v/N*100).toFixed(2);});};
  praRebuildChart('jobLevel','bar',PRA.job_labels,PRA.job_values,{bg:PRA_PAL,counts:PRA.job_values,total:N});
  // Job-level split is a part-of-whole view -> pie, so it differs from the Job Level columns.
  praRebuildChart('jobSplit','bar',PRA.job_split_labels,pcts(PRA.job_split_values),{bg:PRA_PAL,indexAxis:'y',percent:true,
    counts:PRA.job_split_values,total:N,pie:6,pieValues:PRA.job_split_values,solidPie:true});
  praRebuildChart('jobFunctions','line',PRA.func_labels,pcts(PRA.func_values),{bg:PRA_PAL,percent:true,counts:PRA.func_values,total:N});
  praRebuildChart('industries','bar',PRA.industry_labels,pcts(PRA.industry_values),{bg:PRA_PAL,indexAxis:'y',percent:true,counts:PRA.industry_values,total:N});
  var sz=praRebuildChart('employeeSize','bar',PRA.size_labels,pcts(PRA.size_values),{bg:PRA_PAL,percent:true,counts:PRA.size_values,total:N});
  var s4=document.getElementById('s4');
  if(s4&&sz&&sz.horiz){praRenameSection(s4,'Employee Size — Column','Employee Size');}
  var dt=sum(PRA.device_values)||1;
  praRebuildChart('devices','bar',PRA.device_labels,PRA.device_values.map(function(v){return +(v/dt*100).toFixed(2);}),
    {bg:[praC(2),praC(1),praC(0),praC(3)],indexAxis:'y',percent:true,counts:PRA.device_values,total:dt,pie:4,pieValues:PRA.device_values});
  praRebuildChart('assetSplit','doughnut',PRA.asset_labels.slice(0,6),PRA.asset_values.slice(0,6),{bg:PRA_PAL,legend:true,counts:PRA.asset_values.slice(0,6),total:N});

  var indDs=PRA.industry_labels.slice(0,2).map(function(ind,i){return {label:ind,
    data:PRA.asset_names.map(function(_,a){return PRA.asset_industry_matrix?PRA.asset_industry_matrix[a][i]:0;}),
    backgroundColor:praC(i)};});
  praRebuildGroupedChart('assetIndustry',PRA.asset_names,indDs,{legend:true});
  praRebuildGroupedChart('assetSize',PRA.size_names,PRA.asset_names.map(function(a,i){return {label:a,data:PRA.asset_size_matrix[i]||[],backgroundColor:PRA_PAL[i%6]};}),{legend:true});
  praRebuildGroupedChart('assetJob',PRA.job_names,PRA.asset_names.map(function(a,i){return {label:a,data:PRA.asset_job_matrix[i]||[],backgroundColor:PRA_PAL[i%6]};}),{legend:true,indexAxis:'y'});

  [['openChart',PRA.asset_open_labels,PRA.asset_open_values],['clickChart',PRA.asset_click_labels,PRA.asset_click_values],['conversionChart',PRA.asset_conv_labels,PRA.asset_conv_values]].forEach(function(x){
    praRebuildChart(x[0],'bar',x[1],x[2],{bg:PRA_PAL,indexAxis:'y',counts:x[2],total:sum(x[2])||1});
  });
  var rates=[PRA.delivery_rate||0,PRA.open_rate||0,PRA.click_rate||0,PRA.conversion_rate||0,PRA.bounce_rate||0];
  praRebuildChart('statsChart','bar',['Delivered','Open','Clicks','Conversion','Bounce'],rates,
    {bg:[praC(2),praC(0),praC(1),praC(3),praC(4)],indexAxis:'y',percent:true,
     counts:[PRA.delivered,PRA.opens,PRA.clicks,PRA.conversion,PRA.bounced],pcts:rates});
  if(document.getElementById('praGeoChart'))praRenderGeoChart();
}
function praRenderGeoChart(){
  praRebuildChart('praGeoChart','bar',PRA.country_labels,PRA.country_values,{bg:praC(1),indexAxis:'y',counts:PRA.country_values,total:PRA.geo_total||1});
}

/* 2D / 3D switch (floating, hidden when printing). */
function praSetMode(m){PRA_MODE=m; try{localStorage.setItem('praChartMode',m);}catch(e){}
  document.querySelectorAll('#praModeSwitch button').forEach(function(b){b.classList.toggle('on',b.dataset.mode===m);});
  praUpdateCharts();}
document.addEventListener('DOMContentLoaded',function(){
  var st=document.createElement('style');
  st.textContent='#praModeSwitch{position:fixed;left:18px;bottom:18px;z-index:9999;display:flex;gap:2px;padding:3px;border-radius:999px;'+
    'background:#fff;box-shadow:0 6px 20px rgba(15,23,42,.18);border:1px solid #e2e8f0;font:600 12px Arial}'+
    '#praModeSwitch span{padding:6px 8px 6px 10px;color:#64748b}'+
    '#praModeSwitch button{border:0;background:transparent;padding:6px 12px;border-radius:999px;cursor:pointer;color:#334155;font:inherit}'+
    '#praModeSwitch button.on{background:linear-gradient(90deg,#f47b20,#f59e0b);color:#fff}'+
    '@media print{#praModeSwitch{display:none}}'+
    '.grid>.panel,.grid>div{min-width:0}.chartbox{min-width:0;overflow:hidden}.chartbox canvas{max-width:100%}';
  document.head.appendChild(st);
  var sw=document.createElement('div'); sw.id='praModeSwitch';
  sw.innerHTML='<span>Charts</span><button data-mode="2d">2D</button><button data-mode="3d">3D</button>';
  sw.addEventListener('click',function(e){var b=e.target.closest('button'); if(b)praSetMode(b.dataset.mode);});
  document.body.appendChild(sw);
  sw.querySelectorAll('button').forEach(function(b){b.classList.toggle('on',b.dataset.mode===PRA_MODE);});
});

/* ---------------- Location slide: self-contained SVG bubble map ----------------
   One bubble per country in the Location Split. Colours are exactly the pie's: ranks 1-7 get their own slice
   colour and every other country shares the "Others" colour. Plain SVG = no Google key, no iframe, no internet,
   and it is captured correctly by the PDF / PowerPoint export. */
var PRA_PIE_COLORS=['#70ad47','#ffc000','#ed7d31','#4472c4','#5b9bd5','#a5a5a5','#9e480e','#7c5ce5'];
var PRA_LAND={"na":[[-168,66],[-162,70],[-156,71.3],[-141,69.6],[-128,70],[-115,68.5],[-108,68],[-95,68],[-90,69],[-85,69.5],[-82,67],[-86,64],[-93,61],[-94.5,58.7],[-91,57],[-85,55.3],[-82,52.5],[-79,51.5],[-79.5,54.5],[-77,57],[-78,62],[-72,62.3],[-65,60.5],[-62,58],[-60,55.5],[-56,52.5],[-60,50.2],[-66,50],[-70,47],[-65,49.2],[-64.5,46.2],[-61,45.7],[-66,44.3],[-70,43.5],[-70.5,41.8],[-74,40.5],[-76,37],[-75.5,35.2],[-78,33.8],[-81,31.5],[-80,27],[-80.2,25.2],[-81.8,26.2],[-82.8,28.8],[-84,30],[-86.5,30.3],[-89.5,30.2],[-91,29.3],[-94,29.6],[-97.2,27.8],[-97.5,24.5],[-97.8,22],[-96,19],[-94.5,18.2],[-91.5,18.5],[-90.5,21],[-87,21.5],[-88,18.5],[-88.5,16],[-84,15.8],[-83.2,14.5],[-83.7,11],[-81.8,9],[-79.5,9.5],[-77.4,8.6],[-77.5,8.3],[-79.5,8.8],[-81.5,8],[-83.5,8.7],[-85.7,10],[-87.5,13],[-91.5,14],[-94.5,16],[-97,15.8],[-101,17.3],[-105.5,20],[-105.5,23],[-109,25.5],[-112.5,29.5],[-114.7,31.7],[-117.1,32.5],[-118.5,34],[-120.6,34.6],[-122.5,37.5],[-124.3,40.3],[-124,46],[-124.7,48.4],[-123,49],[-127,51],[-130,54.5],[-134,58],[-139,59.8],[-146,60.7],[-152,59],[-156,57],[-162,55],[-158,58.5],[-162,60],[-165,62],[-161,64.5],[-168,65.7]],"baja":[[-114.7,31.7],[-115.8,30.5],[-114,28],[-112.2,25.5],[-110,23],[-109.5,23.4],[-111.5,26],[-113,29],[-114.7,31.7]],"sa":[[-77.3,8.5],[-75.5,10.8],[-72,12],[-71,11],[-68,10.6],[-64,10.6],[-61.5,10.5],[-60,8.5],[-57,6],[-54,5.8],[-51.5,4.2],[-50,1.5],[-48,-0.8],[-44.5,-2.4],[-40,-2.8],[-35.2,-5.5],[-35,-9],[-38.5,-13.2],[-39,-17.8],[-41,-22],[-44,-23.2],[-48.5,-26],[-48.7,-28.5],[-52,-32],[-54,-34.8],[-57,-35],[-58.4,-34.3],[-57.2,-38],[-62,-39],[-62.3,-41],[-65,-41],[-64.3,-43],[-67.5,-46],[-66,-48],[-69,-51],[-68.5,-53],[-71,-54],[-74,-52],[-75.5,-47],[-73.7,-42],[-73.5,-37],[-71.7,-33],[-71.5,-28],[-70.3,-18],[-76,-14],[-79,-8],[-81.2,-5.8],[-80,-3],[-80.2,-1],[-79.8,1.5],[-78,2.5],[-77.5,6]],"gl":[[-73,78.5],[-65,81.5],[-40,83.4],[-20,82],[-18,77],[-20,72],[-22,70],[-26,68],[-33,66.5],[-40,65],[-43,60],[-48,61],[-52,65],[-54,69],[-58,75],[-68,76.2]],"eu":[[-9.5,37],[-9.3,39],[-9,43],[-2,43.5],[-1.5,46],[-4.5,48.3],[-1.5,49.7],[1.5,50.5],[4,51.5],[8,53.8],[8.5,57],[10.5,57.7],[10.5,55],[12,54.3],[14,54],[19,54.5],[21,57],[24,57],[24,59.4],[28,59.5],[30,60],[23,60],[21.5,61],[21,63],[25,65],[22,65.8],[18,63],[17,61],[19,60],[16.5,57],[14,55.5],[12.5,56.3],[11,59],[8,58],[5.5,59],[5,62],[10,64],[14,67.5],[18,69.5],[25,71],[31,70],[33,69],[41,67],[34,66.4],[35,64.5],[38,64.5],[41,66.5],[44,66],[44,68.5],[53,68.5],[58,68.8],[60,69.8],[68,68.5],[69,73],[73,72.5],[80,73.5],[87,75],[100,76.2],[104,77.7],[113,74],[128,72],[140,72.5],[150,71.3],[160,69.5],[170,70],[180,69],[180,65],[178,64.5],[172,64.5],[170,60],[163,59.8],[163,57],[156.7,51],[156,57],[160,61],[155,59.5],[143,59.3],[137,54],[141,52.5],[140.5,48.5],[135,43.8],[131,42.5],[129.5,41],[128,39],[129.5,36],[129,35],[126.5,34.5],[126.5,37.5],[125,39.5],[121.5,39],[121.5,40.8],[118,39],[119,37.2],[122.5,37],[119.5,35],[121.5,32],[122,30],[121,28],[119,25],[116.5,22.7],[113,22],[110.5,21.2],[109.5,19.7],[108,21.5],[106.7,20],[105.8,18.5],[109,15],[109.2,11.7],[107,10.4],[105,8.6],[104.8,10.2],[103,11],[100.5,13.3],[99.5,11],[100,8.5],[102,6],[103.5,4],[103.5,1.3],[101,2.8],[100.3,5.5],[98.3,8.3],[98.5,10.5],[98.5,13.5],[97.6,16.5],[95.3,15.8],[94.3,18],[92.2,21.5],[91.5,22.8],[90,22],[87,21.5],[86.5,19.5],[84,18],[82,16.5],[80.2,15.5],[80,13],[79.8,10.3],[78,8.5],[77,8],[76,10],[75,12.5],[73.5,16],[72.8,19],[72.7,21.7],[70.5,20.8],[69,22.3],[70.5,23.1],[68.5,23.7],[67,24.8],[66.5,25.4],[61.5,25.2],[57.5,25.7],[56.5,27],[54,26.7],[51.5,27.9],[50,30],[48.8,30.2],[48.5,28.5],[50.5,26.2],[51.5,24.5],[54,24.2],[56,26.2],[56.5,24.5],[58.7,23.5],[59.8,22.5],[58.5,20.5],[55.5,17.5],[52,16.2],[48,14],[45,12.8],[43.3,12.7],[42.8,15],[41,19],[39,21.5],[37,25],[35,28],[34.8,29.5],[34.3,31.2],[35,33],[36,34.5],[36,36.5],[34,36.2],[31,36.8],[28.5,36.7],[27,38.5],[26.5,40.3],[24,40.5],[23,39.5],[24,38],[22.5,36.5],[21.5,37],[21,39],[19.5,41],[19,42],[15.5,45],[13.7,45.2],[12.3,45.4],[12.5,44],[14,42.5],[16,41.8],[18.5,40.2],[17,39],[16.5,38],[15.7,38.2],[16,40],[14,40.8],[12,42],[10.5,43.5],[8.8,44.4],[7,43.7],[4,43.5],[3,42.5],[0.5,40.5],[-0.3,38.5],[-2,36.7],[-5.3,36],[-6.3,36.8],[-8.8,37.2]],"af":[[-17,21],[-16.5,24],[-13,27.7],[-10,29.5],[-9.8,31.5],[-6.8,34],[-5.9,35.8],[-2,35.1],[3,36.8],[10,37.3],[11,35],[10.2,33.5],[15,32.3],[20,32],[20,30.5],[24,32],[29,31],[32.3,31.3],[32.6,29.9],[33.5,28],[35,24],[37.2,21],[38.5,18],[41,14.5],[43.2,12.7],[44.5,10.5],[51,11.8],[51,10],[48,5],[44,1],[41.5,-1.8],[39.2,-5],[39,-8],[40.5,-11],[40.5,-15],[37,-17.8],[35,-20],[35.5,-24],[32.8,-26],[32.5,-28.5],[30,-31.5],[27,-33.7],[22,-34.2],[19,-34.8],[17.7,-32],[15.2,-27],[14.5,-22.5],[11.8,-17.2],[13.5,-12],[13,-9],[12,-5],[9,-1],[9.5,3.5],[8,4.5],[5,5.8],[2,6.3],[-2,4.8],[-5,5.2],[-7.7,4.4],[-10.5,6.3],[-13.2,8.3],[-15.2,11.2],[-16.8,13],[-17.5,14.7],[-16.5,16.5],[-16.2,19.5]],"au":[[113.5,-22],[114.5,-26],[115,-33.8],[118,-35],[123,-34],[126,-32.2],[131,-31.5],[135,-34.8],[138,-35.2],[140,-37.5],[144,-38.5],[147,-38.8],[150,-37.3],[153,-31],[153.2,-27],[150.5,-22.5],[146.2,-19],[145.4,-14.8],[143.5,-14],[142.3,-10.8],[141.5,-13],[139.5,-17.5],[136.5,-15.8],[135.5,-12],[132.5,-11.4],[130.5,-12.5],[129.2,-15],[126,-14],[122.3,-17.5],[121,-19.6],[116.5,-20.7]],"uk":[[-5.5,50],[1.3,51.2],[1.7,52.8],[0,53.5],[-1.5,55.5],[-2,57.5],[-4,57.7],[-3,58.6],[-5.2,58.6],[-6,56.5],[-4.8,55],[-3,54.8],[-3.2,53.3],[-4.7,52.7],[-5.2,51.7],[-3,51.4],[-4.5,50.4]],"ie":[[-10,51.8],[-6,52.2],[-6,54.2],[-8,55.3],[-10,54],[-9.5,52.5]],"is":[[-24,65.5],[-22,66.4],[-16,66.5],[-13.5,65],[-18,63.5],[-22.5,63.8]],"jp":[[130.8,31],[132,34],[135,34.5],[136.8,34.3],[140,35],[141,38.5],[142,40.5],[141.2,41.5],[140,40],[139.5,38],[137,37],[136,36],[133,35.5],[131,34.5],[130,33.5]],"hk":[[140,42],[141.5,45.4],[145.5,43.3],[143,42],[141,42]],"tw":[[120.2,23],[121.5,25.2],[122,24.5],[120.8,22]],"lz":[[120.5,18.5],[122.2,18.4],[121.5,15],[124,13],[121.5,13.8],[120.6,14.2],[119.9,16.3]],"md":[[122,7],[125.5,9.8],[126.5,7],[125.5,5.8],[123,7.3]],"bo":[[109,1.5],[109.5,-0.5],[110.5,-3],[114,-4],[116,-3.5],[116.5,-1],[117.8,1],[119,5],[117,7],[115.5,5],[113,3],[111,1.8]],"su":[[95.3,5.6],[98,4],[100.5,2],[104,-1],[106,-3],[105.7,-5.9],[102,-4],[99.5,-1],[97,2.5]],"ja":[[105.2,-6.8],[108,-6.3],[111,-6.5],[114.4,-7.7],[114.5,-8.7],[110,-8.3],[106.5,-7.4]],"sl":[[119.5,-5.5],[120.5,-3],[121.5,-1],[120,1],[124.8,1.5],[123,0.5],[121.2,-1.8],[123,-4],[122,-5],[120.5,-5.6]],"ng":[[131,-0.8],[134,-0.7],[138,-1.8],[141,-2.6],[145,-4.5],[147.5,-6],[150.5,-10.5],[147,-10],[143.5,-8.5],[141,-9.2],[138,-8.4],[138.5,-7],[135,-4.5],[132.5,-3.3]],"nzn":[[172.7,-34.5],[175,-37],[178.5,-37.7],[177,-39.5],[175.2,-41.5],[174.8,-39.3],[173.8,-39.2],[174.8,-37]],"nzs":[[172.7,-40.5],[174.3,-41.7],[173,-43.5],[171,-44.5],[169,-46.6],[166.5,-46],[168.5,-44],[171.5,-41.8]],"mg":[[49.3,-12],[50.4,-15.5],[49.5,-17.5],[47.3,-24.8],[45,-25.6],[43.3,-22],[44.4,-16.5],[47,-15]],"lk":[[79.8,9.8],[81.8,7.5],[81,6],[80,6.2],[79.8,8]],"cu":[[-85,21.9],[-82,23.1],[-77,21.5],[-74.2,20.3],[-77.5,19.9],[-80,21.7]],"hi":[[-74.4,19.8],[-71.7,19.9],[-68.4,18.6],[-70,18.2],[-73.5,18.2]],"nf":[[-59.3,47.7],[-56,51.5],[-53,47],[-55.5,46.9]],"baf":[[-80,73.5],[-68,70],[-62,67],[-65,63],[-72,64.5],[-77,66],[-85,70.5]],"vic":[[-118,70],[-105,73],[-101,69],[-112,68.5]],"ell":[[-90,76.5],[-75,79],[-62,82.5],[-80,83],[-92,81],[-95,77]],"nz2":[[51.5,71.5],[56,75],[68,77],[60,75],[55,71.5]]};
var PRA_WATER={"black":[[28,41.2],[29,45],[31,46.6],[33.5,46],[33,44.6],[36.5,45.3],[38,47],[39.5,43.5],[41.5,41.7],[37,41],[33,42],[29,41.2]],"casp":[[47,45],[49,46.5],[53,46.5],[53.5,44],[51,41.5],[53,40],[54,37.5],[51,36.8],[49,38.5],[49,40.5],[47.5,43]]};
function praS5Visible(){var s=document.getElementById('s5');return !!s&&!s.classList.contains('hidden');}
function praGeoColor(i){return PRA_PIE_COLORS[i<7?i:7];}
function praTextOn(hex){var m=/^#([0-9a-f]{6})$/i.exec(hex||'');if(!m)return '#fff';var n=parseInt(m[1],16);
  return (0.299*(n>>16)+0.587*((n>>8)&255)+0.114*(n&255))>165?'#1e293b':'#fff';}

function praRenderGeoMap(){
  var host=document.getElementById('geoMap'); if(!host)return;
  var names=PRA.country_labels||[], vals=PRA.country_values||[], N=PRA.geo_total||1;
  var pts={}; ((PRA.geo_spec&&PRA.geo_spec.points)||[]).forEach(function(p){pts[String(p[0]).toLowerCase()]={lat:p[1],lon:p[2]};});
  var rows=[], missing=((PRA.geo_spec&&PRA.geo_spec.unplaced)||[]).slice();
  names.forEach(function(n,i){var p=pts[String(n).toLowerCase()];
    if(p)rows.push({name:n,v:vals[i],i:i,lat:p.lat,lon:p.lon}); else if(missing.indexOf(n)<0)missing.push(n);});
  host.innerHTML=''; host.style.overflow='hidden';
  if(!rows.length){host.innerHTML='<div style="padding:24px;font:13px Arial;color:#64748b">No mappable countries were found in the raw lead file.</div>';return;}

  /* frame the map around the plotted countries (whole world when they are far apart) */
  var lons=rows.map(function(r){return r.lon;}), lats=rows.map(function(r){return r.lat;});
  var minLo=Math.min.apply(null,lons), maxLo=Math.max.apply(null,lons), minLa=Math.min.apply(null,lats), maxLa=Math.max.apply(null,lats);
  var ASPECT=2.4, sLon=Math.max((maxLo-minLo)*1.4+30,80), sLat=Math.max((maxLa-minLa)*1.4+20,40);
  if(sLon/sLat<ASPECT)sLon=sLat*ASPECT; else sLat=sLon/ASPECT;
  if(sLat>142){sLat=142;sLon=Math.min(360,sLat*ASPECT);}
  var x0=Math.min(Math.max((minLo+maxLo)/2-sLon/2,-180),180-sLon);
  var top=Math.min(84,Math.max((minLa+maxLa)/2+sLat/2,-58+sLat));
  var W=1000, S=W/sLon, MH=sLat*S;
  var X=function(lon){return (lon-x0)*S;}, Y=function(lat){return (top-lat)*S;};
  var d=function(poly){return 'M'+poly.map(function(p){return X(p[0]).toFixed(1)+','+Y(p[1]).toFixed(1);}).join('L')+'Z';};

  var grid='';
  for(var lo=Math.ceil(x0/30)*30;lo<=x0+sLon;lo+=30)grid+='<line x1="'+X(lo).toFixed(1)+'" y1="0" x2="'+X(lo).toFixed(1)+'" y2="'+MH.toFixed(1)+'"/>';
  for(var la=Math.ceil((top-sLat)/30)*30;la<=top;la+=30)grid+='<line x1="0" y1="'+Y(la).toFixed(1)+'" x2="'+W+'" y2="'+Y(la).toFixed(1)+'"/>';
  var holes=Object.keys(PRA_WATER).map(function(k){return d(PRA_WATER[k]);}).join('');
  var land=Object.keys(PRA_LAND).map(function(k){return '<path fill-rule="evenodd" d="'+d(PRA_LAND[k])+(k==='eu'?holes:'')+'"/>';}).join('');

  var max=Math.max.apply(null,rows.map(function(r){return r.v;}).concat([1]));
  rows.forEach(function(r){r.rad=11+17*Math.sqrt(r.v/max); r.cx=X(r.lon); r.cy=Y(r.lat); r.x0=r.cx; r.y0=r.cy;});
  /* nearby countries (e.g. UK / Netherlands / France / Germany) would sit on top of each other: nudge them
     apart just enough to stay readable, never more than ~45 map units from their true position */
  for(var it=0;it<80;it++){var moved=false;
    for(var a=0;a<rows.length;a++)for(var b=a+1;b<rows.length;b++){
      var A=rows[a],B=rows[b],dx=B.cx-A.cx,dy=B.cy-A.cy,dist=Math.sqrt(dx*dx+dy*dy)||0.01,need=(A.rad+B.rad)*0.92;
      if(dist<need){var push=(need-dist)/2,ux=dx/dist,uy=dy/dist; if(dist<0.5){ux=1;uy=0;}
        A.cx-=ux*push;A.cy-=uy*push;B.cx+=ux*push;B.cy+=uy*push;moved=true;}}
    rows.forEach(function(r){var ox=r.cx-r.x0,oy=r.cy-r.y0,od=Math.sqrt(ox*ox+oy*oy); if(od>45){r.cx=r.x0+ox/od*45;r.cy=r.y0+oy/od*45;}});
    if(!moved)break;}
  var dots=rows.slice().sort(function(a,b){return b.v-a.v;}).map(function(r){
    var col=praGeoColor(r.i), rad=r.rad, cx=r.cx, cy=r.cy, pc=(r.v/N*100).toFixed(1);
    return '<g class="pra-geo-dot" data-name="'+praEsc(r.name)+'" data-count="'+r.v+'" style="cursor:pointer">'+
      '<title>'+praEsc(r.name)+': '+praFmtN(r.v)+' leads ('+pc+'%)</title>'+
      '<circle cx="'+cx.toFixed(1)+'" cy="'+cy.toFixed(1)+'" r="'+rad.toFixed(1)+'" fill="'+col+'" fill-opacity=".92" stroke="#fff" stroke-width="2"/>'+
      (rad>=14?'<text x="'+cx.toFixed(1)+'" y="'+(cy+5).toFixed(1)+'" text-anchor="middle" font-size="15" font-weight="700" fill="'+praTextOn(col)+'" style="pointer-events:none">'+praFmtN(r.v)+'</text>':'')+
      '</g>';}).join('');

  /* legend = the pie's legend: top 7 + Others */
  var items=names.slice(0,7).map(function(n,i){return {t:n+' \u00b7 '+praFmtN(vals[i]),c:praGeoColor(i)};});
  if(names.length>7){var rest=vals.slice(7).reduce(function(a,b){return a+b;},0); items.push({t:'Others ('+(names.length-7)+') \u00b7 '+praFmtN(rest),c:praGeoColor(7)});}
  var lx=14,ly=MH+24,leg='';
  items.forEach(function(it){var w=24+it.t.length*7.4+18; if(lx+w>W-8){lx=14;ly+=22;}
    leg+='<circle cx="'+(lx+7)+'" cy="'+(ly-5)+'" r="7" fill="'+it.c+'"/><text x="'+(lx+20)+'" y="'+ly+'" font-size="15" font-weight="600" fill="currentColor">'+praEsc(it.t)+'</text>'; lx+=w;});
  if(missing.length){ly+=22; leg+='<text x="14" y="'+ly+'" font-size="13" fill="#b45309">Not plotted (no map position for): '+praEsc(missing.join(', '))+'</text>';}
  var H=ly+12;

  host.innerHTML='<svg viewBox="0 0 '+W+' '+H.toFixed(0)+'" preserveAspectRatio="xMidYMid meet" xmlns="http://www.w3.org/2000/svg" role="img" '+
    'aria-label="Map of lead locations" style="display:block;width:100%;height:100%;color:inherit" font-family="Calibri,Segoe UI,Arial,sans-serif">'+
    '<defs><clipPath id="praGeoClip"><rect x="0" y="0" width="'+W+'" height="'+MH.toFixed(1)+'"/></clipPath></defs>'+
    '<g clip-path="url(#praGeoClip)"><g stroke="currentColor" stroke-opacity=".12" stroke-width="1">'+grid+'</g>'+
    '<g fill="currentColor" fill-opacity=".2" stroke="currentColor" stroke-opacity=".28" stroke-width="1" stroke-linejoin="round">'+land+'</g>'+
    '</g>'+
    '<g id="praGeoDots">'+dots+'</g><g id="praGeoTip"></g>'+leg+'</svg>';
  var svg=host.querySelector('svg');
  svg.addEventListener('click',function(e){var g=e.target.closest&&e.target.closest('.pra-geo-dot'); if(g)praFocusCountry(g.getAttribute('data-name'));});
}

function praFocusCountry(name){
  var N=PRA.geo_total||1;
  document.querySelectorAll('#s5 .geo-item').forEach(function(it){
    var a=it.querySelector('a'); var on=a&&a.dataset.mapQuery===name;
    it.style.background=on?'#fff4ea':''; it.style.borderLeftWidth=on?'6px':'';
  });
  var host=document.getElementById('geoMap'), svg=host&&host.querySelector('svg'); if(!svg)return;
  var hit=null;
  svg.querySelectorAll('.pra-geo-dot').forEach(function(g){var on=g.getAttribute('data-name')===name, c=g.querySelector('circle');
    c.setAttribute('stroke',on?'#0f172a':'#fff'); c.setAttribute('stroke-width',on?4:2); if(on)hit=g;});
  var tip=svg.querySelector('#praGeoTip'); tip.innerHTML=''; if(!hit)return;
  svg.querySelector('#praGeoDots').appendChild(hit);
  var c=hit.querySelector('circle'), cx=+c.getAttribute('cx'), cy=+c.getAttribute('cy'), r=+c.getAttribute('r'), n=+hit.getAttribute('data-count');
  var txt=name+' \u00b7 '+praFmtN(n)+' leads \u00b7 '+(n/N*100).toFixed(1)+'%', w=txt.length*7.6+22;
  var x=Math.min(Math.max(cx-w/2,4),996-w), y=cy-r-34; if(y<4)y=cy+r+8;
  tip.innerHTML='<rect x="'+x.toFixed(1)+'" y="'+y.toFixed(1)+'" width="'+w.toFixed(1)+'" height="26" rx="6" fill="#0f172a" fill-opacity=".92"/>'+
    '<text x="'+(x+w/2).toFixed(1)+'" y="'+(y+18).toFixed(1)+'" text-anchor="middle" font-size="15" font-weight="700" fill="#fff">'+praEsc(txt)+'</text>';
}

function praUpdateSlide5(){
  var s=document.getElementById('s5'); if(!s)return;
  var names=PRA.country_labels, vals=PRA.country_values, N=PRA.geo_total||1, unique=names.length;
  var mv=s.querySelector('.metric .value'); if(mv)mv.textContent=String(unique);
  var ml=s.querySelector('.metric .label'); if(ml){var t=ml.lastChild; if(t&&t.nodeType===3)t.nodeValue='Unique Countries';}
  var title=s.querySelector('.title');
  if(title)title.textContent='Location Split \u00b7 '+(unique<=3?names.join(' \u00b7 '):names.slice(0,2).join(' \u00b7 ')+' +'+(unique-2)+' more');
  var navItems=document.querySelectorAll('.slide-nav-item'); if(navItems[4]){var sm=navItems[4].querySelector('small'); if(sm)sm.textContent='Location Map';}
  var geoList=s.querySelector('.geo-list');
  if(geoList){geoList.innerHTML=PRA.geo_rows; geoList.style.maxHeight='230px'; geoList.style.overflowY='auto';
    geoList.addEventListener('click',function(e){var a=e.target.closest('a[data-map-query]'); if(!a)return; e.preventDefault(); praFocusCountry(a.dataset.mapQuery);});}
  var callout=s.querySelector('.callout');
  if(callout){
    var k=Math.min(5,unique), top=vals.slice(0,k).reduce(function(a,b){return a+b;},0);
    callout.textContent=unique?('Top '+k+' location'+(k>1?'s':'')+' ('+names.slice(0,k).join(', ')+') account for '+praFmtN(top)+' of '+praFmtN(N)+' leads ('+(top/N*100).toFixed(1)+'%).'):'No geographic data is available in the raw lead file.';
  }
  praRenderGeoMap();
  try{window.geoInitialized=true;}catch(e){}
  try{window.initMap=function(){return true;};}catch(e){}
}

/* ---------------- Decision-maker mix replaces the made-up device split ---------------- */
function praRenameSection(root,oldT,newT){
  root.querySelectorAll('.section').forEach(function(sec){
    var c=sec.cloneNode(true); var ic=c.querySelector('.section-icon'); if(ic)ic.remove();
    if(c.textContent.trim()!==oldT)return;
    var nodes=[].filter.call(sec.childNodes,function(n){return n.nodeType===3;});
    if(nodes.length)nodes[nodes.length-1].nodeValue=newT; else sec.appendChild(document.createTextNode(newT));
  });
}
document.addEventListener('DOMContentLoaded',function(){
  if(PRA.device_title==='Decision Maker Mix'){
    var s4=document.getElementById('s4');
    if(s4){
      praRenameSection(s4,'Devices','Decision Makers vs Influencers');
      praRenameSection(s4,'Device Mix','Decision Maker Mix');
      var t=s4.querySelector('.title'); if(t)t.textContent='Industry, Company Size & Decision Maker Profile';
    }
    var nav=document.querySelectorAll('.slide-nav-item')[3]; if(nav){var b=nav.querySelector('b'); if(b)b.textContent='Industry, Decision Makers & Size';}
  }
});
'''

def js_dynamic_layer(data: dict[str, Any], df: pd.DataFrame, geo_spec: dict[str, Any]) -> str:
    inputs_maps_key = data.get("maps_key", "")
    # JSON data used by browser-side JS for text + charts.
    top_job = data["job_levels"][0] if data["job_levels"] else ("N/A", 0)
    top_industries = data["industries"][:2]
    top_sizes = data["company_sizes"]
    top_assets = data["assets"]

    job_labels, job_values = chart_payload(data["job_levels"])
    job_split_labels, job_split_values = chart_payload(data["job_levels"])
    func_labels, func_values = chart_payload(data["job_functions"])
    industry_labels, industry_values = chart_payload(data["industries"][:6])
    size_labels, size_values = chart_payload(top_sizes)
    asset_labels, asset_values = chart_payload(top_assets)
    decision_labels, decision_values = chart_payload(data["decisions"])
    device_labels, device_values = chart_payload(data["devices"])
    asset_open_labels, asset_open_values = chart_payload(data["asset_opens"][:6])
    asset_click_labels, asset_click_values = chart_payload(data["asset_clicks"][:6])
    asset_conv_labels, asset_conv_values = chart_payload(data["asset_conversions"][:6])

    # Industry x Asset matrix for Slide 6. The table uses its first two columns;
    # the chart can use all six top industries.
    industry_names = [x[0] for x in data["industries"][:6]]
    asset_names = [x[0] for x in top_assets]
    matrix = build_matrix(df, data["asset_col"], data["industry_col"], asset_names, industry_names)
    _country_names = [c for c, _ in top_counts(df, find_column(df, ["Country", "Country Name"]), 6)]

    # Asset x employee size / job level datasets
    company_size_col = find_column(df, ["Company Size", "Employee Size", "Company Size Range"])
    job_level_col = find_column(df, ["Job Level", "Seniority", "Seniority Level"])

    size_names = [x[0] for x in data["company_sizes"][:6]]
    asset_size_matrix = build_matrix(df, data["asset_col"], company_size_col, asset_names, size_names)

    job_names = [x[0] for x in data["job_levels"][:6]]
    asset_job_matrix = build_matrix(df, data["asset_col"], job_level_col, asset_names, job_names)

    # Build list HTML for dynamic slide 5.
    geo_sorted = sorted(data["geo"], key=lambda x: x[3], reverse=True)
    geo_top = geo_sorted[:5]
    geo_other_count = sum(x[3] for x in geo_sorted[5:])
    geo_other_pct = pct(geo_other_count, data["lead_count"]) or 0

    geo_rows = "".join(
      f'<div class="geo-item"><a href="https://www.google.com/maps/search/?api=1&amp;query={quote(name)}" '
      f'data-map-query="{esc(name)}" target="_blank" rel="noopener">{esc(name)}</a>'
      f'<b>{count:,} leads · {pct_text(pct(count, data["geo_total"]))}</b></div>'
      for name, count in data["country_counts"]
    )

    # Dynamic top two industries.
    industry_split_html = ""
    if top_industries:
        cols = []
        for i, (name, count) in enumerate(top_industries):
            color = "var(--blue)" if i == 0 else "var(--orange)"
            cols.append(
                f'<div><div style="font-size:30px;color:{color};font-weight:900">{pct_text(pct(count, data["lead_count"]))}</div>'
                f'<div style="font-size:10px;color:#64748b">{esc(name)}</div></div>'
            )
        industry_split_html = "".join(cols)
    else:
        industry_split_html = '<div><div style="font-size:14px;color:#64748b;font-weight:800">No industry data</div></div>'

    # Dynamic device mix text.
    device_total = sum(count for _, count in data["devices"])
    if data["devices"]:
        device_mix_text = " &nbsp; ".join(
            f'<b style="color:{"var(--cyan)" if i == 0 else "var(--orange)"}">{pct_text(pct(count, device_total))}</b> {esc(name)} ({count:,})'
            for i, (name, count) in enumerate(data["devices"])
        )
        if data["device_mix_is_estimate"]:
            device_mix_text += '<div style="margin-top:6px;color:#64748b;font-size:10px">Illustrative allocation; device type is not present in the source workbook.</div>'
    else:
        device_mix_text = '<span style="color:#64748b">Device data not available in the raw file.</span>'

    # Asset table on slide 6.
    table_asset_rows = ""
    for i, (asset, count) in enumerate(top_assets):
        industries_for_row = matrix[i] if i < len(matrix) else [0] * len(industry_names)
        c1 = industries_for_row[0] if len(industries_for_row) >= 1 else 0
        c2 = industries_for_row[1] if len(industries_for_row) >= 2 else 0
        table_asset_rows += (
            f"<tr><td>{esc(asset)}</td><td class='num'>{c1:,}</td><td class='num'>{c2:,}</td>"
            f"<td class='num'>{count:,}</td></tr>"
        )

    industry_headers = industry_names + ["Industry 2"]
    if len(industry_names) == 1:
        industry_headers = [industry_names[0], "—"]

    # Senior audience cards — retain the three-card format.
    senior_candidates = []
    for name in ["C-Level", "Director", "Vice President"]:
        for actual_name, count in data["job_levels"]:
            if actual_name.strip().lower() == name.strip().lower():
                senior_candidates.append((actual_name, count))
                break
    if len(senior_candidates) < 3:
        for pair in data["job_levels"]:
            if pair not in senior_candidates:
                senior_candidates.append(pair)
            if len(senior_candidates) == 3:
                break

    while len(senior_candidates) < 3:
        senior_candidates.append(("N/A", 0))

    senior_cards = "".join(
        f'<div class="audience-stat"><span class="pill">{esc(name)}</span><b>{count:,}</b>'
        f'<small>{pct_text(pct(count, data["lead_count"]))}</small></div>'
        for name, count in senior_candidates[:3]
    )

    top_job_name, top_job_count = top_job
    top_asset_name = top_assets[0][0] if top_assets else "N/A"
    top_asset_count = top_assets[0][1] if top_assets else 0
    top_industry_name = top_industries[0][0] if top_industries else "N/A"
    top_industry_count = top_industries[0][1] if top_industries else 0
    top_size_name = top_sizes[0][0] if top_sizes else "N/A"
    top_size_count = top_sizes[0][1] if top_sizes else 0

    priority_locations = ", ".join(x[0] for x in geo_top[:5]) if geo_top else "No geographic data"
    priority_location_share = pct(sum(x[3] for x in geo_top), data["lead_count"]) if geo_top else None

    duration_text = (
        f'{data["business_days"]} business days'
        if data["business_days"] is not None
        else "the supplied campaign period"
    )
    lead_type_labels = [name for name, _ in data.get("lead_types", [])]
    has_bant = any(re.search(r"\bBANT\b", name, flags=re.IGNORECASE) for name in lead_type_labels)
    has_cs = any(re.search(r"\bCS\b", name, flags=re.IGNORECASE) for name in lead_type_labels)

    if has_bant and has_cs:
        engagement_observation = (
            "BANT potential customers were engaged through telephonic calls and promotional emails; "
            "CS potential customers were engaged through promotional emails."
        )
    elif has_bant:
        engagement_observation = "Potential customers were engaged through telephonic calls and promotional emails."
    elif has_cs:
        engagement_observation = "Potential customers were engaged through promotional emails."
    else:
        engagement_observation = None

    # These observations intentionally use only measured lead-level / campaign inputs.
    observations = [
        f'Campaign generated {data["lead_count"]:,} leads in {duration_text}.',
        *([engagement_observation] if engagement_observation else []),
        f'{esc(top_job_name)} is the largest job-level segment with {top_job_count:,} leads ({pct_text(pct(top_job_count, data["lead_count"]))}).',
        f'{esc(top_industry_name)} is the largest industry segment with {top_industry_count:,} leads ({pct_text(pct(top_industry_count, data["lead_count"]))}).',
        f'{esc(top_size_name)} is the largest company-size segment with {top_size_count:,} leads ({pct_text(pct(top_size_count, data["lead_count"]))}).',
        f'{esc(top_asset_name)} is the largest assigned asset segment with {top_asset_count:,} leads ({pct_text(pct(top_asset_count, data["lead_count"]))}).',
        f'Delivery rate is {pct_text(data["delivery_rate"])}, open rate is {pct_text(data["open_rate"])}, click rate is {pct_text(data["click_rate"])}, and conversion rate is {pct_text(data["conversion_rate"])}.',
    ]

    recommendations = [
        f'Prioritize {esc(top_job_name)} and {esc(top_industry_name)} audiences in the next campaign cycle based on the current lead mix.',
        f'Use the strongest asset segment, {esc(top_asset_name)}, as a starting point for follow-up and asset planning.',
        f'Focus geographic follow-up on {esc(priority_locations)} where the top five locations represent {pct_text(priority_location_share)} of the raw lead base.' if priority_location_share is not None else 'Add complete location data before applying geographic prioritization.',
    ]

    # Encode everything into browser JS.
    data_json = safe_json({
        "campaign_name": data["campaign_name"],
        "report_date": data["report_date"],
        "period": data.get("period", data["report_date"]),
        "palette": data.get("palette") or [],
        "start_date": data["start_date"],
        "end_date": data["end_date"],
        "prepared_by": data["prepared_by"],
        "lead_count": data["lead_count"],
        "sent": data["sent"],
        "delivered": data["delivered"],
        "opens": data["opens"],
        "clicks": data["clicks"],
        "conversion": data["conversion"],
        "bounced": data["bounced"],
        "delivery_rate": data["delivery_rate"],
        "open_rate": data["open_rate"],
        "click_rate": data["click_rate"],
        "conversion_rate": data["conversion_rate"],
        "bounce_rate": data["bounce_rate"],
        "business_days": data["business_days"],
        "job_labels": json.loads(job_labels),
        "job_values": json.loads(job_values),
        "job_split_labels": json.loads(job_split_labels),
        "job_split_values": json.loads(job_split_values),
        "func_labels": json.loads(func_labels),
        "func_values": json.loads(func_values),
        "industry_labels": json.loads(industry_labels),
        "industry_values": json.loads(industry_values),
        "size_labels": json.loads(size_labels),
        "size_values": json.loads(size_values),
        "asset_labels": json.loads(asset_labels),
        "asset_values": json.loads(asset_values),
        "unique_asset_count": data["unique_asset_count"],
        "unique_industry_count": data["unique_industry_count"],
        "unique_country_count": data["unique_country_count"],
        "country_names": _country_names,
        "asset_country_matrix": build_matrix(df, data["asset_col"], find_column(df, ["Country", "Country Name"]), asset_names, _country_names),
        "decision_labels": json.loads(decision_labels),
        "decision_values": json.loads(decision_values),
        "device_labels": json.loads(device_labels),
        "device_values": json.loads(device_values),
        "device_mix_is_estimate": data["device_mix_is_estimate"],
        "device_title": data["device_title"],
        "device_is_pct": data.get("device_is_pct", False),
        "maps_key": inputs_maps_key,
        "asset_open_labels": json.loads(asset_open_labels),
        "asset_open_values": json.loads(asset_open_values),
        "asset_click_labels": json.loads(asset_click_labels),
        "asset_click_values": json.loads(asset_click_values),
        "asset_conv_labels": json.loads(asset_conv_labels),
        "asset_conv_values": json.loads(asset_conv_values),
        "country_labels": [name for name, _ in data["country_counts"]],
        "country_values": [count for _, count in data["country_counts"]],
        "asset_event_note": data["asset_event_note"],
        "geo_total": data["geo_total"],
        "top_job_name": top_job_name,
        "top_job_count": top_job_count,
        "top_industry_name": top_industry_name,
        "top_industry_count": top_industry_count,
        "top_size_name": top_size_name,
        "top_size_count": top_size_count,
        "top_asset_name": top_asset_name,
        "top_asset_count": top_asset_count,
        "senior_cards": senior_cards,
        "industry_split_html": industry_split_html,
        "device_mix_text": device_mix_text,
        "geo_rows": geo_rows,
        "table_asset_rows": table_asset_rows,
        "industry_names": industry_names,
        "asset_names": asset_names,
        "size_names": size_names,
        "asset_size_matrix": asset_size_matrix,
        "job_names": job_names,
        "asset_job_matrix": asset_job_matrix,
        "observations": observations,
        "recommendations": recommendations,
        "geo_spec": geo_spec,
    })

    # Use JSON inside a script tag; escape </script> if a campaign name contains it.
    data_json = data_json.replace("</script>", "<\\/script>")

    # JS helper code. The original HTML/CSS is untouched; this layer only updates
    # the existing DOM and recreates the existing Chart.js charts.
    dynamic_layer = f"""
<!-- ==========================================================
     DYNAMIC PRA LAYER - generated by generate_dynamic_pra.py
     ========================================================== -->
<script>
const PRA = {data_json};

function praText(root, exactText, newText) {{
  const walker = document.createTreeWalker(root || document.body, NodeFilter.SHOW_TEXT);
  let node;
  while ((node = walker.nextNode())) {{
    if (node.nodeValue.trim() === exactText) {{
      node.nodeValue = node.nodeValue.replace(exactText, String(newText));
      return true;
    }}
  }}
  return false;
}}

function praAllText(root, oldText, newText) {{
  const walker = document.createTreeWalker(root || document.body, NodeFilter.SHOW_TEXT);
  let node;
  while ((node = walker.nextNode())) {{
    if (node.nodeValue.includes(oldText)) {{
      node.nodeValue = node.nodeValue.split(oldText).join(String(newText));
    }}
  }}
}}

function praNumber(v) {{
  return v === null || v === undefined ? 'N/A' : Number(v).toLocaleString('en-US');
}}

function praPct(v) {{
  return v === null || v === undefined ? 'N/A' : Number(v).toFixed(2) + '%';
}}

function praRebuildGroupedChart(containerId, labels, datasets, options={{}}) {{
  const canvas = document.getElementById(containerId);
  if (!canvas || typeof Chart === 'undefined') return;
  try {{ const old = Chart.getChart(canvas); if (old) old.destroy(); }} catch (e) {{}}
  new Chart(canvas, {{
    type:'bar',
    data:{{labels:labels,datasets:datasets}},
    options:{{
      responsive:true,
      maintainAspectRatio:false,
      indexAxis:options.indexAxis || 'x',
      plugins:{{legend:{{display:!!options.legend,position:'bottom'}}}},
      scales:{{
        x:{{ticks:{{color:'#64748b',font:{{size:9}}}},grid:{{display:options.indexAxis !== 'y'}}}},
        y:{{beginAtZero:true,ticks:{{color:'#64748b',font:{{size:9}}}},grid:{{color:'#edf0f4'}}}}
      }}
    }}
  }});
}}

function praRebuildChart(canvasId, type, labels, values, options={{}}) {{
  const el = document.getElementById(canvasId);
  if (!el || typeof Chart === 'undefined') return;
  try {{ const old = Chart.getChart(el); if (old) old.destroy(); }} catch (e) {{}}
  const background = options.bg || ['#3f5bd8','#f47b20','#12a8b8','#18a878','#7c5ce5','#e66aa4'];
  const dataset = {{
    data:values,
    backgroundColor:background,
    borderColor:options.border || background,
    borderWidth:options.borderWidth || 0,
    borderRadius:options.radius || 4,
    pointRadius:4,tension:0.28,fill:options.fill || false
  }};
  new Chart(el, {{
    type:type,
    data:{{labels:labels,datasets:[dataset]}},
    options:{{
      responsive:true,
      maintainAspectRatio:false,
      indexAxis:options.indexAxis || 'x',
      plugins:{{legend:{{display:!!options.legend,position:'bottom'}}}},
      scales:type === 'doughnut' ? {{}} : {{
        x:options.indexAxis === 'y'
          ? {{beginAtZero:true,grid:{{color:'#edf0f4'}},ticks:{{callback:function(v){{return options.percent ? Number(v).toFixed(2) + '%' : v;}}}}}}
          : {{grid:{{display:false}}}},
        y:options.indexAxis === 'y'
          ? {{grid:{{display:false}}}}
          : {{beginAtZero:true,grid:{{color:'#edf0f4'}},ticks:{{callback:function(v){{return options.percent ? Number(v).toFixed(2) + '%' : v;}}}}}}
      }}
    }}
  }});
}}

function praUpdateTopbarAndCover() {{
  document.title = 'VAIS | ' + PRA.campaign_name + ' | Program Analysis Report';

  const topH1 = document.querySelector('.brand-copy h1');
  const topPeriod = document.querySelector('.meta-block:nth-of-type(1) b');
  const topPrepared = document.querySelector('.meta-block:nth-of-type(2) b');

  if (topH1) topH1.textContent = PRA.campaign_name;
  if (topPeriod) topPeriod.textContent = PRA.period;
  if (topPrepared) topPrepared.textContent = PRA.prepared_by;

  const coverCampaign = document.querySelector('#s1 .campaign');
  const coverPrepared = document.querySelector('#s1 .prepared');
  if (coverCampaign) coverCampaign.textContent = PRA.campaign_name;
  if (coverPrepared) coverPrepared.innerHTML = 'Prepared by: ' + PRA.prepared_by + '<br/>' + PRA.report_date;
}}

function praUpdateSlide3() {{
  const s = document.getElementById('s3');
  if (!s) return;

  // Lead count
  const leadValue = s.querySelector('.metric .value');
  if (leadValue) leadValue.textContent = praNumber(PRA.lead_count);

  // Senior audience cards
  const stats = s.querySelectorAll('.audience-stat');
  const cards = [
    ...[
      ['C-Level', 'c-level'], ['Director', 'director'], ['Vice President', 'vice president']
    ]
  ];
  const raw = PRA.job_labels.map((n,i)=>({{name:n,count:PRA.job_values[i]}}));
  const preferred = ['c-level','director','vice president'];
  let selected = preferred
    .map(k => raw.find(x => x.name.toLowerCase().trim() === k))
    .filter(Boolean);

  raw.forEach(x => {{
    if (selected.length < 3 && !selected.some(y => y.name === x.name)) selected.push(x);
  }});
  while (selected.length < 3) selected.push({{name:'N/A',count:0}});

  stats.forEach((el, i) => {{
    const pair = selected[i];
    const pill = el.querySelector('.pill');
    const val = el.querySelector('b');
    const small = el.querySelector('small');
    if (pill) pill.textContent = pair.name;
    if (val) val.textContent = praNumber(pair.count);
    if (small) small.textContent = praPct((pair.count / PRA.lead_count) * 100);
  }});

  const callout = s.querySelector('.callout');
  if (callout) {{
    callout.textContent = 'The campaign generated ' + praNumber(PRA.lead_count) +
      ' leads, with ' + PRA.top_job_name + ' contacts representing the largest job-level segment at ' +
      praPct((PRA.top_job_count / PRA.lead_count) * 100) + '.';
  }}
}}

function praUpdateSlide4() {{
  const s = document.getElementById('s4');
  if (!s) return;

  const metricValue = s.querySelector('.metric .value');
  const metricNote = s.querySelector('.metric .note');
  if (metricValue) metricValue.textContent = String(PRA.industry_labels.length);
  if (metricNote) metricNote.textContent = PRA.industry_labels.slice(0,3).join(' • ') || 'No industry data';

  const industrySplit = s.querySelector('.panel .section + .grid.g2');
  const panels = s.querySelectorAll('.panel');
  const sectionTitle = panel => {{
    const section = panel.querySelector('.section');
    if (!section) return '';
    const copy = section.cloneNode(true);
    copy.querySelector('.section-icon')?.remove();
    return copy.textContent.trim();
  }};
  const industryBox = [...panels].find(x => sectionTitle(x) === 'Industry Split');
  const deviceBox = [...panels].find(x => sectionTitle(x) === 'Device Mix');

  if (industryBox) {{
    const grid = industryBox.querySelector('.grid.g2');
    if (grid) grid.innerHTML = PRA.industry_split_html;
  }}
  if (deviceBox) {{
    const wrap = deviceBox.querySelector('div[style*="font-size:13px"]');
    if (wrap) wrap.innerHTML = PRA.device_mix_text;
  }}
}}

function praUpdateSlide5() {{
  const s = document.getElementById('s5');
  if (!s) return;
  const metricValue = s.querySelector('.metric .value');
  if (metricValue) metricValue.textContent = String(PRA.geo_total);

  const geoList = s.querySelector('.geo-list');
  if (geoList) geoList.innerHTML = PRA.geo_rows;

  const callout = s.querySelector('.callout');
  if (callout) {{
    const top = PRA.geo_rows ? 'Geographic concentration is based on the top five locations in the raw lead file.' :
      'No geographic data is available in the raw lead file.';
    callout.textContent = top;
  }}

  const map = document.getElementById('geoMap');
  if (map) {{
    const firstCountry = PRA.geo_rows ? (PRA.geo_rows.match(/data-map-query="([^"]+)"/) || [])[1] : '';
    const mapQuery = encodeURIComponent(firstCountry || 'North America');
    map.innerHTML = '<div style="width:100%;height:210px"><canvas id="praGeoChart"></canvas></div>' +
      '<iframe id="praGoogleMap" title="Google Maps country location" ' +
      'src="https://www.google.com/maps?q=' + mapQuery + '&output=embed" ' +
      'style="width:100%;height:210px;border:0" loading="lazy" referrerpolicy="no-referrer-when-downgrade"></iframe>';
    map.style.height = '430px';
    map.style.overflow = 'hidden';
    const geoCanvas = document.getElementById('praGeoChart');
    if (geoCanvas && typeof Chart !== 'undefined') {{
      new Chart(geoCanvas, {{
        type:'bar',
        data:{{
          labels:PRA.country_labels,
          datasets:[{{label:'Leads',data:PRA.country_values,backgroundColor:'#f47b20',borderRadius:3}}]
        }},
        options:{{indexAxis:'y',responsive:true,maintainAspectRatio:false,
          plugins:{{legend:{{display:false}},tooltip:{{callbacks:{{label:c=>c.raw + ' leads'}}}}}},
          scales:{{x:{{beginAtZero:true,ticks:{{precision:0,color:'#64748b'}},grid:{{color:'#edf0f4'}}}},
            y:{{ticks:{{color:'#334155',font:{{size:10}}}},grid:{{display:false}}}}}}
        }}
      }});
    }}

    if (geoList) geoList.addEventListener('click', function(event) {{
      const link = event.target.closest('a[data-map-query]');
      const frame = document.getElementById('praGoogleMap');
      if (link && frame) frame.src = 'https://www.google.com/maps?q=' + encodeURIComponent(link.dataset.mapQuery) + '&output=embed';
    }});
  }}

  // Prevent the old Leaflet initializer from running on this map.
  try {{ window.geoInitialized = true; }} catch (e) {{}}
  try {{ window.initMap = function() {{ return true; }}; }} catch (e) {{}}
}}

function praUpdateSlide6() {{
  const s = document.getElementById('s6');
  if (!s) return;

  const uniqueAssets = s.querySelector('.metric .value');
  const assetNote = s.querySelector('.metric .note');
  if (uniqueAssets) uniqueAssets.textContent = String(PRA.unique_asset_count);
  if (assetNote) assetNote.textContent = 'Unique assets in source data';

  const table = s.querySelector('table tbody');
  if (table) table.innerHTML = PRA.table_asset_rows;

  const headers = s.querySelectorAll('table th');
  if (headers.length >= 4) {{
    headers[0].textContent = 'Asset';
    headers[1].textContent = PRA.industry_labels[0] || 'Industry 1';
    headers[2].textContent = PRA.industry_labels[1] || 'Industry 2';
    headers[3].textContent = 'Total';
  }}

  const chartGrid = s.querySelector('.grid.g3');
  const industryPanel = s.querySelector('#assetIndustry')?.closest('.panel');
  if (chartGrid && industryPanel) {{
    chartGrid.style.gridTemplateColumns = 'repeat(3, minmax(0, 1fr))';
    chartGrid.style.alignItems = 'start';
    industryPanel.style.gridColumn = 'auto';
    const industryChart = industryPanel.querySelector('.chartbox');
    if (industryChart) industryChart.style.height = '300px';
    if (!document.getElementById('praIndustryChartLayout')) {{
      const style = document.createElement('style');
      style.id = 'praIndustryChartLayout';
      style.textContent = '@media(max-width:760px){{#s6 .grid.g3{{grid-template-columns:1fr!important}}}}';
      document.head.appendChild(style);
    }}
  }}
}}

function praUpdateSlide7() {{
  const s = document.getElementById('s7');
  if (!s) return;
  const callout = s.querySelector('.callout');
  if (callout) {{
    callout.textContent = 'Audience and asset matrices are calculated directly from the raw lead file. The top assigned asset is ' +
      PRA.top_asset_name + ' with ' + praNumber(PRA.top_asset_count) + ' leads.';
  }}
}}

function praUpdateAssetSlide(slideId, metricValue, metricLabel, noteText, tableMode) {{
  const s = document.getElementById(slideId);
  if (!s) return;
  const value = s.querySelector('.asset-hero .value');
  const note = s.querySelector('.asset-hero .note');
  if (value) value.textContent = praNumber(metricValue);
  if (note) {{
    if ((slideId === 's8' || slideId === 's9') && !noteText) {{
      note.remove();
    }} else {{
      note.textContent = noteText;
    }}
  }}

  if (slideId === 's8' || slideId === 's9' || slideId === 's10') {{
    const layout = s.querySelector('.asset-split-layout');
    const legendPanel = s.querySelector('.asset-legend-panel');
    if (legendPanel) legendPanel.remove();
    if (layout) layout.style.gridTemplateColumns = '175px minmax(0, 1fr)';
  }}

  const tbody = s.querySelector('table tbody');
  if (tbody) {{
    const source = tableMode === 'open' ? PRA.asset_open_values :
                  tableMode === 'click' ? PRA.asset_click_values :
                  PRA.asset_conv_values;
    tbody.innerHTML = PRA.asset_labels.slice(0, source.length).map((name,i) => {{
      const share = metricValue ? (source[i] / metricValue * 100) : 0;
      const valueLabel = Number(source[i] || 0).toLocaleString('en-US');
      return '<tr><td>' + name + '</td><td>' + valueLabel + '</td><td>' + share.toFixed(2) + '%</td></tr>';
    }}).join('');
  }}

  const callout = s.querySelector('.callout');
  if (callout) callout.textContent = PRA.asset_event_note;
}}

function praUpdateSlide11() {{
  const s = document.getElementById('s11');
  if (!s) return;
  const cards = s.querySelectorAll('.stats-card');
  const values = [PRA.sent, PRA.delivered, PRA.opens, PRA.clicks, PRA.conversion, PRA.bounced];
  const notes = ['', praPct(PRA.delivery_rate), praPct(PRA.open_rate), praPct(PRA.click_rate), praPct(PRA.conversion_rate), praPct(PRA.bounce_rate)];

  cards.forEach((card, i) => {{
    const value = card.querySelector('.value');
    const note = card.querySelector('.note');
    if (value) value.textContent = praNumber(values[i]);
    if (note && i > 0) note.textContent = notes[i];
  }});

  const readout = s.querySelector('.exec-readout .callout');
  if (readout) {{
    readout.textContent =
      'Delivery rate is ' + praPct(PRA.delivery_rate) +
      ', while the campaign generated ' + praNumber(PRA.opens) +
      ' opens, ' + praNumber(PRA.clicks) +
      ' clicks and ' + praNumber(PRA.conversion) +
      ' conversions from ' + praNumber(PRA.sent) + ' sends.';
  }}
}}

function praUpdateSlide12() {{
  const s = document.getElementById('s12');
  if (!s) return;

  const list = s.querySelector('.observation-list');
  if (list) list.innerHTML = PRA.observations.map(x => '<li>' + x + '</li>').join('');

  const callouts = s.querySelectorAll('.panel:nth-of-type(2) .callout');
  PRA.recommendations.forEach((text, i) => {{
    if (callouts[i]) callouts[i].textContent = text;
  }});
}}

function praUpdateCharts() {{
  praRebuildChart('jobLevel', 'bar',
    PRA.job_labels,
    PRA.job_values,
    {{bg:['#f28a3d','#4e7fd8','#12e7d4','#18a878','#7c5ce5'], indexAxis:'x'}}
  );

  praRebuildChart('jobSplit', 'bar',
    PRA.job_split_labels,
    PRA.job_split_values.map(v => PRA.lead_count ? (v/PRA.lead_count*100) : 0),
    {{bg:['#12e7d4','#8d6de7','#f28a3d','#4e7fd8'], indexAxis:'y', percent:true}}
  );

  praRebuildChart('jobFunctions', 'line',
    PRA.func_labels,
    PRA.func_values.map(v => PRA.lead_count ? (v/PRA.lead_count*100) : 0),
    {{border:'#f47b20', bg:'rgba(244,123,32,.12)', borderWidth:3, fill:true, percent:true}}
  );

  praRebuildChart('industries', 'bar',
    PRA.industry_labels,
    PRA.industry_values.map(v => PRA.lead_count ? (v/PRA.lead_count*100) : 0),
    {{bg:['#4e7fd8','#f28a3d','#12a8b8','#18a878','#7c5ce5','#e66aa4'], indexAxis:'y', percent:true}}
  );

  praRebuildChart('employeeSize', 'bar',
    PRA.size_labels,
    PRA.size_values.map(v => PRA.lead_count ? (v/PRA.lead_count*100) : 0),
    {{bg:'#f28a3d', percent:true}}
  );

  const deviceTotal = PRA.device_values.reduce((sum, value) => sum + Number(value || 0), 0) || 1;
  const deviceShares = PRA.device_values.map(value => Number(value || 0) / deviceTotal * 100);
  praRebuildChart('devices', 'bar', PRA.device_labels, deviceShares,
    {{bg:['#12a8b8','#f47b20','#4e7fd8','#18a878','#7c5ce5','#e66aa4'],indexAxis:'y',percent:true}});

  praRebuildChart('assetSplit', 'doughnut',
    PRA.asset_labels.slice(0,6),
    PRA.asset_values.slice(0,6),
    {{bg:['#4e7fd8','#f28a3d','#12a8b8','#18a878','#7c5ce5','#e66aa4'], legend:true, percent:true}}
  );

  // Asset industry matrix
  const industryDatasets = PRA.industry_labels.slice(0,6).map((industry, i) => ({{
    label: industry,
    data: PRA.asset_names.map((_, a) => PRA.asset_industry_matrix?.[a]?.[i] || 0),
    backgroundColor: PRA_PAL[i % PRA_PAL.length]
  }}));
  praRebuildGroupedChart('assetIndustry', PRA.asset_names, industryDatasets, {{legend:true}});

  // Asset x employee size
  const sizeDatasets = PRA.asset_names.map((asset, i) => ({{
    label: asset,
    data: PRA.asset_size_matrix[i] || [],
    backgroundColor: i === 0 ? '#3f5bd8' : '#f47b20'
  }}));
  praRebuildGroupedChart('assetSize', PRA.size_names, sizeDatasets, {{legend:true}});

  // Asset x job level
  const jobDatasets = PRA.asset_names.map((asset, i) => ({{
    label: asset,
    data: PRA.asset_job_matrix[i] || [],
    backgroundColor: PRA_PAL[i % PRA_PAL.length]
  }}));
  praRebuildGroupedChart('assetJob', PRA.job_names, jobDatasets, {{legend:true,indexAxis:'y'}});

  praRebuildChart('openChart', 'bar',
    PRA.asset_open_labels,
    PRA.asset_open_values,
    {{bg:['#4e7fd8','#f28a3d','#12a8b8','#18a878','#7c5ce5','#e66aa4'], indexAxis:'y'}}
  );

  praRebuildChart('clickChart', 'bar',
    PRA.asset_click_labels,
    PRA.asset_click_values,
    {{bg:['#4e7fd8','#f28a3d','#12a8b8','#18a878','#7c5ce5','#e66aa4'], indexAxis:'y'}}
  );

  praRebuildChart('conversionChart', 'bar',
    PRA.asset_conv_labels,
    PRA.asset_conv_values,
    {{bg:['#4e7fd8','#f28a3d','#12a8b8','#18a878','#7c5ce5','#e66aa4'], indexAxis:'y'}}
  );

  praRebuildChart('statsChart', 'bar',
    ['Delivered','Open','Clicks','Conversion','Bounce'],
    [PRA.delivery_rate || 0, PRA.open_rate || 0, PRA.click_rate || 0, PRA.conversion_rate || 0, PRA.bounce_rate || 0],
    {{bg:['#12e7d4','#4e7fd8','#f28a3d','#69d86e','#8d6de7'], indexAxis:'y', percent:true}}
  );
}}

function praUpdateSlide8to10() {{
  // Open
  praUpdateAssetSlide(
    's8',
    PRA.opens,
    'Total Open Count',
    '',
    'open'
  );

  // Click
  praUpdateAssetSlide(
    's9',
    PRA.clicks,
    'Total Click Count',
    '',
    'click'
  );

  // Conversion
  praUpdateAssetSlide(
    's10',
    PRA.conversion,
    'Total Conversion Count',
    PRA.asset_conv_values.slice(0,2).join(' + ') + ' conversions',
    'conversion'
  );
}}

function praApplyAll() {{
  praUpdateTopbarAndCover();
  praUpdateSlide3();
  praUpdateSlide4();
  praUpdateSlide5();
  praUpdateSlide6();
  praUpdateSlide7();
  praUpdateSlide8to10();
  praUpdateSlide11();

  // Observations/recommendations use safe HTML strings generated by Python.
  const obs = document.querySelector('#s12 .observation-list');
  if (obs) obs.innerHTML = PRA.observations.map(x => '<li>' + x + '</li>').join('');

  const recs = document.querySelectorAll('#s12 .panel');
  const recPanel = [...recs].find(p => p.querySelector('.section')?.textContent.trim() === 'Recommendations');
  if (recPanel) {{
    const c = recPanel.querySelectorAll('.callout');
    PRA.recommendations.forEach((r,i)=>{{ if(c[i]) c[i].innerHTML = r; }});
  }}

  // Rebuild charts after the existing template charts.
  setTimeout(() => {{
    praUpdateChartsDataOnly();
  }}, 0);
}}

function praUpdateChartsDataOnly() {{
  // Add calculated matrix data which is generated in Python.
  PRA.asset_industry_matrix = {safe_json(matrix)};
  praUpdateChartsWithMatrices();
}}

function praUpdateChartsWithMatrices() {{
  // Re-run the same chart generation function but with matrix data available.
  praUpdateCharts();
}}

document.addEventListener('DOMContentLoaded', function() {{
  praApplyAll();
}});
</script>
<script>
""" + PRA_EXTRA_JS + """
</script>
"""

    return dynamic_layer