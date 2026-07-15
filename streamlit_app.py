#!/usr/bin/env python3
"""
Patches make_dashboard.py in place with:
  1. Fixed Alpha/X-Factor column names (matches your actual Raw Data headers).
  2. BM breakdown data (bm_df, per-BM category/dow/top10/bottom10, dailyBM).
  3. New keys added to the data dict.
  4. A <div id="bmFilters"> tab row added to BODY_HTML.
  5. APP_JS fully replaced with the BM-aware version.

Usage:
    python patch_dashboard.py make_dashboard.py
    (writes make_dashboard.py, keeping a make_dashboard.py.bak backup)
"""
import re
import sys
import shutil
from pathlib import Path

NEW_APP_JS = r'''
Chart.defaults.font.family = "'Inter',sans-serif";
Chart.defaults.color = '#8592AE';
Chart.defaults.borderColor = 'rgba(255,255,255,.06)';

const fmtCr = n => '\u20b9' + (n/10000000).toFixed(2) + ' Cr';
const fmtL = n => '\u20b9' + (n/100000).toFixed(1) + ' L';
const fmtShort = n => n>=10000000 ? '\u20b9'+(n/10000000).toFixed(1)+'Cr' : n>=100000 ? '\u20b9'+(n/100000).toFixed(1)+'L' : '\u20b9'+Math.round(n).toLocaleString('en-IN');
const fmtNum = n => Math.round(n).toLocaleString('en-IN');
const fmtDate = s => { const d=new Date(s+'T00:00:00'); return d.toLocaleDateString('en-IN',{day:'2-digit',month:'short',weekday:'short'}); };
const fmtDateShort = s => { const d=new Date(s+'T00:00:00'); return d.getDate(); };

let currentRegion = 'All';
let currentBM = 'All';
let selectedDate = DATA.daily[DATA.daily.length-1].DateStr;
const charts = {};

function scopeType(){ return currentBM !== 'All' ? 'bm' : 'region'; }
function scopeValue(){ return currentBM !== 'All' ? currentBM : currentRegion; }
function scopeLabel(){ return currentBM !== 'All' ? currentBM : (currentRegion==='All' ? 'All Regions' : currentRegion); }

function destroy(id){ if(charts[id]){ charts[id].destroy(); } }

function renderBmFilters(){
  const el = document.getElementById('bmFilters');
  if(!el || !DATA.bmList || !DATA.bmList.length) return;
  el.innerHTML = ['All'].concat(DATA.bmList).map(b=>
    `<button class="tab ${currentBM===b?'active':''}" data-bm="${b}">${b==='All' ? 'All BMs' : b}</button>`
  ).join('');
  el.querySelectorAll('.tab').forEach(btn=>{
    btn.addEventListener('click', ()=>{
      currentBM = btn.dataset.bm;
      if(currentBM !== 'All'){
        currentRegion = 'All';
        document.querySelectorAll('#regionFilters .tab').forEach(b=>b.classList.remove('active'));
        document.querySelector('#regionFilters .tab[data-region="All"]')?.classList.add('active');
      }
      renderBmFilters();
      renderAll();
    });
  });
}

function renderKPIs(){
  const k = DATA.kpi;
  let target = k.totalTarget, achieved = k.totalAchieved, reps = k.totalReps, above = k.repsAbove100;
  if(scopeType()==='region' && currentRegion !== 'All'){
    const r = DATA.region.find(x=>x.Region===currentRegion);
    target = r.Target; achieved = r.Achieved; reps = r.Reps;
  } else if(scopeType()==='bm'){
    const r = DATA.bm.find(x=>x.BM===currentBM);
    if(r){ target = r.Target; achieved = r.Achieved; reps = r.Reps; }
  }
  const achPct = target ? achieved/target*100 : 0;
  document.getElementById('statusPill').textContent = 'OVERALL: ' + achPct.toFixed(1) + '% ATTAINED';
  document.getElementById('repCountLbl').textContent = reps + ' field reps' + (scopeValue()==='All' ? '' : ' \u00b7 ' + scopeLabel());

  const items = [
    ['Revenue Target', fmtCr(target), null],
    ['Revenue Achieved', fmtCr(achieved), (achPct>=100?'up':'down')],
    ['Achievement %', achPct.toFixed(1)+'%', (achPct>=100?'up':'down')],
    ['Total Transactions', fmtNum(k.totalTransactions), null],
    ['Units Sold', fmtNum(k.totalUnits), null],
    ['Reps Above 100%', above + ' / ' + reps, (above/reps>=0.5?'up':'down')],
  ];
  document.getElementById('kpiRow').innerHTML = items.map(([l,v,d])=>`
    <div class="kpi">
      <div class="lbl">${l}</div>
      <div class="val">${v}</div>
      ${d?`<div class="delta ${d}">${d==='up'?'\u25b2 on pace':'\u25bc behind pace'}</div>`:'<div class="delta" style="color:var(--muted-2)">\u2014</div>'}
    </div>`).join('');
}

function renderDayStrip(){
  const rows = DATA.daily;
  const useVal = d => {
    if(scopeType()==='bm'){
      const row = (DATA.dailyBM||[]).find(x=>x.DateStr===d.DateStr) || {};
      return row[currentBM] || 0;
    }
    return currentRegion==='All' ? d.Revenue : (DATA.dailyRegion.find(x=>x.DateStr===d.DateStr)||{})[currentRegion] || 0;
  };
  const max = Math.max(...rows.map(useVal));
  let target = DATA.dailyTargetLine;
  if(scopeType()==='region' && currentRegion !== 'All'){
    target = DATA.dailyTargetLine * (DATA.region.find(r=>r.Region===currentRegion).Target/DATA.kpi.totalTarget);
  } else if(scopeType()==='bm'){
    const r = DATA.bm.find(x=>x.BM===currentBM);
    if(r) target = DATA.dailyTargetLine * (r.Target/DATA.kpi.totalTarget);
  }

  document.getElementById('dayStrip').innerHTML = rows.map(d=>{
    const v = useVal(d);
    const h = Math.max(3, (v/max)*100);
    const cls = d.DateStr===selectedDate ? 'selected' : (v>=target ? 'above':'below');
    return `<div class="daybar ${cls}" style="height:${h}%" data-date="${d.DateStr}" tabindex="0" title="${fmtDate(d.DateStr)} \u2014 ${fmtShort(v)}"></div>`;
  }).join('');
  document.getElementById('dayLabels').innerHTML = rows.map(d=>`<span>${fmtDateShort(d.DateStr)}</span>`).join('');

  document.querySelectorAll('.daybar').forEach(el=>{
    el.addEventListener('click', ()=>{ selectedDate = el.dataset.date; renderDayStrip(); renderDayDetail(); });
    el.addEventListener('keydown', e=>{ if(e.key==='Enter'||e.key===' '){ selectedDate = el.dataset.date; renderDayStrip(); renderDayDetail(); }});
  });
}

function renderDayDetail(){
  const dd = DATA.dayDetail[selectedDate];
  if(!dd) return;
  const target = DATA.dailyTargetLine;
  let revForScope = dd.revenue;
  let vsTarget = null;
  if(scopeType()==='region' && currentRegion !== 'All'){
    revForScope = (dd.regionSplit.find(([n])=>n===currentRegion)||[null,0])[1];
  } else if(scopeType()==='bm'){
    const row = (DATA.dailyBM||[]).find(x=>x.DateStr===selectedDate) || {};
    revForScope = row[currentBM] || 0;
  } else {
    vsTarget = ((revForScope-target)/target*100);
  }

  const topReps = dd.topReps;
  const topCats = dd.topCats.filter(([,v])=>v>0);

  document.getElementById('dayDetail').innerHTML = `
    <div class="dd-block">
      <div class="dd-lbl">Selected Day</div>
      <div class="dd-date">${fmtDate(selectedDate)}</div>
      <div class="dd-rev">${fmtShort(revForScope)}</div>
      ${vsTarget!==null ? `<div class="dd-vs">${vsTarget>=0?'\u25b2':'\u25bc'} <b style="color:${vsTarget>=0?'var(--green)':'var(--red)'}">${Math.abs(vsTarget).toFixed(0)}%</b> vs required daily pace (${fmtShort(target)})</div>` : ''}
      <div class="dd-stats">
        <div><div class="n">${fmtNum(dd.transactions)}</div><div class="l">Transactions</div></div>
        <div><div class="n">${fmtNum(dd.units)}</div><div class="l">Units</div></div>
        <div><div class="n">${dd.activeReps}</div><div class="l">Active Reps</div></div>
      </div>
    </div>
    <div class="dd-block">
      <div class="dd-lbl">Top Performers That Day</div>
      <div class="chiplist">
        ${topReps.slice(0,5).map(([n,v],i)=>`<div class="chip"><span class="name"><span class="rankbadge">${i+1}</span>${n}</span><span class="amt">${fmtShort(v)}</span></div>`).join('')}
      </div>
    </div>
    <div class="dd-block">
      <div class="dd-lbl">Category Mix</div>
      <div class="chiplist">
        ${topCats.slice(0,5).map(([n,v])=>`<div class="chip"><span class="name">${n}</span><span class="amt">${fmtShort(v)}</span></div>`).join('')}
      </div>
    </div>
    <div class="dd-block">
      <div class="dd-lbl">Region &amp; Top Cities</div>
      <div class="chiplist">
        ${dd.regionSplit.map(([n,v])=>`<div class="chip"><span class="name" style="color:${n==='North'?'var(--north)':'var(--south)'}">${n}</span><span class="amt">${fmtShort(v)}</span></div>`).join('')}
        ${dd.topCities.slice(0,3).map(([n,v])=>`<div class="chip"><span class="name">${n}</span><span class="amt">${fmtShort(v)}</span></div>`).join('')}
      </div>
    </div>
  `;
}

const gridOpt = { grid:{color:'rgba(255,255,255,.055)', drawBorder:false}, ticks:{color:'#8592AE', font:{family:"'JetBrains Mono',monospace", size:10.5}} };

function renderTrendChart(){
  destroy('trend');
  const labels = DATA.daily.map(d=>fmtDateShort(d.DateStr));
  const target = DATA.dailyTargetLine;
  let datasets;
  if(scopeType()==='bm'){
    datasets = [
      { label:currentBM, data: (DATA.dailyBM||[]).map(d=>d[currentBM]||0), borderColor:'#7C9CFF', backgroundColor:'rgba(124,156,255,.14)', fill:true, tension:.35, pointRadius:0, borderWidth:2 },
    ];
  } else if(currentRegion==='All'){
    datasets = [
      { label:'North', data: DATA.dailyRegion.map(d=>d.North), borderColor:'#2DD4BF', backgroundColor:'rgba(45,212,191,.12)', fill:true, tension:.35, pointRadius:0, borderWidth:2 },
      { label:'South', data: DATA.dailyRegion.map(d=>d.South), borderColor:'#F5A524', backgroundColor:'rgba(245,165,36,.10)', fill:true, tension:.35, pointRadius:0, borderWidth:2 },
      { label:'Daily pace needed', data: DATA.daily.map(()=>target), borderColor:'#FF6B4A', borderDash:[5,4], pointRadius:0, borderWidth:1.5, fill:false },
    ];
  } else {
    datasets = [
      { label:currentRegion, data: DATA.dailyRegion.map(d=>d[currentRegion]), borderColor: currentRegion==='North'?'#2DD4BF':'#F5A524', backgroundColor: currentRegion==='North'?'rgba(45,212,191,.14)':'rgba(245,165,36,.12)', fill:true, tension:.35, pointRadius:0, borderWidth:2 },
    ];
  }
  charts.trend = new Chart(document.getElementById('trendChart'), {
    type:'line',
    data:{ labels, datasets },
    options:{
      responsive:true, interaction:{mode:'index', intersect:false},
      plugins:{ legend:{position:'top', labels:{boxWidth:10, usePointStyle:true, color:'#8592AE', font:{size:11.5}}},
        tooltip:{ callbacks:{ label:c=>`${c.dataset.label}: ${fmtShort(c.raw)}` } } },
      scales:{ x:gridOpt, y:{ ...gridOpt, ticks:{...gridOpt.ticks, callback:v=>fmtShort(v)} } }
    }
  });
}

function renderPaceChart(){
  destroy('pace');
  const labels = DATA.daily.map(d=>fmtDateShort(d.DateStr));
  const cum = DATA.daily.map(d=>d.CumRevenue);
  const cumTarget = DATA.daily.map((d,i)=> DATA.dailyTargetLine*(i+1));
  charts.pace = new Chart(document.getElementById('paceChart'), {
    type:'line',
    data:{ labels, datasets:[
      { label:'Cumulative Achieved', data:cum, borderColor:'#2DD4BF', backgroundColor:'rgba(45,212,191,.12)', fill:true, tension:.25, pointRadius:0, borderWidth:2.5 },
      { label:'Cumulative Target Pace', data:cumTarget, borderColor:'#FF6B4A', borderDash:[5,4], pointRadius:0, borderWidth:1.5, fill:false },
    ]},
    options:{ responsive:true, plugins:{legend:{position:'top', labels:{boxWidth:10, usePointStyle:true, color:'#8592AE', font:{size:11.5}}},
      tooltip:{callbacks:{label:c=>`${c.dataset.label}: ${fmtShort(c.raw)}`}}},
      scales:{ x:gridOpt, y:{...gridOpt, ticks:{...gridOpt.ticks, callback:v=>fmtShort(v)}} } }
  });
}

function renderRegionChart(){
  destroy('region');
  const regions = DATA.region;
  charts.region = new Chart(document.getElementById('regionChart'), {
    type:'bar',
    data:{ labels: regions.map(r=>r.Region), datasets:[
      { label:'Target', data: regions.map(r=>r.Target), backgroundColor:'rgba(255,255,255,.14)', borderRadius:6, maxBarThickness:46 },
      { label:'Achieved', data: regions.map(r=>r.Achieved), backgroundColor: regions.map(r=>r.Region==='North'?'#2DD4BF':'#F5A524'), borderRadius:6, maxBarThickness:46 },
    ]},
    options:{ responsive:true, plugins:{legend:{position:'top', labels:{boxWidth:10, usePointStyle:true, color:'#8592AE', font:{size:11.5}}},
      tooltip:{callbacks:{label:c=>`${c.dataset.label}: ${fmtShort(c.raw)}`}}},
      scales:{ x:gridOpt, y:{...gridOpt, ticks:{...gridOpt.ticks, callback:v=>fmtShort(v)}} } }
  });
}

function renderCatChart(){
  destroy('cat');
  let cats;
  if(scopeType()==='bm') cats = (DATA.categoryBM[currentBM] || []).slice(0,7);
  else cats = (currentRegion==='All' ? DATA.category : DATA.categoryRegion[currentRegion]).slice(0,7);
  const palette = ['#2DD4BF','#F5A524','#FF6B4A','#8592AE','#3ECF8E','#7C9CFF','#E879F9'];
  charts.cat = new Chart(document.getElementById('catChart'), {
    type:'doughnut',
    data:{ labels: cats.map(c=>c['Product Category']), datasets:[{ data: cats.map(c=>c.Revenue), backgroundColor:palette, borderColor:'#111827', borderWidth:2 }] },
    options:{ responsive:true, cutout:'62%', plugins:{ legend:{position:'right', labels:{boxWidth:9, color:'#8592AE', font:{size:10.5}, padding:8}},
      tooltip:{callbacks:{label:c=>`${c.label}: ${fmtShort(c.raw)}`}} } }
  });
}

function renderRegionCards(){
  document.getElementById('regionCards').innerHTML = DATA.region.map(r=>{
    const pct = (r.Achieved/r.Target*100);
    const color = r.Region==='North' ? 'var(--north)' : 'var(--south)';
    return `<div class="rcard">
      <div class="rcard-top"><div class="rcard-name"><i style="background:${color}"></i>${r.Region}</div><div class="rcard-pct" style="color:${color}">${pct.toFixed(1)}%</div></div>
      <div class="bar-outer"><div class="bar-inner" style="width:${Math.min(100,pct)}%;background:${color}"></div></div>
      <div class="rcard-meta"><span>${r.Reps} reps</span><span>${fmtShort(r.Achieved)} / ${fmtShort(r.Target)}</span></div>
    </div>`;
  }).join('');
}

function renderCatStackChart(){
  destroy('catStack');
  const cats = DATA.topCategories.slice(0,5);
  const palette = ['#2DD4BF','#F5A524','#FF6B4A','#8592AE','#3ECF8E'];
  const labels = DATA.dailyCategory.map(d=>fmtDateShort(d.DateStr));
  charts.catStack = new Chart(document.getElementById('catStackChart'), {
    type:'bar',
    data:{ labels, datasets: cats.map((c,i)=>({ label:c, data: DATA.dailyCategory.map(d=>d[c]), backgroundColor:palette[i], stack:'s' })) },
    options:{ responsive:true, plugins:{legend:{position:'top', labels:{boxWidth:9, usePointStyle:true, color:'#8592AE', font:{size:10.5}}},
      tooltip:{callbacks:{label:c=>`${c.dataset.label}: ${fmtShort(c.raw)}`}}},
      scales:{ x:{...gridOpt, stacked:true}, y:{...gridOpt, stacked:true, ticks:{...gridOpt.ticks, callback:v=>fmtShort(v)}} } }
  });
}

function renderDowChart(){
  destroy('dow');
  let src;
  if(scopeType()==='bm') src = DATA.dowBM[currentBM] || [];
  else src = currentRegion==='All' ? DATA.dow : DATA.dowRegion[currentRegion];
  charts.dow = new Chart(document.getElementById('dowChart'), {
    type:'bar',
    data:{ labels: src.map(d=>d.DOW.slice(0,3)), datasets:[{ label:'Avg-style Revenue', data: src.map(d=>d.Revenue), backgroundColor:'#7C9CFF', borderRadius:6, maxBarThickness:52 }] },
    options:{ responsive:true, plugins:{legend:{display:false}, tooltip:{callbacks:{label:c=>fmtShort(c.raw)}}},
      scales:{ x:gridOpt, y:{...gridOpt, ticks:{...gridOpt.ticks, callback:v=>fmtShort(v)}} } }
  });
}

function renderSlabChart(){
  destroy('slab');
  charts.slab = new Chart(document.getElementById('slabChart'), {
    type:'bar',
    data:{ labels: DATA.slab.labels, datasets:[{ data: DATA.slab.values, backgroundColor:['#F2495C','#F5A524','#F5A524','#3ECF8E','#2DD4BF'], borderRadius:6, maxBarThickness:52 }] },
    options:{ indexAxis:'y', responsive:true, plugins:{legend:{display:false}, tooltip:{callbacks:{label:c=>c.raw+' reps'}}},
      scales:{ x:{...gridOpt, ticks:{...gridOpt.ticks}}, y:gridOpt } }
  });
}

function renderTable(id, rows){
  const tbody = document.querySelector('#'+id+' tbody');
  tbody.innerHTML = rows.map((r,i)=>{
    const pct = r['Achievement in %']*100;
    return `<tr>
      <td class="rankcell">${i+1}</td>
      <td>${r.Name}</td>
      <td style="color:${r.Region==='North'?'var(--north)':'var(--south)'}">${r.Region}</td>
      <td class="num">${fmtShort(r['Revenue Target'])}</td>
      <td class="num">${fmtShort(r['Revenue Achived'])}</td>
      <td class="num"><span class="pct-tag ${pct>=100?'hi':'lo'}">${pct.toFixed(0)}%</span></td>
    </tr>`;
  }).join('');
}

function renderTables(){
  let top, bottom;
  if(scopeType()==='bm'){
    top = DATA.top10BM[currentBM] || [];
    bottom = DATA.bottom10BM[currentBM] || [];
  } else {
    top = currentRegion==='All' ? DATA.top10 : DATA.top10Region[currentRegion];
    bottom = currentRegion==='All' ? DATA.bottom10 : DATA.bottom10Region[currentRegion];
  }
  renderTable('topTable', top);
  renderTable('bottomTable', bottom);
}

function renderFomTable(){
  const tbody = document.querySelector('#fomTable tbody');
  tbody.innerHTML = DATA.fom.map(f=>{
    const pct = f.AchPct;
    return `<tr>
      <td>${f['Field Operations Manager']}</td>
      <td class="num">${f.Reps}</td>
      <td class="num">${fmtShort(f.Target)}</td>
      <td class="num">${fmtShort(f.Achieved)}</td>
      <td class="num"><span class="pct-tag ${pct>=100?'hi':'lo'}">${pct.toFixed(0)}%</span></td>
    </tr>`;
  }).join('');
}

document.getElementById('regionFilters')?.addEventListener('click', e=>{
  const btn = e.target.closest('.tab');
  if(!btn) return;
  document.querySelectorAll('#regionFilters .tab').forEach(b=>b.classList.remove('active'));
  btn.classList.add('active');
  currentRegion = btn.dataset.region;
  if(currentRegion !== 'All'){
    currentBM = 'All';
    renderBmFilters();
  }
  renderAll();
});

function renderAbKPIs(){
  const ab = DATA.alphaBooster;
  if(!ab) return;
  const k = ab.kpi;
  const items = [
    ["Alpha + Booster Revenue", fmtShort(k.totalRevenue), null],
    ["% of Total Revenue", k.pctOfRevenue.toFixed(1)+"%", null],
    ["Units Sold", fmtNum(k.totalUnits), null],
    ["Transactions", fmtNum(k.totalTransactions), null],
  ];
  document.getElementById("abKpiRow").innerHTML = items.map(([l,v])=>`
    <div class="kpi">
      <div class="lbl">${l}</div>
      <div class="val">${v}</div>
      <div class="delta" style="color:var(--muted-2)">\u2014</div>
    </div>`).join("");
}

function renderAbComboChart(){
  destroy("abCombo");
  const ab = DATA.alphaBooster; if(!ab || !ab.combo || !ab.combo.length) return;
  const combo = ab.combo;
  const palette = ["#2DD4BF","#F5A524","#FF6B4A","#7C9CFF"];
  charts.abCombo = new Chart(document.getElementById("abComboChart"), {
    type:"bar",
    data:{ labels: combo.map(c=>c.Label), datasets:[{ data: combo.map(c=>c.Revenue), backgroundColor: combo.map((c,i)=>palette[i%palette.length]), borderRadius:6, maxBarThickness:56 }] },
    options:{ responsive:true, plugins:{legend:{display:false}, tooltip:{callbacks:{label:c=>fmtShort(c.raw)}}},
      scales:{ x:gridOpt, y:{...gridOpt, ticks:{...gridOpt.ticks, callback:v=>fmtShort(v)}} } }
  });
}

function renderAbDailyChart(){
  destroy("abDaily");
  const ab = DATA.alphaBooster; if(!ab || !ab.daily || !ab.daily.length) return;
  const labels = ab.daily.map(d=>fmtDateShort(d.DateStr));
  charts.abDaily = new Chart(document.getElementById("abDailyChart"), {
    type:"line",
    data:{ labels, datasets:[
      { label:"Alpha", data: ab.daily.map(d=>d.Alpha||0), borderColor:"#2DD4BF", backgroundColor:"rgba(45,212,191,.12)", fill:true, tension:.3, pointRadius:0, borderWidth:2 },
      { label:"X-Factor (Booster)", data: ab.daily.map(d=>d["X-Factor"]||0), borderColor:"#F5A524", backgroundColor:"rgba(245,165,36,.10)", fill:true, tension:.3, pointRadius:0, borderWidth:2 },
    ]},
    options:{ responsive:true, plugins:{legend:{position:"top", labels:{boxWidth:10, usePointStyle:true, color:"#8592AE", font:{size:11.5}}},
      tooltip:{callbacks:{label:c=>`${c.dataset.label}: ${fmtShort(c.raw)}`}}},
      scales:{ x:gridOpt, y:{...gridOpt, ticks:{...gridOpt.ticks, callback:v=>fmtShort(v)}} } }
  });
}

function renderAbTypeChart(){
  destroy("abType");
  const ab = DATA.alphaBooster; if(!ab || !ab.typeSummary || !ab.typeSummary.length) return;
  charts.abType = new Chart(document.getElementById("abTypeChart"), {
    type:"doughnut",
    data:{ labels: ab.typeSummary.map(t=>t.Type), datasets:[{ data: ab.typeSummary.map(t=>t.Revenue), backgroundColor:["#2DD4BF","#FF6B4A"], borderColor:"#111827", borderWidth:2 }] },
    options:{ responsive:true, cutout:"62%", plugins:{ legend:{position:"right", labels:{boxWidth:9, color:"#8592AE", font:{size:10.5}, padding:8}},
      tooltip:{callbacks:{label:c=>`${c.label}: ${fmtShort(c.raw)}`}} } }
  });
}

function renderAbRegionChart(){
  destroy("abRegion");
  const ab = DATA.alphaBooster; if(!ab || !ab.regionProgram || !ab.regionProgram.length) return;
  const rp = ab.regionProgram;
  charts.abRegion = new Chart(document.getElementById("abRegionChart"), {
    type:"bar",
    data:{ labels: rp.map(r=>r.Region), datasets:[
      { label:"Alpha", data: rp.map(r=>r.Alpha||0), backgroundColor:"#2DD4BF", borderRadius:6, maxBarThickness:40 },
      { label:"X-Factor", data: rp.map(r=>r["X-Factor"]||0), backgroundColor:"#F5A524", borderRadius:6, maxBarThickness:40 },
    ]},
    options:{ responsive:true, plugins:{legend:{position:"top", labels:{boxWidth:10, usePointStyle:true, color:"#8592AE", font:{size:11.5}}},
      tooltip:{callbacks:{label:c=>`${c.dataset.label}: ${fmtShort(c.raw)}`}}},
      scales:{ x:{...gridOpt, stacked:true}, y:{...gridOpt, stacked:true, ticks:{...gridOpt.ticks, callback:v=>fmtShort(v)}} } }
  });
}

function renderAbTopRepsTable(){
  const ab = DATA.alphaBooster; if(!ab) return;
  const tbody = document.querySelector("#abTopRepsTable tbody");
  if(!tbody) return;
  tbody.innerHTML = (ab.topReps||[]).map((r,i)=>`
    <tr>
      <td class="rankcell">${i+1}</td>
      <td>${r.Name}</td>
      <td style="color:${r.Region==="North"?"var(--north)":"var(--south)"}">${r.Region}</td>
      <td class="num">${fmtShort(r.Revenue)}</td>
      <td class="num">${fmtNum(r.Units)}</td>
    </tr>`).join("");
}

function renderExecSummary(){
  const bullets = [];
  const push = (html, tone) => bullets.push({ html, tone: tone || "" });

  let target = DATA.kpi.totalTarget, achieved = DATA.kpi.totalAchieved, reps = DATA.kpi.totalReps, above = DATA.kpi.repsAbove100;
  if(scopeType()==='region' && currentRegion !== "All"){
    const r = DATA.region.find(x=>x.Region===currentRegion);
    target = r.Target; achieved = r.Achieved; reps = r.Reps;
  } else if(scopeType()==='bm'){
    const r = DATA.bm.find(x=>x.BM===currentBM);
    if(r){ target = r.Target; achieved = r.Achieved; reps = r.Reps; }
  }
  const achPct = target ? (achieved/target*100) : 0;
  const label = scopeLabel();
  push(`<b>${label}: ${achPct.toFixed(1)}% achieved</b> \u2014 ${fmtShort(achieved)} of ${fmtShort(target)} target, ${DATA.kpi.activeDays}/${DATA.kpi.daysInMonth} days in.`, achPct>=100?"pos":"neg");

  push(`<b>${above}/${reps} reps</b> at 100%+ target${scopeValue()==="All"?"":" in "+label}.`, (above/reps)>=0.5?"pos":"neg");

  if(scopeType()==='region' && currentRegion==="All" && DATA.region.length>1){
    const sorted = [...DATA.region].sort((a,b)=> (b.Achieved/b.Target) - (a.Achieved/a.Target));
    const best = sorted[0], worst = sorted[sorted.length-1];
    if(best.Region !== worst.Region){
      push(`<b>${best.Region}</b> leads at ${(best.Achieved/best.Target*100).toFixed(0)}%, <b>${worst.Region}</b> at ${(worst.Achieved/worst.Target*100).toFixed(0)}%.`);
    }
  }

  let top, bottom;
  if(scopeType()==='bm'){ top = DATA.top10BM[currentBM]; bottom = DATA.bottom10BM[currentBM]; }
  else { top = currentRegion==="All" ? DATA.top10 : DATA.top10Region[currentRegion]; bottom = currentRegion==="All" ? DATA.bottom10 : DATA.bottom10Region[currentRegion]; }
  if(top && top.length){
    const t = top[0];
    push(`Top: <b>${t.Name}</b> (${t.Region}) \u2014 ${(t["Achievement in %"]*100).toFixed(0)}%.`, "pos");
  }
  if(bottom && bottom.length){
    const b = bottom[0];
    push(`Needs attention: <b>${b.Name}</b> (${b.Region}) \u2014 ${(b["Achievement in %"]*100).toFixed(0)}%.`, "neg");
  }

  if(DATA.daily && DATA.daily.length){
    const bestDay = [...DATA.daily].sort((a,b)=>b.Revenue-a.Revenue)[0];
    push(`Best day: <b>${fmtDate(bestDay.DateStr)}</b>, ${fmtShort(bestDay.Revenue)}.`);
  }

  let cats;
  if(scopeType()==='bm') cats = DATA.categoryBM[currentBM];
  else cats = currentRegion==="All" ? DATA.category : DATA.categoryRegion[currentRegion];
  if(cats && cats.length){
    const topCat = cats[0];
    const totalRev = cats.reduce((s,c)=>s+c.Revenue,0);
    push(`Top category: <b>${topCat["Product Category"]}</b> \u2014 ${(topCat.Revenue/totalRev*100).toFixed(0)}% of revenue.`);
  }

  if(scopeValue()==="All" && DATA.alphaBooster && DATA.alphaBooster.kpi && DATA.alphaBooster.kpi.totalRevenue > 0){
    const ab = DATA.alphaBooster.kpi;
    push(`Alpha / X Factor: <b>${fmtShort(ab.totalRevenue)}</b> \u2014 ${ab.pctOfRevenue.toFixed(0)}% of revenue.`);
  }

  const daysLeft = DATA.kpi.daysInMonth - DATA.kpi.activeDays;
  if(daysLeft > 0 && achPct < 100){
    const dailyNeeded = (target - achieved) / daysLeft;
    push(`Need <b>${fmtShort(Math.max(0,dailyNeeded))}/day</b> for ${daysLeft} more day${daysLeft===1?"":"s"} to hit target.`, "neg");
  } else if(achPct >= 100){
    push(`Target met${scopeValue()==="All"?"":" in "+label}${daysLeft>0?" with "+daysLeft+" day(s) left":""}.`, "pos");
  }

  document.getElementById("execSummaryList").innerHTML = bullets.map(b=>
    `<li class="${b.tone}"><span class="dot"></span><span>${b.html}</span></li>`
  ).join("");
}

function renderAll(){
  renderKPIs();
  renderExecSummary();
  renderDayStrip();
  renderDayDetail();
  renderTrendChart();
  renderPaceChart();
  renderRegionChart();
  renderCatChart();
  renderRegionCards();
  renderCatStackChart();
  renderDowChart();
  renderSlabChart();
  renderTables();
  renderFomTable();
  renderAbKPIs();
  renderAbComboChart();
  renderAbDailyChart();
  renderAbTypeChart();
  renderAbRegionChart();
  renderAbTopRepsTable();
}

renderBmFilters();
renderAll();
'''

BM_BUILD_BLOCK = '''
    # ---- BM (Business/Branch Manager) breakdown ----
    bm_df = tgt.groupby("BM").agg(
        Target=("Revenue Target", "sum"), Achieved=("Revenue Achived", "sum"), Reps=("Denave ID", "count")
    ).reset_index()
    bm_df["AchPct"] = bm_df["Achieved"] / bm_df["Target"] * 100
    bm_df = bm_df.sort_values("Achieved", ascending=False)
    bm_list = bm_df["BM"].tolist()

    rep_bm = tgt.set_index("Denave ID")["BM"].to_dict()
    raw["BM"] = raw["Employee ID"].map(rep_bm)

    db = raw.groupby(["DateStr", "BM"])["Revenue"].sum().unstack(fill_value=0).reset_index().sort_values("DateStr")
    for b in bm_list:
        if b not in db.columns:
            db[b] = 0

    categoryBM, dowBM, top10BM, bottom10BM = {}, {}, {}, {}
    for b in bm_list:
        braw = raw[raw["BM"] == b]
        if braw.empty:
            continue
        categoryBM[b] = braw.groupby("Product Category").agg(
            Revenue=("Revenue", "sum"), Units=("Quantity", "sum"), Transactions=("RowId", "count")
        ).reset_index().sort_values("Revenue", ascending=False).to_dict("records")
        dowBM[b] = braw.groupby("DOW").agg(
            Revenue=("Revenue", "sum"), Transactions=("RowId", "count"), Units=("Quantity", "sum")
        ).reindex(DOW_ORDER).fillna(0).reset_index().to_dict("records")
        btgt = tgt[tgt["BM"] == b].sort_values("Achievement in %", ascending=False)
        top10BM[b] = btgt.head(10)[cols].to_dict("records")
        bottom10BM[b] = btgt.tail(10).sort_values("Achievement in %").to_dict("records")

'''


def main():
    if len(sys.argv) != 2:
        print("Usage: python patch_dashboard.py make_dashboard.py")
        sys.exit(1)

    path = Path(sys.argv[1])
    if not path.exists():
        print(f"File not found: {path}")
        sys.exit(1)

    src = path.read_text(encoding="utf-8")
    original = src
    changes = []

    # 1. Fix Alpha/X-Factor column names
    old_cols = '''    AB_TYPE_COL = "Alpha & Booster InK/Laser"
    AB_PROG_COL = "Alpha / Booster"'''
    new_cols = '''    AB_TYPE_COL = "Alpha & X Factor InK/Laser"
    AB_PROG_COL = "Alpha / X Factor"'''
    if old_cols in src:
        src = src.replace(old_cols, new_cols)
        changes.append("Fixed Alpha/X-Factor column names")
    elif new_cols in src:
        changes.append("Alpha/X-Factor column names already correct")
    else:
        print("WARNING: Could not find the AB_TYPE_COL/AB_PROG_COL block to patch. Check manually.")

    # 2. Insert BM build block after the tier block
    anchor = '''    tier["AchPct"] = tier["Achieved"] / tier["Target"] * 100'''
    if anchor in src and "bm_df = tgt.groupby" not in src:
        src = src.replace(anchor, anchor + "\n" + BM_BUILD_BLOCK.rstrip("\n"))
        changes.append("Inserted BM aggregation block")
    elif "bm_df = tgt.groupby" in src:
        changes.append("BM aggregation block already present")
    else:
        print("WARNING: Could not find the tier AchPct anchor line. Check manually.")

    # 3. Add BM keys to the data dict
    old_dict_tail = '''        categoryRegion=category_region, dowRegion=dow_region, top10Region=top10_region, bottom10Region=bottom10_region,
        alphaBooster=alpha_booster,
    )'''
    new_dict_tail = '''        categoryRegion=category_region, dowRegion=dow_region, top10Region=top10_region, bottom10Region=bottom10_region,
        alphaBooster=alpha_booster,
        bmList=bm_list, bm=bm_df.to_dict("records"), dailyBM=db[["DateStr"] + bm_list].to_dict("records"),
        categoryBM=categoryBM, dowBM=dowBM, top10BM=top10BM, bottom10BM=bottom10BM,
    )'''
    if old_dict_tail in src:
        src = src.replace(old_dict_tail, new_dict_tail)
        changes.append("Added BM keys to the data dict")
    elif "bmList=bm_list" in src:
        changes.append("BM keys already present in data dict")
    else:
        print("WARNING: Could not find the data dict closing block. Check manually.")

    # 4. Add bmFilters div to BODY_HTML
    old_filters = '''    <button class="tab" data-region="South">South</button>
  </div>'''
    new_filters = '''    <button class="tab" data-region="South">South</button>
  </div>

  <div class="filters" id="bmFilters"></div>'''
    if old_filters in src:
        src = src.replace(old_filters, new_filters)
        changes.append("Added bmFilters div to BODY_HTML")
    elif 'id="bmFilters"' in src:
        changes.append("bmFilters div already present")
    else:
        print("WARNING: Could not find the region filters block in BODY_HTML. Check manually.")

    # 5. Replace APP_JS = '''...''' entirely
    m = re.search(r"APP_JS = '''.*?'''\n", src, re.S)
    if m:
        src = src[:m.start()] + "APP_JS = '''" + NEW_APP_JS + "'''\n" + src[m.end():]
        changes.append("Replaced APP_JS with BM-aware version")
    else:
        print("WARNING: Could not find the APP_JS = '''...''' block (check your quote style: it must be triple single-quotes). Check manually.")

    if src == original:
        print("No changes were made -- nothing matched. Your file may already differ from the expected structure.")
        sys.exit(1)

    backup = path.with_suffix(path.suffix + ".bak")
    shutil.copy(path, backup)
    path.write_text(src, encoding="utf-8")

    print(f"Backed up original to: {backup}")
    print(f"Patched: {path}")
    print("Changes applied:")
    for c in changes:
        print(f"  - {c}")


if __name__ == "__main__":
    main()
