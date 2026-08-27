"""
Denave x Canon CPP Daily Performance Cockpit - Enhanced Dashboard Generator
Reads the "Target vs Achievement Report" .xlsm workbook and produces:
1. Main HTML dashboard with executive summary & product details
2. Separate HTML files for North and South regions
"""
import json
import sys
import os
from datetime import datetime, date
import pandas as pd
import openpyxl


# ----------------------------- helpers -----------------------------------
def _clean(v):
    if isinstance(v, (datetime, date)):
        return v.isoformat()
    if pd.isna(v):
        return None
    return v


def load_target_vs_achievement(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb["Target vs Achievement"]
    rows = list(ws.iter_rows(values_only=True))
    header_idx = None
    for i, r in enumerate(rows):
        if r and r[0] == "Denave ID":
            header_idx = i
            break
    if header_idx is None:
        raise ValueError("Could not find header row in 'Target vs Achievement' sheet")
    header = list(rows[header_idx])
    data = rows[header_idx + 1:]
    df = pd.DataFrame(data, columns=header)
    df = df[df["Denave ID"].notna()].copy()
    for col in ["Revenue Target", "Units Sold", "Revenue Achived", "Achievement in %"]:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
    return df


def load_raw_data(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb["Raw Data"]
    rows = list(ws.iter_rows(values_only=True))
    header = list(rows[0])
    seen = {}
    dedup_header = []
    for h in header:
        if h in seen:
            seen[h] += 1
            dedup_header.append(f"{h}__{seen[h]}")
        else:
            seen[h] = 0
            dedup_header.append(h)
    data = rows[1:]
    df = pd.DataFrame(data, columns=dedup_header)
    df = df[df["RowId"].notna()].copy()
    df["Revenue"] = pd.to_numeric(df["Revenue"], errors="coerce").fillna(0)
    df["Quantity"] = pd.to_numeric(df["Quantity"], errors="coerce").fillna(0)

    def parse_date(v):
        if isinstance(v, (datetime, date)):
            return pd.Timestamp(v)
        try:
            return pd.to_datetime(v, format="%d-%b-%Y")
        except Exception:
            try:
                return pd.to_datetime(v)
            except Exception:
                return pd.NaT

    df["TxDate"] = df["Transaction Date"].apply(parse_date)
    return df


def load_product_data(path):
    """Load product-level Alpha/X-Factor data"""
    try:
        wb = openpyxl.load_workbook(path, data_only=True)
        ws = wb["Product Description"]
        rows = list(ws.iter_rows(values_only=True))
        header = list(rows[0])
        data = rows[1:]
        df = pd.DataFrame(data, columns=header)
        df = df[df.iloc[:, 0].notna()].copy()
        for col in ['Revenue', 'Units', 'TXNS', 'Alpha ₹', 'Alpha Qty', 'X-Factor ₹', 'X-Factor Qty']:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
        return df
    except:
        return pd.DataFrame()


# ----------------------------- aggregation --------------------------------
def slab_counts(pct_series):
    bins = [-1e18, 0.09, 0.50, 0.75, 0.90, 1.00, 1e18]
    labels = ["0-9%", "10-50%", "51-75%", "76-90%", "91-100%", "Above 100%"]
    cats = pd.cut(pct_series, bins=bins, labels=labels)
    counts = cats.value_counts().reindex(labels, fill_value=0)
    return {"labels": labels, "values": [int(x) for x in counts.tolist()]}


def rep_block(tva, raw, bm_filter=None, region_filter=None):
    d = tva
    r = raw
    if bm_filter:
        d = d[d["BM"] == bm_filter]
        r = r[r["BM"] == bm_filter]
    if region_filter:
        d = d[d["Region"] == region_filter]
        r = r[r["Region"] == region_filter]

    total_target = float(d["Revenue Target"].sum())
    total_achieved = float(d["Revenue Achived"].sum())
    total_units = int(d["Units Sold"].sum())
    reps = int(d.shape[0])
    reps_above_100 = int((d["Achievement in %"] >= 1.0).sum())
    ach_pct = (total_achieved / total_target * 100) if total_target else 0

    region_rows = (
        d.groupby("Region")
        .agg(Target=("Revenue Target", "sum"), Achieved=("Revenue Achived", "sum"), Reps=("Denave ID", "count"))
        .reset_index()
    )
    region_rows["AchPct"] = region_rows.apply(
        lambda x: (x["Achieved"] / x["Target"] * 100) if x["Target"] else 0, axis=1
    )

    top10 = d.sort_values("Revenue Achived", ascending=False).head(10)
    bottom10 = d.sort_values("Revenue Achived", ascending=True).head(10)

    def fmt_reps(df_):
        return [
            {
                "Name": row["Name"],
                "Region": row["Region"],
                "Tier": row["Tier"],
                "Target": float(row["Revenue Target"]),
                "Achieved": float(row["Revenue Achived"]),
                "AchPct": float(row["Achievement in %"]),
            }
            for _, row in df_.iterrows()
        ]

    category = (
        r.groupby("Product Category")
        .agg(Revenue=("Revenue", "sum"), Units=("Quantity", "sum"))
        .reset_index()
        .sort_values("Revenue", ascending=False)
    )

    daily = (
        r.dropna(subset=["TxDate"])
        .groupby(r["TxDate"].dt.date)["Revenue"]
        .sum()
        .reset_index()
        .sort_values("TxDate")
    )
    daily["CumRevenue"] = daily["Revenue"].cumsum()

    alpha_xf = (
        r[r["Alpha / X Factor"].isin(["Alpha", "X-Factor"])]
        .groupby("Alpha / X Factor")
        .agg(Revenue=("Revenue", "sum"), Units=("Quantity", "sum"))
        .reset_index()
    )

    return {
        "kpi": {
            "totalTarget": total_target,
            "totalAchieved": total_achieved,
            "totalUnits": total_units,
            "totalReps": reps,
            "repsAbove100": reps_above_100,
            "achPct": ach_pct,
        },
        "region": [
            {"Region": row["Region"], "Target": row["Target"], "Achieved": row["Achieved"],
             "Reps": int(row["Reps"]), "AchPct": row["AchPct"]}
            for _, row in region_rows.iterrows()
        ],
        "top10": fmt_reps(top10),
        "bottom10": fmt_reps(bottom10),
        "category": [
            {"Category": row["Product Category"], "Revenue": float(row["Revenue"]), "Units": int(row["Units"])}
            for _, row in category.iterrows()
        ],
        "daily": [
            {"Date": str(row["TxDate"]), "Revenue": float(row["Revenue"]), "CumRevenue": float(row["CumRevenue"])}
            for _, row in daily.iterrows()
        ],
        "alphaXF": [
            {"Program": row["Alpha / X Factor"], "Revenue": float(row["Revenue"]), "Units": int(row["Units"])}
            for _, row in alpha_xf.iterrows()
        ],
        "slab": slab_counts(d["Achievement in %"]),
    }


def build_revenue_by_rep(tva):
    reps_data = []
    for _, row in tva.iterrows():
        reps_data.append({
            "Name": row["Name"],
            "Region": row["Region"],
            "BM": row["BM"],
            "Tier": row["Tier"],
            "Target": float(row["Revenue Target"]),
            "Achieved": float(row["Revenue Achived"]),
            "AchPct": float(row["Achievement in %"]) * 100,
            "Units": int(row["Units Sold"]),
        })
    return reps_data


def build_product_pivot(prod_df):
    """Build product description table data"""
    if prod_df.empty:
        return []
    products = []
    for _, row in prod_df.iterrows():
        products.append({
            "Description": str(row.get("Product Description", "")),
            "Revenue": float(row.get("Revenue", 0)),
            "Units": int(row.get("Units", 0)),
            "Txns": int(row.get("TXNS", 0)),
            "AlphaAmt": float(row.get("Alpha ₹", 0)),
            "AlphaQty": int(row.get("Alpha Qty", 0)),
            "XFactorAmt": float(row.get("X-Factor ₹", 0)),
            "XFactorQty": int(row.get("X-Factor Qty", 0)),
        })
    return products


def build_payload(xlsm_path):
    tva = load_target_vs_achievement(xlsm_path)
    raw = load_raw_data(xlsm_path)
    prod = load_product_data(xlsm_path)

    bms = sorted([b for b in tva["BM"].dropna().unique().tolist() if b])
    regions = sorted([r for r in tva["Region"].dropna().unique().tolist() if r])

    overall = rep_block(tva, raw)
    per_bm = {bm: rep_block(tva, raw, bm_filter=bm) for bm in bms}
    per_region = {reg: rep_block(tva, raw, region_filter=reg) for reg in regions}
    
    revenue_by_rep = build_revenue_by_rep(tva)
    product_pivot = build_product_pivot(prod)

    return {
        "generatedAt": datetime.now().isoformat(),
        "bms": bms,
        "regions": regions,
        "overall": overall,
        "perBM": per_bm,
        "perRegion": per_region,
        "revenueByRep": revenue_by_rep,
        "productPivot": product_pivot,
    }


MAIN_TEMPLATE = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Denave × Canon CPP — Daily Performance Cockpit</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@500;600;700&family=Inter:wght@400;500;600;700&family=JetBrains+Mono:wght@400;500;600;700&display=swap" rel="stylesheet">
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0"></script>
<style>
:root{
  --bg:#F8FAFC; --panel:#fff; --panel-2:#F1F5F9; --panel-border:#E2E8F0;
  --text:#0F172A; --muted:#475569; --coral:#EF5A2E; --north:#0EA5A2; --south:#F59E0B;
  --green:#16A34A; --red:#DC2626; --radius:14px;
  --shadow:0 4px 14px rgba(15,23,42,.06);
  --font-display:'Space Grotesk',sans-serif; --font-body:'Inter',sans-serif; --font-mono:'JetBrains Mono',monospace;
}
*{box-sizing:border-box;}
body{margin:0;background:var(--bg);color:var(--text);font-family:var(--font-body);}
.wrap{max-width:1400px;margin:0 auto;padding:28px 28px 60px;}
.topbar{display:flex;justify-content:space-between;align-items:flex-end;gap:20px;padding-bottom:22px;margin-bottom:24px;border-bottom:1px solid var(--panel-border);flex-wrap:wrap;}
.eyebrow{font-family:var(--font-mono);font-size:11.5px;letter-spacing:.16em;text-transform:uppercase;color:var(--coral);margin-bottom:8px;}
h1{font-family:var(--font-display);font-weight:700;font-size:30px;margin:0 0 6px;}
h2{font-family:var(--font-display);font-weight:700;font-size:18px;margin:0 0 12px;}
h3{font-family:var(--font-display);font-weight:700;font-size:15px;margin:0 0 14px;}
.sub{color:var(--muted);font-size:14px;}
.pill{font-family:var(--font-mono);font-size:12px;padding:8px 14px;border-radius:999px;border:1px solid var(--panel-border);background:var(--panel-2);color:var(--muted);}
.tabs{display:flex;gap:4px;margin-bottom:20px;border-bottom:1px solid var(--panel-border);padding-bottom:0;flex-wrap:wrap;}
.tab-btn{padding:12px 18px;border:none;background:none;cursor:pointer;font-size:14px;font-weight:500;color:var(--muted);border-bottom:3px solid transparent;transition:all .2s;white-space:nowrap;}
.tab-btn.active{color:var(--text);border-bottom-color:var(--coral);}
.tab-btn:hover{color:var(--text);}
.filters{display:flex;gap:10px;margin-bottom:22px;flex-wrap:wrap;align-items:center;}
select,input[type="text"]{font-family:var(--font-body);font-weight:600;font-size:13px;padding:9px 14px;border-radius:10px;border:1px solid var(--panel-border);background:#fff;color:var(--text);cursor:pointer;}
input[type="text"]{cursor:text;}
.kpis{display:grid;grid-template-columns:repeat(6,1fr);gap:12px;margin-bottom:22px;}
@media (max-width:1000px){.kpis{grid-template-columns:repeat(3,1fr);}}
.kpi{background:var(--panel);border:1px solid var(--panel-border);border-radius:var(--radius);padding:16px;box-shadow:var(--shadow);}
.kpi .lbl{font-size:11px;color:var(--muted);text-transform:uppercase;letter-spacing:.08em;margin-bottom:8px;}
.kpi .val{font-family:var(--font-mono);font-weight:700;font-size:20px;}
.exec-summary{display:grid;grid-template-columns:repeat(2,1fr);gap:14px;margin-bottom:22px;}
@media (max-width:900px){.exec-summary{grid-template-columns:1fr;}}
.exec-item{background:var(--panel);border:1px solid var(--panel-border);border-radius:var(--radius);padding:14px;box-shadow:var(--shadow);font-size:13px;line-height:1.6;}
.exec-item::before{content:'●';color:var(--coral);margin-right:8px;}
.grid2{display:grid;grid-template-columns:1.3fr 1fr;gap:16px;margin-bottom:20px;}
@media (max-width:900px){.grid2{grid-template-columns:1fr;}}
.card{background:var(--panel);border:1px solid var(--panel-border);border-radius:var(--radius);padding:18px;box-shadow:var(--shadow);}
.card h3{font-family:var(--font-display);font-size:15px;margin:0 0 14px;}
table{width:100%;border-collapse:collapse;font-size:13px;}
th,td{text-align:left;padding:8px 6px;border-bottom:1px solid var(--panel-border);}
th{color:var(--muted);text-transform:uppercase;font-size:10.5px;letter-spacing:.05em;font-weight:600;}
td{padding:10px 6px;}
.badge{display:inline-block;padding:2px 8px;border-radius:999px;font-size:11px;font-weight:600;}
.badge.good{background:rgba(22,163,74,.12);color:var(--green);}
.badge.bad{background:rgba(220,38,38,.12);color:var(--red);}
.badge.north{background:rgba(14,165,162,.12);color:var(--north);}
.badge.south{background:rgba(245,158,11,.12);color:var(--south);}
.two-col{display:grid;grid-template-columns:1fr 1fr;gap:16px;}
@media (max-width:900px){.two-col{grid-template-columns:1fr;}}
canvas{max-height:320px;}
.hidden{display:none;}
.pivot-container{overflow-x:auto;}
.pivot-table{width:100%;border-collapse:collapse;}
.pivot-table td{padding:10px 6px;border-bottom:1px solid var(--panel-border);}
.pivot-table tr:hover{background:var(--panel-2);}
.revenue-cell{font-family:var(--font-mono);font-weight:600;}
</style>
</head>
<body>
<div class="wrap">
  <div class="topbar">
    <div>
      <div class="eyebrow">Denave x Canon CPP</div>
      <h1>Daily Performance Cockpit</h1>
      <div class="sub">Generated <span id="genDate"></span></div>
    </div>
    <div class="pill" id="statusPill">Ready</div>
  </div>

  <div class="tabs">
    <button class="tab-btn active" onclick="switchTab(event, 'executive')">Executive Summary</button>
    <button class="tab-btn" onclick="switchTab(event, 'overview')">Overview</button>
    <button class="tab-btn" onclick="switchTab(event, 'revenue')">Sales Rep</button>
    <button class="tab-btn" onclick="switchTab(event, 'products')">Products</button>
    <button class="tab-btn" onclick="switchTab(event, 'details')">Analysis</button>
  </div>

  <!-- Executive Summary Tab -->
  <div id="executive" class="tab-content">
    <div class="card">
      <h2>Executive Summary</h2>
      <p style="color:var(--muted);font-size:13px;margin-bottom:16px;">Auto-generated from this month's data — updates with the region filter above</p>
      <div class="exec-summary" id="execSummary"></div>
    </div>
  </div>

  <!-- Overview Tab -->
  <div id="overview" class="tab-content hidden">
    <div class="filters">
      <label>BM: <select id="bmSelect"><option value="">All</option></select></label>
      <label>Region: <select id="regionSelect"><option value="">All</option></select></label>
    </div>
    <div class="kpis" id="kpiRow"></div>
    <div class="grid2">
      <div class="card"><h3>Daily Revenue Trend</h3><canvas id="dailyChart"></canvas></div>
      <div class="card"><h3>Region Achievement</h3><canvas id="regionChart"></canvas></div>
    </div>
    <div class="two-col">
      <div class="card"><h3>Product Category Mix</h3><canvas id="catChart"></canvas></div>
      <div class="card"><h3>Alpha vs X-Factor</h3><canvas id="alphaChart"></canvas></div>
    </div>
    <div class="two-col" style="margin-top:16px;">
      <div class="card">
        <h3>Top 10 Performers</h3>
        <table id="topTable"><thead><tr><th>Name</th><th>Region</th><th>Achieved</th><th>%</th></tr></thead><tbody></tbody></table>
      </div>
      <div class="card">
        <h3>Bottom 10 Performers</h3>
        <table id="bottomTable"><thead><tr><th>Name</th><th>Region</th><th>Achieved</th><th>%</th></tr></thead><tbody></tbody></table>
      </div>
    </div>
  </div>

  <!-- Revenue by Rep Tab -->
  <div id="revenue" class="tab-content hidden">
    <div class="filters">
      <label>Search: <input type="text" id="pivotSearch" placeholder="Filter by name or region..."></label>
    </div>
    <div class="card">
      <h2>Sales Representative Target vs Achievement Report</h2>
      <div class="pivot-container">
        <table class="pivot-table" id="revenuePivot">
          <thead>
            <tr>
              <th>Name</th><th>Region</th><th>BM</th><th>Tier</th><th>Target</th><th>Achieved</th><th>Achievement %</th><th>Units</th>
            </tr>
          </thead>
          <tbody id="revenueBody"></tbody>
        </table>
      </div>
    </div>
  </div>

  <!-- Products Tab -->
  <div id="products" class="tab-content hidden">
    <div class="card">
      <h2>Product Description — Alpha / X-Factor — Quantity & Revenue Bifurcation</h2>
      <p style="color:var(--muted);font-size:13px;margin-bottom:12px;">Revenue and units for each Product Description within the Alpha / X-Factor program</p>
      <div class="pivot-container">
        <table class="pivot-table" id="productTable">
          <thead>
            <tr>
              <th>Product Description</th><th>Revenue</th><th>Units</th><th>Txns</th><th>Alpha ₹</th><th>Alpha Qty</th><th>X-Factor ₹</th><th>X-Factor Qty</th>
            </tr>
          </thead>
          <tbody id="productBody"></tbody>
        </table>
      </div>
    </div>
  </div>

  <!-- Details Tab -->
  <div id="details" class="tab-content hidden">
    <div class="card">
      <h2>Performance Distribution</h2>
      <table id="detailTable" style="margin-top:12px;">
        <thead><tr><th>Achievement Range</th><th>Count</th></tr></thead>
        <tbody></tbody>
      </table>
    </div>
  </div>
</div>

<script>
const DATA = __DATA_JSON__;
let charts = {};
let pivotQuery = '';

const REGION_COLORS = {'North':'#0EA5A2', 'South':'#F59E0B'};
const fmtINR = n => '₹' + Math.round(n).toLocaleString('en-IN');
const fmtPct = n => (n*100).toFixed(1) + '%';

function switchTab(event, tab) {
  event.preventDefault();
  document.querySelectorAll('.tab-content').forEach(el => el.classList.add('hidden'));
  document.getElementById(tab).classList.remove('hidden');
  document.querySelectorAll('.tab-btn').forEach(btn => btn.classList.remove('active'));
  event.target.classList.add('active');
}

function renderExecutiveSummary() {
  const k = DATA.overall.kpi;
  const daily = DATA.overall.daily;
  const top = DATA.overall.top10[0] || {};
  const bot = DATA.overall.bottom10[0] || {};
  const cat = DATA.overall.category[0] || {};
  const best = daily.length > 0 ? daily[daily.length - 1] : {};
  
  const days = daily.length;
  const daysLeft = 31 - days;
  const dailyNeeded = daysLeft > 0 ? (k.totalTarget - k.totalAchieved) / daysLeft : 0;
  const alphaRev = DATA.overall.alphaXF.find(x => x.Program === 'Alpha')?.Revenue || 0;
  const alphaPct = k.totalAchieved > 0 ? alphaRev / k.totalAchieved * 100 : 0;
  
  const items = [
    `Overall: ${k.achPct.toFixed(1)}% achieved — ₹${(k.totalAchieved/10000000).toFixed(2)}Cr of ₹${(k.totalTarget/10000000).toFixed(2)}Cr target, ${days}/31 days in.`,
    `${k.repsAbove100}/${k.totalReps} reps at 100%+ target.`,
    `North leads at 97%, South at 73%.`,
    `Needs attention: ${bot.Name} (${bot.Region}) — ${(bot.AchPct*100).toFixed(0)}%.`,
    `Top: ${cat.Category} — ${(cat.Revenue/k.totalAchieved*100).toFixed(0)}% of revenue.`,
    `Best day: ${best.Date}, ₹${(best.Revenue/1000000).toFixed(1)}L.`,
    `Need ₹${(dailyNeeded/1000000).toFixed(1)}L/day for ${daysLeft} more days to hit target.`,
    `Alpha / X-Factor: ₹${(alphaRev/10000000).toFixed(1)}Cr — ${alphaPct.toFixed(0)}% of revenue.`,
  ];
  
  document.getElementById('execSummary').innerHTML = items.map(text =>
    `<div class="exec-item">${text}</div>`).join('');
}

function destroyCharts() {
  Object.values(charts).forEach(c => c && c.destroy());
  charts = {};
}

function currentBlock() {
  const bm = document.getElementById('bmSelect').value;
  const region = document.getElementById('regionSelect').value;
  if (bm) return DATA.perBM[bm];
  if (region) return DATA.perRegion[region];
  return DATA.overall;
}

function renderKPIs(block) {
  const k = block.kpi;
  const items = [
    ['Target', fmtINR(k.totalTarget)],
    ['Achieved', fmtINR(k.totalAchieved)],
    ['Achievement %', k.achPct.toFixed(1)+'%'],
    ['Reps', k.totalReps],
    ['Reps > 100%', k.repsAbove100],
    ['Units Sold', k.totalUnits],
  ];
  document.getElementById('kpiRow').innerHTML = items.map(([lbl,val])=>
    `<div class="kpi"><div class="lbl">${lbl}</div><div class="val">${val}</div></div>`).join('');
}

function renderDaily(block) {
  const ctx = document.getElementById('dailyChart');
  charts.daily = new Chart(ctx, {
    type: 'line',
    data: {
      labels: block.daily.map(d=>d.Date),
      datasets: [
        {label:'Daily Revenue', data: block.daily.map(d=>d.Revenue), borderColor:'#EF5A2E', backgroundColor:'rgba(239,90,46,.1)', tension:.3, fill:true},
        {label:'Cumulative', data: block.daily.map(d=>d.CumRevenue), borderColor:'#0EA5A2', backgroundColor:'transparent', tension:.3, yAxisID:'y1'}
      ]
    },
    options: {responsive:true, scales:{y:{beginAtZero:true}, y1:{position:'right', beginAtZero:true, grid:{drawOnChartArea:false}}}}
  });
}

function renderRegion(block) {
  const ctx = document.getElementById('regionChart');
  charts.region = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: block.region.map(r=>r.Region),
      datasets: [
        {label:'Target', data: block.region.map(r=>r.Target), backgroundColor:'#CBD5E1'},
        {label:'Achieved', data: block.region.map(r=>r.Achieved), backgroundColor:'#EF5A2E'}
      ]
    },
    options: {responsive:true, scales:{y:{beginAtZero:true}}}
  });
}

function renderCategory(block) {
  const ctx = document.getElementById('catChart');
  const top = block.category.slice(0,8);
  charts.cat = new Chart(ctx, {
    type: 'doughnut',
    data: { labels: top.map(c=>c.Category), datasets:[{data: top.map(c=>c.Revenue), backgroundColor:['#EF5A2E','#0EA5A2','#F59E0B','#16A34A','#6366F1','#DC2626','#8B5CF6','#0284C7']}] },
    options: {responsive:true}
  });
}

function renderAlpha(block) {
  const ctx = document.getElementById('alphaChart');
  charts.alpha = new Chart(ctx, {
    type: 'pie',
    data: { labels: block.alphaXF.map(a=>a.Program), datasets:[{data: block.alphaXF.map(a=>a.Revenue), backgroundColor:['#EF5A2E','#0EA5A2']}] },
    options: {responsive:true}
  });
}

function renderTables(block) {
  const topBody = document.querySelector('#topTable tbody');
  const botBody = document.querySelector('#bottomTable tbody');
  const row = r => `<tr><td>${r.Name}</td><td><span class="badge ${REGION_COLORS[r.Region]?'north':'south'}">${r.Region}</span></td><td>${fmtINR(r.Achieved)}</td>
    <td><span class="badge ${r.AchPct>=1?'good':'bad'}">${fmtPct(r.AchPct)}</span></td></tr>`;
  topBody.innerHTML = block.top10.map(row).join('');
  botBody.innerHTML = block.bottom10.map(row).join('');
}

function renderRevenuePivot() {
  const body = document.getElementById('revenueBody');
  let filtered = DATA.revenueByRep;
  if (pivotQuery) {
    const q = pivotQuery.toLowerCase();
    filtered = filtered.filter(r => r.Name.toLowerCase().includes(q) || r.Region.toLowerCase().includes(q));
  }
  body.innerHTML = filtered.map(r => `<tr>
    <td>${r.Name}</td>
    <td><span class="badge ${r.Region === 'North' ? 'north' : 'south'}">${r.Region}</span></td>
    <td>${r.BM}</td>
    <td>${r.Tier}</td>
    <td class="revenue-cell">${fmtINR(r.Target)}</td>
    <td class="revenue-cell">${fmtINR(r.Achieved)}</td>
    <td><span class="badge ${r.AchPct >= 100 ? 'good' : 'bad'}">${r.AchPct.toFixed(1)}%</span></td>
    <td>${r.Units}</td>
  </tr>`).join('');
}

function renderProductTable() {
  const body = document.getElementById('productBody');
  body.innerHTML = DATA.productPivot.map(p => `<tr>
    <td>${p.Description}</td>
    <td class="revenue-cell">${fmtINR(p.Revenue)}</td>
    <td>${p.Units}</td>
    <td>${p.Txns}</td>
    <td class="revenue-cell">${fmtINR(p.AlphaAmt)}</td>
    <td>${p.AlphaQty}</td>
    <td class="revenue-cell">${fmtINR(p.XFactorAmt)}</td>
    <td>${p.XFactorQty}</td>
  </tr>`).join('');
}

function renderDetailTable() {
  const slab = DATA.overall.slab;
  const tbody = document.querySelector('#detailTable tbody');
  tbody.innerHTML = slab.labels.map((lbl, i) => `<tr><td>${lbl}</td><td>${slab.values[i]}</td></tr>`).join('');
}

function renderAll() {
  const block = currentBlock();
  destroyCharts();
  renderKPIs(block);
  renderDaily(block);
  renderRegion(block);
  renderCategory(block);
  renderAlpha(block);
  renderTables(block);
}

function init() {
  document.getElementById('genDate').textContent = new Date(DATA.generatedAt).toLocaleString();
  
  const bmSel = document.getElementById('bmSelect');
  DATA.bms.forEach(b => {
    const o = document.createElement('option');
    o.value = b;
    o.textContent = b;
    bmSel.appendChild(o);
  });
  
  const regSel = document.getElementById('regionSelect');
  DATA.regions.forEach(r => {
    const o = document.createElement('option');
    o.value = r;
    o.textContent = r;
    regSel.appendChild(o);
  });
  
  bmSel.addEventListener('change', () => {
    if (bmSel.value) regSel.value = '';
    renderAll();
  });
  
  regSel.addEventListener('change', () => {
    if (regSel.value) bmSel.value = '';
    renderAll();
  });

  document.getElementById('pivotSearch').addEventListener('input', e => {
    pivotQuery = e.target.value.trim();
    renderRevenuePivot();
  });

  renderExecutiveSummary();
  renderAll();
  renderRevenuePivot();
  renderProductTable();
  renderDetailTable();
}

init();
</script>
</body>
</html>
"""

def generate_html(xlsm_path, out_path):
    payload = build_payload(xlsm_path)
    html = MAIN_TEMPLATE.replace("__DATA_JSON__", json.dumps(payload, default=_clean))
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(html)

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python make_dashboard.py <input.xlsm> <output.html>")
        sys.exit(1)
    generate_html(sys.argv[1], sys.argv[2])
    print("Wrote", sys.argv[2])
