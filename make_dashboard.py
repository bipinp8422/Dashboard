#!/usr/bin/env python3
"""
Denave x Canon CPP -- Daily Performance Cockpit generator
==========================================================

Turns a "Sales_Representative_Target_vs_Achievement_Report_Denave_<Month>.xlsm"
export into a single, self-contained HTML dashboard (Day Explorer, daily
trend, pacing chart, category mix, region tabs, leaderboards, etc.)

USAGE
-----
    pip install pandas openpyxl
    python make_dashboard.py "Sales_Representative_Target_vs_Achievement_Report_Denave_July.xlsm"

    # optional custom output path:
    python make_dashboard.py input.xlsm -o report.html

The script expects the workbook to contain two sheets, same as the source
reports this was built from:
  - "Raw Data"               -> transaction-level rows (Revenue, Quantity,
                                 Transaction Date, Region, Product Category, ...)
  - "Target vs Achievement"  -> one row per rep, with a 4-row header block
                                 (real headers start on row 4 of the sheet)

Requires an internet connection when the resulting HTML is *opened* (it loads
Chart.js and Google Fonts from a CDN). No internet is needed to *generate* it.
"""

import argparse
import calendar
import json
import sys
from pathlib import Path

import pandas as pd

DOW_ORDER = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
SLAB_ORDER = ["0-9%", "10-50%", "51-75%", "76-90%", "91-100%", "Above 100%"]


def load_workbook(path: Path):
    raw = pd.read_excel(path, sheet_name="Raw Data")
    tgt = pd.read_excel(path, sheet_name="Target vs Achievement", header=3)

    raw["Revenue"] = pd.to_numeric(raw["Revenue"], errors="coerce").fillna(0)
    raw["Date"] = pd.to_datetime(raw["Transaction Date"], format="%d-%b-%Y", errors="coerce")
    raw = raw.dropna(subset=["Date"])
    raw["DateStr"] = raw["Date"].dt.strftime("%Y-%m-%d")
    raw["DOW"] = raw["Date"].dt.day_name()

    tgt["Revenue Achived"] = pd.to_numeric(tgt["Revenue Achived"], errors="coerce").fillna(0)
    tgt["Revenue Target"] = pd.to_numeric(tgt["Revenue Target"], errors="coerce").fillna(0)
    return raw, tgt


def build_dataset(raw: pd.DataFrame, tgt: pd.DataFrame):
    # ---- auto-detect month/year straight from the transaction dates ----
    month_dt = raw["Date"].mode().iloc[0] if not raw.empty else pd.Timestamp.today()
    year, month_num = month_dt.year, month_dt.month
    days_in_month = calendar.monthrange(year, month_num)[1]
    month_label = f"{calendar.month_name[month_num]} {year}"

    total_target = float(tgt["Revenue Target"].sum())
    total_achieved = float(tgt["Revenue Achived"].sum())
    total_reps = int(tgt.shape[0])
    reps_above_100 = int((tgt["Achievement in %"] >= 1).sum())
    total_units = int(raw["Quantity"].sum())
    total_transactions = int(raw.shape[0])
    active_days = int(raw["DateStr"].nunique())
    daily_target = total_target / days_in_month if days_in_month else 0

    kpi = dict(
        totalTarget=total_target, totalAchieved=total_achieved, totalReps=total_reps,
        repsAbove100=reps_above_100, totalUnits=total_units, totalTransactions=total_transactions,
        newSaleTx=int((raw["Transaction type"] == "NEW SALE").sum()),
        returnTx=int((raw["Transaction type"] == "SALES RETURN").sum()),
        activeDays=active_days, daysInMonth=days_in_month,
        achPct=(total_achieved / total_target * 100 if total_target else 0),
    )

    region = tgt.groupby("Region").agg(
        Target=("Revenue Target", "sum"), Achieved=("Revenue Achived", "sum"), Reps=("Denave ID", "count")
    ).reset_index()
    region["AchPct"] = region["Achieved"] / region["Target"] * 100

    daily = raw.groupby("DateStr").agg(
        Revenue=("Revenue", "sum"), Units=("Quantity", "sum"), Transactions=("RowId", "count")
    ).reset_index().sort_values("DateStr")
    daily["CumRevenue"] = daily["Revenue"].cumsum()

    dr = raw.groupby(["DateStr", "Region"])["Revenue"].sum().unstack(fill_value=0).reset_index().sort_values("DateStr")
    for c in ["North", "South"]:
        if c not in dr.columns:
            dr[c] = 0

    topcats = raw.groupby("Product Category")["Revenue"].sum().sort_values(ascending=False).index.tolist()
    dc = raw.groupby(["DateStr", "Product Category"])["Revenue"].sum().unstack(fill_value=0).reset_index().sort_values("DateStr")
    for c in topcats:
        if c not in dc.columns:
            dc[c] = 0

    rep_tier = tgt.set_index("Denave ID")["Tier"].to_dict()
    raw["Tier"] = raw["Employee ID"].map(rep_tier)
    tiers = sorted([t for t in raw["Tier"].dropna().unique()])

    dow = raw.groupby("DOW").agg(
        Revenue=("Revenue", "sum"), Transactions=("RowId", "count"), Units=("Quantity", "sum")
    ).reindex(DOW_ORDER).fillna(0).reset_index()

    cat = raw.groupby("Product Category").agg(
        Revenue=("Revenue", "sum"), Units=("Quantity", "sum"), Transactions=("RowId", "count")
    ).reset_index().sort_values("Revenue", ascending=False)

    slab_present = [s for s in SLAB_ORDER if s in tgt["Achievement SLAB %"].unique()]
    extra_slabs = [s for s in tgt["Achievement SLAB %"].dropna().unique() if s not in slab_present]
    slab_present += extra_slabs
    slab_counts = tgt["Achievement SLAB %"].value_counts().reindex(slab_present).fillna(0).astype(int)
    slab = dict(labels=slab_present, values=slab_counts.tolist())

    tgt_sorted = tgt.sort_values("Achievement in %", ascending=False)
    cols = ["Name", "Region", "Tier", "Revenue Target", "Revenue Achived", "Achievement in %"]
    top10 = tgt_sorted.head(10)[cols].to_dict("records")
    bottom10 = tgt_sorted.tail(10).sort_values("Achievement in %")[cols].to_dict("records")

    day_detail = {}
    for d, grp in raw.groupby("DateStr"):
        by_rep = grp.groupby("Name")["Revenue"].sum().sort_values(ascending=False)
        by_cat = grp.groupby("Product Category")["Revenue"].sum().sort_values(ascending=False)
        by_region = grp.groupby("Region")["Revenue"].sum().sort_values(ascending=False)
        by_city = grp.groupby("Partner City")["Revenue"].sum().sort_values(ascending=False)
        day_detail[d] = dict(
            revenue=float(grp["Revenue"].sum()), units=int(grp["Quantity"].sum()),
            transactions=int(grp.shape[0]), activeReps=int(grp["Employee ID"].nunique()),
            topReps=[[n, float(v)] for n, v in by_rep.head(5).items()],
            topCats=[[n, float(v)] for n, v in by_cat.head(6).items()],
            regionSplit=[[n, float(v)] for n, v in by_region.items()],
            topCities=[[n, float(v)] for n, v in by_city.head(5).items()],
        )

    fom = tgt.groupby("Field Operations Manager").agg(
        Target=("Revenue Target", "sum"), Achieved=("Revenue Achived", "sum"), Reps=("Denave ID", "count")
    ).reset_index()
    fom["AchPct"] = fom["Achieved"] / fom["Target"] * 100

    tier = tgt.groupby("Tier").agg(
        Target=("Revenue Target", "sum"), Achieved=("Revenue Achived", "sum"), Reps=("Denave ID", "count")
    ).reset_index()
    tier["AchPct"] = tier["Achieved"] / tier["Target"] * 100

    category_region, dow_region, top10_region, bottom10_region = {}, {}, {}, {}
    for r in ["North", "South"]:
        rraw = raw[raw["Region"] == r]
        if rraw.empty:
            continue
        category_region[r] = rraw.groupby("Product Category").agg(
            Revenue=("Revenue", "sum"), Units=("Quantity", "sum"), Transactions=("RowId", "count")
        ).reset_index().sort_values("Revenue", ascending=False).to_dict("records")
        dow_region[r] = rraw.groupby("DOW").agg(
            Revenue=("Revenue", "sum"), Transactions=("RowId", "count"), Units=("Quantity", "sum")
        ).reindex(DOW_ORDER).fillna(0).reset_index().to_dict("records")
        rtgt = tgt[tgt["Region"] == r].sort_values("Achievement in %", ascending=False)
        top10_region[r] = rtgt.head(10)[cols].to_dict("records")
        bottom10_region[r] = rtgt.tail(10).sort_values("Achievement in %")[cols].to_dict("records")

    data = dict(
        kpi=kpi, region=region.to_dict("records"), daily=daily.to_dict("records"),
        dailyRegion=dr.to_dict("records"), dailyCategory=dc[["DateStr"] + topcats].to_dict("records"),
        topCategories=topcats, dailyTier=raw.groupby(["DateStr", "Tier"])["Revenue"].sum().unstack(fill_value=0).reset_index().sort_values("DateStr").to_dict("records"),
        tiers=tiers, dow=dow.to_dict("records"), category=cat.to_dict("records"), slab=slab,
        top10=top10, bottom10=bottom10, dayDetail=day_detail, fom=fom.sort_values("Achieved", ascending=False).to_dict("records"),
        tier=tier.to_dict("records"), dailyTargetLine=daily_target,
        categoryRegion=category_region, dowRegion=dow_region, top10Region=top10_region, bottom10Region=bottom10_region,
    )

    def fmt_short(n):
        if n >= 10_000_000:
            return f"Rs {n/10_000_000:.2f} Cr"
        if n >= 100_000:
            return f"Rs {n/100_000:.1f} L"
        return f"Rs {n:,.0f}"

    meta = dict(
        month_label=month_label, rep_count=total_reps, days_active=active_days,
        days_in_month=days_in_month, txn_count=f"{total_transactions:,}",
        daily_pace=fmt_short(daily_target), source_file="",
    )
    return data, meta


HEAD_TOP = '<!DOCTYPE html>\n<html lang="en">\n<head>\n<meta charset="UTF-8">\n<meta name="viewport" content="width=device-width, initial-scale=1.0">\n<title>Denave x Canon CPP -- Daily Performance Cockpit -- {{MONTH_LABEL}}</title>\n<link rel="preconnect" href="https://fonts.googleapis.com">\n<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>\n<link href="https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@500;600;700&family=Inter:wght@400;500;600;700&family=JetBrains+Mono:wght@400;500;600;700&display=swap" rel="stylesheet">\n<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.4/dist/chart.umd.min.js"></script>\n'

STYLE_CSS = '<style>\n:root{\n  --bg:#0A0E17;\n  --bg-grad:#0D1424;\n  --panel:#111827;\n  --panel-2:#141C30;\n  --panel-border:rgba(255,255,255,.07);\n  --panel-border-hi:rgba(255,255,255,.14);\n  --text:#EAF0FA;\n  --muted:#8592AE;\n  --muted-2:#5C6883;\n  --north:#2DD4BF;\n  --south:#F5A524;\n  --coral:#FF6B4A;\n  --coral-dim:rgba(255,107,74,.16);\n  --green:#3ECF8E;\n  --red:#F2495C;\n  --grid-line:rgba(255,255,255,.055);\n  --radius:14px;\n  --font-display:\'Space Grotesk\',sans-serif;\n  --font-body:\'Inter\',sans-serif;\n  --font-mono:\'JetBrains Mono\',monospace;\n}\n*{box-sizing:border-box;}\nhtml{-webkit-text-size-adjust:100%;}\nbody{\n  margin:0;\n  background:\n    radial-gradient(1100px 620px at 12% -8%, rgba(45,212,191,.10), transparent 60%),\n    radial-gradient(900px 560px at 100% 0%, rgba(255,107,74,.08), transparent 55%),\n    linear-gradient(180deg,var(--bg-grad),var(--bg) 340px);\n  color:var(--text);\n  font-family:var(--font-body);\n  min-height:100vh;\n  line-height:1.45;\n}\n::selection{background:var(--coral);color:#0A0E17;}\n.wrap{max-width:1360px;margin:0 auto;padding:28px 28px 60px;}\na{color:inherit;}\n\n/* Focus visibility */\nbutton:focus-visible, .daybar:focus-visible, .tab:focus-visible{\n  outline:2px solid var(--coral); outline-offset:2px;\n}\n\n/* ---------- Topbar ---------- */\n.topbar{\n  display:flex;justify-content:space-between;align-items:flex-end;gap:20px;\n  padding-bottom:22px;margin-bottom:24px;border-bottom:1px solid var(--panel-border);\n  flex-wrap:wrap;\n}\n.eyebrow{\n  font-family:var(--font-mono);font-size:11.5px;letter-spacing:.16em;text-transform:uppercase;\n  color:var(--coral);margin-bottom:8px;display:flex;align-items:center;gap:8px;\n}\n.eyebrow::before{content:"";width:7px;height:7px;border-radius:50%;background:var(--coral);box-shadow:0 0 0 4px var(--coral-dim);}\nh1{\n  font-family:var(--font-display);font-weight:700;font-size:32px;line-height:1.1;margin:0 0 6px;\n  letter-spacing:-.01em;\n}\n.sub{color:var(--muted);font-size:14px;}\n.sub b{color:var(--text);font-weight:600;}\n.topmeta{display:flex;gap:10px;flex-wrap:wrap;justify-content:flex-end;}\n.pill{\n  font-family:var(--font-mono);font-size:12px;padding:8px 14px;border-radius:999px;\n  border:1px solid var(--panel-border-hi);background:var(--panel);color:var(--muted);white-space:nowrap;\n}\n.pill.hi{\n  background:linear-gradient(135deg, rgba(45,212,191,.18), rgba(255,107,74,.14));\n  border-color:rgba(255,255,255,.18);color:var(--text);font-weight:600;\n}\n\n/* ---------- Region filter tabs ---------- */\n.filters{display:flex;gap:8px;margin-bottom:22px;}\n.tab{\n  font-family:var(--font-body);font-weight:600;font-size:13px;padding:9px 18px;border-radius:10px;\n  border:1px solid var(--panel-border);background:var(--panel);color:var(--muted);cursor:pointer;\n  transition:all .15s ease;\n}\n.tab:hover{color:var(--text);border-color:var(--panel-border-hi);}\n.tab.active{background:var(--text);color:#0A0E17;border-color:var(--text);}\n\n/* ---------- KPI row ---------- */\n.kpis{display:grid;grid-template-columns:repeat(6,1fr);gap:12px;margin-bottom:22px;}\n.kpi{\n  background:var(--panel);border:1px solid var(--panel-border);border-radius:var(--radius);\n  padding:16px 16px 14px;position:relative;overflow:hidden;\n}\n.kpi::after{\n  content:"";position:absolute;inset:auto 0 0 0;height:2px;\n  background:linear-gradient(90deg,var(--north),var(--coral));opacity:.7;\n}\n.kpi .lbl{font-size:11px;color:var(--muted);text-transform:uppercase;letter-spacing:.08em;margin-bottom:8px;}\n.kpi .val{font-family:var(--font-mono);font-weight:700;font-size:21px;letter-spacing:-.01em;}\n.kpi .delta{font-size:11.5px;margin-top:6px;font-weight:600;}\n.kpi .delta.up{color:var(--green);}\n.kpi .delta.down{color:var(--red);}\n\n/* ---------- Day Explorer (signature element) ---------- */\n.explorer{\n  background:var(--panel);border:1px solid var(--panel-border);border-radius:var(--radius);\n  padding:20px 22px 18px;margin-bottom:22px;\n}\n.explorer-head{display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:16px;flex-wrap:wrap;gap:10px;}\n.explorer-head h3{font-family:var(--font-display);font-size:17px;margin:0 0 4px;}\n.explorer-head .cap{color:var(--muted);font-size:12.5px;}\n.legend{display:flex;gap:16px;font-size:11.5px;color:var(--muted);align-items:center;}\n.legend span{display:flex;align-items:center;gap:6px;}\n.legend i{width:9px;height:9px;border-radius:2px;display:inline-block;}\n\n.daystrip{\n  display:flex;align-items:flex-end;gap:3px;height:120px;padding:0 2px 2px;\n  border-bottom:1px solid var(--grid-line);margin-bottom:8px;\n}\n.daybar{\n  flex:1;min-width:6px;border-radius:3px 3px 0 0;cursor:pointer;position:relative;\n  background:var(--muted-2);opacity:.55;transition:opacity .12s ease, transform .12s ease;\n}\n.daybar:hover{opacity:.9;}\n.daybar.above{background:var(--north);}\n.daybar.below{background:var(--south);}\n.daybar.selected{opacity:1;background:var(--coral);transform:scaleY(1.02);}\n.daylabels{display:flex;gap:3px;padding:0 2px;}\n.daylabels span{\n  flex:1;min-width:6px;text-align:center;font-family:var(--font-mono);font-size:9px;color:var(--muted-2);\n}\n\n.day-detail{\n  margin-top:18px;padding-top:18px;border-top:1px dashed var(--panel-border);\n  display:grid;grid-template-columns:1.1fr 1fr 1fr 1.2fr;gap:22px;\n}\n.dd-block .dd-lbl{font-size:10.5px;text-transform:uppercase;letter-spacing:.08em;color:var(--muted);margin-bottom:8px;}\n.dd-date{font-family:var(--font-display);font-size:20px;font-weight:700;}\n.dd-rev{font-family:var(--font-mono);font-size:22px;font-weight:700;color:var(--coral);margin-top:4px;}\n.dd-vs{font-size:12px;color:var(--muted);margin-top:4px;}\n.dd-vs b{font-weight:700;}\n.dd-stats{display:flex;gap:18px;margin-top:10px;}\n.dd-stats div{font-family:var(--font-mono);}\n.dd-stats .n{font-size:15px;font-weight:700;}\n.dd-stats .l{font-size:10px;color:var(--muted);text-transform:uppercase;}\n.chiplist{display:flex;flex-direction:column;gap:7px;}\n.chip{\n  display:flex;justify-content:space-between;align-items:center;font-size:12.5px;\n  background:var(--panel-2);border:1px solid var(--panel-border);border-radius:8px;padding:7px 10px;\n}\n.chip .name{color:var(--text);font-weight:500;}\n.chip .amt{font-family:var(--font-mono);color:var(--muted);font-size:11.5px;}\n.rankbadge{\n  display:inline-flex;align-items:center;justify-content:center;width:16px;height:16px;border-radius:4px;\n  background:var(--coral-dim);color:var(--coral);font-family:var(--font-mono);font-size:9.5px;font-weight:700;margin-right:8px;\n}\n\n/* ---------- Card grids ---------- */\n.grid{display:grid;grid-template-columns:1.4fr 1fr;gap:16px;margin-bottom:16px;}\n.grid3{display:grid;grid-template-columns:1fr 1fr 1fr;gap:16px;margin-bottom:16px;}\n.card{\n  background:var(--panel);border:1px solid var(--panel-border);border-radius:var(--radius);padding:18px 20px 14px;\n}\n.card h3{font-family:var(--font-display);font-size:15.5px;margin:0 0 3px;font-weight:600;}\n.card .cap{color:var(--muted);font-size:12px;margin-bottom:12px;}\ncanvas{max-width:100%;}\n\n/* ---------- Region snapshot mini cards ---------- */\n.regionCards{display:flex;flex-direction:column;gap:10px;}\n.rcard{background:var(--panel-2);border:1px solid var(--panel-border);border-radius:10px;padding:12px 14px;}\n.rcard-top{display:flex;justify-content:space-between;align-items:center;margin-bottom:8px;}\n.rcard-name{font-weight:700;font-size:13.5px;display:flex;align-items:center;gap:7px;}\n.rcard-name i{width:9px;height:9px;border-radius:50%;display:inline-block;}\n.rcard-pct{font-family:var(--font-mono);font-weight:700;font-size:13px;}\n.bar-outer{height:6px;background:rgba(255,255,255,.08);border-radius:99px;overflow:hidden;margin-bottom:8px;}\n.bar-inner{height:100%;border-radius:99px;}\n.rcard-meta{display:flex;justify-content:space-between;font-size:11px;color:var(--muted);font-family:var(--font-mono);}\n\n/* ---------- Tables ---------- */\n.two-tables{display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-bottom:16px;}\ntable{width:100%;border-collapse:collapse;font-size:12.8px;}\nthead th{\n  text-align:left;color:var(--muted);font-weight:600;font-size:10.5px;text-transform:uppercase;\n  letter-spacing:.06em;padding:0 8px 8px;border-bottom:1px solid var(--panel-border);\n}\ntbody td{padding:9px 8px;border-bottom:1px solid var(--grid-line);}\ntbody tr:last-child td{border-bottom:none;}\ntbody tr:hover{background:rgba(255,255,255,.025);}\ntd.num, th.num{font-family:var(--font-mono);text-align:right;}\n.rankcell{font-family:var(--font-mono);color:var(--muted);}\n.pct-tag{\n  font-family:var(--font-mono);font-size:11px;font-weight:700;padding:3px 8px;border-radius:6px;display:inline-block;\n}\n.pct-tag.hi{background:rgba(62,207,142,.15);color:var(--green);}\n.pct-tag.lo{background:rgba(242,73,92,.15);color:var(--red);}\n\n/* ---------- FOM / Tier table ---------- */\n.fom-table td, .fom-table th{font-size:12.5px;}\n\nfooter{\n  text-align:center;color:var(--muted-2);font-size:11.5px;padding-top:18px;font-family:var(--font-mono);\n}\n\n@media (max-width:1100px){\n  .kpis{grid-template-columns:repeat(3,1fr);}\n  .grid{grid-template-columns:1fr;}\n  .grid3{grid-template-columns:1fr;}\n  .two-tables{grid-template-columns:1fr;}\n  .day-detail{grid-template-columns:1fr 1fr;}\n}\n@media (max-width:640px){\n  .kpis{grid-template-columns:repeat(2,1fr);}\n  .day-detail{grid-template-columns:1fr;}\n  h1{font-size:24px;}\n}\n</style>\n'

BODY_HTML = '<body>\n<div class="wrap">\n\n  <div class="topbar">\n    <div>\n      <div class="eyebrow">Denave × Canon CPP Program</div>\n      <h1>Daily Performance Cockpit</h1>\n      <div class="sub">Sales Representative Target vs Achievement &nbsp;•&nbsp; <b>{{MONTH_LABEL}}</b> &nbsp;•&nbsp; <span id="repCountLbl">{{REP_COUNT}} field reps</span></div>\n    </div>\n    <div class="topmeta">\n      <div class="pill hi" id="statusPill">LOADING…</div>\n      <div class="pill">{{DAYS_ACTIVE}} of {{DAYS_IN_MONTH}} days active</div>\n      <div class="pill">{{TXN_COUNT}} transactions</div>\n    </div>\n  </div>\n\n  <div class="filters" id="regionFilters">\n    <button class="tab active" data-region="All">All Regions</button>\n    <button class="tab" data-region="North">North</button>\n    <button class="tab" data-region="South">South</button>\n  </div>\n\n  <div class="kpis" id="kpiRow"></div>\n\n  <!-- Signature: Day Explorer -->\n  <div class="explorer">\n    <div class="explorer-head">\n      <div>\n        <h3>Day Explorer — click any day for a full breakdown</h3>\n        <div class="cap">Bar height = revenue booked that day. Color shows performance vs the flat daily pace needed ({{DAILY_PACE}}/day) to hit the monthly target.</div>\n      </div>\n      <div class="legend">\n        <span><i style="background:var(--north)"></i>Above daily pace</span>\n        <span><i style="background:var(--south)"></i>Below daily pace</span>\n        <span><i style="background:var(--coral)"></i>Selected</span>\n      </div>\n    </div>\n    <div class="daystrip" id="dayStrip"></div>\n    <div class="daylabels" id="dayLabels"></div>\n\n    <div class="day-detail" id="dayDetail"></div>\n  </div>\n\n  <div class="grid">\n    <div class="card">\n      <h3>Daily Revenue Trend</h3>\n      <div class="cap">Booked revenue by day, split by region, vs required daily pace</div>\n      <canvas id="trendChart" height="260"></canvas>\n    </div>\n    <div class="card">\n      <h3>Pacing: Cumulative Achieved vs Target</h3>\n      <div class="cap">Running total across the month</div>\n      <canvas id="paceChart" height="260"></canvas>\n    </div>\n  </div>\n\n  <div class="grid3">\n    <div class="card">\n      <h3>Target vs Achieved by Region</h3>\n      <div class="cap">Revenue in ₹, with attainment %</div>\n      <canvas id="regionChart" height="210"></canvas>\n    </div>\n    <div class="card">\n      <h3>Revenue by Product Category</h3>\n      <div class="cap">Share of total booked revenue</div>\n      <canvas id="catChart" height="210"></canvas>\n    </div>\n    <div class="card">\n      <h3>Region Snapshot</h3>\n      <div class="cap">Attainment &amp; headcount</div>\n      <div class="regionCards" id="regionCards"></div>\n    </div>\n  </div>\n\n  <div class="grid">\n    <div class="card">\n      <h3>Revenue by Product Category — Daily Mix</h3>\n      <div class="cap">Stacked daily revenue across top categories</div>\n      <canvas id="catStackChart" height="230"></canvas>\n    </div>\n    <div class="card">\n      <h3>Average Revenue by Day of Week</h3>\n      <div class="cap">Where the calendar naturally peaks</div>\n      <canvas id="dowChart" height="230"></canvas>\n    </div>\n  </div>\n\n  <div class="two-tables">\n    <div class="card">\n      <h3>Top 10 Performers</h3>\n      <div class="cap">Ranked by achievement %</div>\n      <table id="topTable"><thead><tr><th>#</th><th>Rep</th><th>Region</th><th class="num">Target</th><th class="num">Achieved</th><th class="num">Ach %</th></tr></thead><tbody></tbody></table>\n    </div>\n    <div class="card">\n      <h3>Needs Attention (Bottom 10)</h3>\n      <div class="cap">Lowest achievement %</div>\n      <table id="bottomTable"><thead><tr><th>#</th><th>Rep</th><th>Region</th><th class="num">Target</th><th class="num">Achieved</th><th class="num">Ach %</th></tr></thead><tbody></tbody></table>\n    </div>\n  </div>\n\n  <div class="grid">\n    <div class="card">\n      <h3>Field Operations Manager — Team Rollup</h3>\n      <div class="cap">Target vs achieved by FOM team</div>\n      <table id="fomTable" class="fom-table"><thead><tr><th>FOM</th><th class="num">Reps</th><th class="num">Target</th><th class="num">Achieved</th><th class="num">Ach %</th></tr></thead><tbody></tbody></table>\n    </div>\n    <div class="card">\n      <h3>Achievement Distribution</h3>\n      <div class="cap">Reps by performance slab</div>\n      <canvas id="slabChart" height="230"></canvas>\n    </div>\n  </div>\n\n  <footer>Source: Raw Data ({{TXN_COUNT}} line items) + Target vs Achievement sheets — {{SOURCE_FILE}} &nbsp;•&nbsp; Generated dashboard</footer>\n</div>\n</body>\n'

APP_JS = '\nChart.defaults.font.family = "\'Inter\',sans-serif";\nChart.defaults.color = \'#8592AE\';\nChart.defaults.borderColor = \'rgba(255,255,255,.06)\';\n\nconst fmtCr = n => \'₹\' + (n/10000000).toFixed(2) + \' Cr\';\nconst fmtL = n => \'₹\' + (n/100000).toFixed(1) + \' L\';\nconst fmtShort = n => n>=10000000 ? \'₹\'+(n/10000000).toFixed(1)+\'Cr\' : n>=100000 ? \'₹\'+(n/100000).toFixed(1)+\'L\' : \'₹\'+Math.round(n).toLocaleString(\'en-IN\');\nconst fmtNum = n => Math.round(n).toLocaleString(\'en-IN\');\nconst fmtDate = s => { const d=new Date(s+\'T00:00:00\'); return d.toLocaleDateString(\'en-IN\',{day:\'2-digit\',month:\'short\',weekday:\'short\'}); };\nconst fmtDateShort = s => { const d=new Date(s+\'T00:00:00\'); return d.getDate(); };\n\nlet currentRegion = \'All\';\nlet selectedDate = DATA.daily[DATA.daily.length-1].DateStr;\nconst charts = {};\n\n// ---------------- KPI ROW ----------------\nfunction renderKPIs(){\n  const k = DATA.kpi;\n  let target = k.totalTarget, achieved = k.totalAchieved, reps = k.totalReps, above = k.repsAbove100;\n  if(currentRegion !== \'All\'){\n    const r = DATA.region.find(x=>x.Region===currentRegion);\n    target = r.Target; achieved = r.Achieved; reps = r.Reps;\n  }\n  const achPct = achieved/target*100;\n  document.getElementById(\'statusPill\').textContent = \'OVERALL: \' + achPct.toFixed(1) + \'% ATTAINED\';\n  document.getElementById(\'repCountLbl\').textContent = reps + \' field reps\' + (currentRegion===\'All\' ? \'\' : \' · \' + currentRegion);\n\n  const items = [\n    [\'Revenue Target\', fmtCr(target), null],\n    [\'Revenue Achieved\', fmtCr(achieved), (achPct>=100?\'up\':\'down\')],\n    [\'Achievement %\', achPct.toFixed(1)+\'%\', (achPct>=100?\'up\':\'down\')],\n    [\'Total Transactions\', fmtNum(k.totalTransactions), null],\n    [\'Units Sold\', fmtNum(k.totalUnits), null],\n    [\'Reps Above 100%\', above + \' / \' + reps, (above/reps>=0.5?\'up\':\'down\')],\n  ];\n  document.getElementById(\'kpiRow\').innerHTML = items.map(([l,v,d])=>`\n    <div class="kpi">\n      <div class="lbl">${l}</div>\n      <div class="val">${v}</div>\n      ${d?`<div class="delta ${d}">${d===\'up\'?\'▲ on pace\':\'▼ behind pace\'}</div>`:\'<div class="delta" style="color:var(--muted-2)">—</div>\'}\n    </div>`).join(\'\');\n}\n\n// ---------------- DAY EXPLORER ----------------\nfunction renderDayStrip(){\n  const rows = DATA.daily;\n  const useRegionVal = d => currentRegion===\'All\' ? d.Revenue : (DATA.dailyRegion.find(x=>x.DateStr===d.DateStr)||{})[currentRegion] || 0;\n  const max = Math.max(...rows.map(useRegionVal));\n  const target = currentRegion===\'All\' ? DATA.dailyTargetLine : DATA.dailyTargetLine * (DATA.region.find(r=>r.Region===currentRegion).Target/DATA.kpi.totalTarget);\n\n  document.getElementById(\'dayStrip\').innerHTML = rows.map(d=>{\n    const v = useRegionVal(d);\n    const h = Math.max(3, (v/max)*100);\n    const cls = d.DateStr===selectedDate ? \'selected\' : (v>=target ? \'above\':\'below\');\n    return `<div class="daybar ${cls}" style="height:${h}%" data-date="${d.DateStr}" tabindex="0" title="${fmtDate(d.DateStr)} — ${fmtShort(v)}"></div>`;\n  }).join(\'\');\n  document.getElementById(\'dayLabels\').innerHTML = rows.map(d=>`<span>${fmtDateShort(d.DateStr)}</span>`).join(\'\');\n\n  document.querySelectorAll(\'.daybar\').forEach(el=>{\n    el.addEventListener(\'click\', ()=>{ selectedDate = el.dataset.date; renderDayStrip(); renderDayDetail(); });\n    el.addEventListener(\'keydown\', e=>{ if(e.key===\'Enter\'||e.key===\' \'){ selectedDate = el.dataset.date; renderDayStrip(); renderDayDetail(); }});\n  });\n}\n\nfunction renderDayDetail(){\n  const dd = DATA.dayDetail[selectedDate];\n  if(!dd) return;\n  const target = DATA.dailyTargetLine;\n  const revForRegion = currentRegion===\'All\' ? dd.revenue : (dd.regionSplit.find(([n])=>n===currentRegion)||[null,0])[1];\n  const vsTarget = currentRegion===\'All\' ? ((revForRegion-target)/target*100) : null;\n\n  const topReps = currentRegion===\'All\' ? dd.topReps : dd.topReps; // rep-level region split not pre-split; show global\n  const topCats = dd.topCats.filter(([,v])=>v>0);\n\n  document.getElementById(\'dayDetail\').innerHTML = `\n    <div class="dd-block">\n      <div class="dd-lbl">Selected Day</div>\n      <div class="dd-date">${fmtDate(selectedDate)}</div>\n      <div class="dd-rev">${fmtShort(revForRegion)}</div>\n      ${vsTarget!==null ? `<div class="dd-vs">${vsTarget>=0?\'▲\':\'▼\'} <b style="color:${vsTarget>=0?\'var(--green)\':\'var(--red)\'}">${Math.abs(vsTarget).toFixed(0)}%</b> vs required daily pace (${fmtShort(target)})</div>` : \'\'}\n      <div class="dd-stats">\n        <div><div class="n">${fmtNum(dd.transactions)}</div><div class="l">Transactions</div></div>\n        <div><div class="n">${fmtNum(dd.units)}</div><div class="l">Units</div></div>\n        <div><div class="n">${dd.activeReps}</div><div class="l">Active Reps</div></div>\n      </div>\n    </div>\n    <div class="dd-block">\n      <div class="dd-lbl">Top Performers That Day</div>\n      <div class="chiplist">\n        ${topReps.slice(0,5).map(([n,v],i)=>`<div class="chip"><span class="name"><span class="rankbadge">${i+1}</span>${n}</span><span class="amt">${fmtShort(v)}</span></div>`).join(\'\')}\n      </div>\n    </div>\n    <div class="dd-block">\n      <div class="dd-lbl">Category Mix</div>\n      <div class="chiplist">\n        ${topCats.slice(0,5).map(([n,v])=>`<div class="chip"><span class="name">${n}</span><span class="amt">${fmtShort(v)}</span></div>`).join(\'\')}\n      </div>\n    </div>\n    <div class="dd-block">\n      <div class="dd-lbl">Region &amp; Top Cities</div>\n      <div class="chiplist">\n        ${dd.regionSplit.map(([n,v])=>`<div class="chip"><span class="name" style="color:${n===\'North\'?\'var(--north)\':\'var(--south)\'}">${n}</span><span class="amt">${fmtShort(v)}</span></div>`).join(\'\')}\n        ${dd.topCities.slice(0,3).map(([n,v])=>`<div class="chip"><span class="name">${n}</span><span class="amt">${fmtShort(v)}</span></div>`).join(\'\')}\n      </div>\n    </div>\n  `;\n}\n\n// ---------------- CHARTS ----------------\nconst gridOpt = { grid:{color:\'rgba(255,255,255,.055)\', drawBorder:false}, ticks:{color:\'#8592AE\', font:{family:"\'JetBrains Mono\',monospace", size:10.5}} };\n\nfunction destroy(id){ if(charts[id]){ charts[id].destroy(); } }\n\nfunction renderTrendChart(){\n  destroy(\'trend\');\n  const labels = DATA.daily.map(d=>fmtDateShort(d.DateStr));\n  const target = DATA.dailyTargetLine;\n  let datasets;\n  if(currentRegion===\'All\'){\n    datasets = [\n      { label:\'North\', data: DATA.dailyRegion.map(d=>d.North), borderColor:\'#2DD4BF\', backgroundColor:\'rgba(45,212,191,.12)\', fill:true, tension:.35, pointRadius:0, borderWidth:2 },\n      { label:\'South\', data: DATA.dailyRegion.map(d=>d.South), borderColor:\'#F5A524\', backgroundColor:\'rgba(245,165,36,.10)\', fill:true, tension:.35, pointRadius:0, borderWidth:2 },\n      { label:\'Daily pace needed\', data: DATA.daily.map(()=>target), borderColor:\'#FF6B4A\', borderDash:[5,4], pointRadius:0, borderWidth:1.5, fill:false },\n    ];\n  } else {\n    datasets = [\n      { label:currentRegion, data: DATA.dailyRegion.map(d=>d[currentRegion]), borderColor: currentRegion===\'North\'?\'#2DD4BF\':\'#F5A524\', backgroundColor: currentRegion===\'North\'?\'rgba(45,212,191,.14)\':\'rgba(245,165,36,.12)\', fill:true, tension:.35, pointRadius:0, borderWidth:2 },\n    ];\n  }\n  charts.trend = new Chart(document.getElementById(\'trendChart\'), {\n    type:\'line\',\n    data:{ labels, datasets },\n    options:{\n      responsive:true, interaction:{mode:\'index\', intersect:false},\n      plugins:{ legend:{position:\'top\', labels:{boxWidth:10, usePointStyle:true, color:\'#8592AE\', font:{size:11.5}}},\n        tooltip:{ callbacks:{ label:c=>`${c.dataset.label}: ${fmtShort(c.raw)}` } } },\n      scales:{ x:gridOpt, y:{ ...gridOpt, ticks:{...gridOpt.ticks, callback:v=>fmtShort(v)} } }\n    }\n  });\n}\n\nfunction renderPaceChart(){\n  destroy(\'pace\');\n  const labels = DATA.daily.map(d=>fmtDateShort(d.DateStr));\n  const cum = DATA.daily.map(d=>d.CumRevenue);\n  const cumTarget = DATA.daily.map((d,i)=> DATA.dailyTargetLine*(i+1));\n  charts.pace = new Chart(document.getElementById(\'paceChart\'), {\n    type:\'line\',\n    data:{ labels, datasets:[\n      { label:\'Cumulative Achieved\', data:cum, borderColor:\'#2DD4BF\', backgroundColor:\'rgba(45,212,191,.12)\', fill:true, tension:.25, pointRadius:0, borderWidth:2.5 },\n      { label:\'Cumulative Target Pace\', data:cumTarget, borderColor:\'#FF6B4A\', borderDash:[5,4], pointRadius:0, borderWidth:1.5, fill:false },\n    ]},\n    options:{ responsive:true, plugins:{legend:{position:\'top\', labels:{boxWidth:10, usePointStyle:true, color:\'#8592AE\', font:{size:11.5}}},\n      tooltip:{callbacks:{label:c=>`${c.dataset.label}: ${fmtShort(c.raw)}`}}},\n      scales:{ x:gridOpt, y:{...gridOpt, ticks:{...gridOpt.ticks, callback:v=>fmtShort(v)}} } }\n  });\n}\n\nfunction renderRegionChart(){\n  destroy(\'region\');\n  const regions = DATA.region;\n  charts.region = new Chart(document.getElementById(\'regionChart\'), {\n    type:\'bar\',\n    data:{ labels: regions.map(r=>r.Region), datasets:[\n      { label:\'Target\', data: regions.map(r=>r.Target), backgroundColor:\'rgba(255,255,255,.14)\', borderRadius:6, maxBarThickness:46 },\n      { label:\'Achieved\', data: regions.map(r=>r.Achieved), backgroundColor: regions.map(r=>r.Region===\'North\'?\'#2DD4BF\':\'#F5A524\'), borderRadius:6, maxBarThickness:46 },\n    ]},\n    options:{ responsive:true, plugins:{legend:{position:\'top\', labels:{boxWidth:10, usePointStyle:true, color:\'#8592AE\', font:{size:11.5}}},\n      tooltip:{callbacks:{label:c=>`${c.dataset.label}: ${fmtShort(c.raw)}`}}},\n      scales:{ x:gridOpt, y:{...gridOpt, ticks:{...gridOpt.ticks, callback:v=>fmtShort(v)}} } }\n  });\n}\n\nfunction renderCatChart(){\n  destroy(\'cat\');\n  const cats = (currentRegion===\'All\' ? DATA.category : DATA.categoryRegion[currentRegion]).slice(0,7);\n  const palette = [\'#2DD4BF\',\'#F5A524\',\'#FF6B4A\',\'#8592AE\',\'#3ECF8E\',\'#7C9CFF\',\'#E879F9\'];\n  charts.cat = new Chart(document.getElementById(\'catChart\'), {\n    type:\'doughnut\',\n    data:{ labels: cats.map(c=>c[\'Product Category\']), datasets:[{ data: cats.map(c=>c.Revenue), backgroundColor:palette, borderColor:\'#111827\', borderWidth:2 }] },\n    options:{ responsive:true, cutout:\'62%\', plugins:{ legend:{position:\'right\', labels:{boxWidth:9, color:\'#8592AE\', font:{size:10.5}, padding:8}},\n      tooltip:{callbacks:{label:c=>`${c.label}: ${fmtShort(c.raw)}`}} } }\n  });\n}\n\nfunction renderRegionCards(){\n  document.getElementById(\'regionCards\').innerHTML = DATA.region.map(r=>{\n    const pct = (r.Achieved/r.Target*100);\n    const color = r.Region===\'North\' ? \'var(--north)\' : \'var(--south)\';\n    return `<div class="rcard">\n      <div class="rcard-top"><div class="rcard-name"><i style="background:${color}"></i>${r.Region}</div><div class="rcard-pct" style="color:${color}">${pct.toFixed(1)}%</div></div>\n      <div class="bar-outer"><div class="bar-inner" style="width:${Math.min(100,pct)}%;background:${color}"></div></div>\n      <div class="rcard-meta"><span>${r.Reps} reps</span><span>${fmtShort(r.Achieved)} / ${fmtShort(r.Target)}</span></div>\n    </div>`;\n  }).join(\'\');\n}\n\nfunction renderCatStackChart(){\n  destroy(\'catStack\');\n  const cats = DATA.topCategories.slice(0,5);\n  const palette = [\'#2DD4BF\',\'#F5A524\',\'#FF6B4A\',\'#8592AE\',\'#3ECF8E\'];\n  const labels = DATA.dailyCategory.map(d=>fmtDateShort(d.DateStr));\n  charts.catStack = new Chart(document.getElementById(\'catStackChart\'), {\n    type:\'bar\',\n    data:{ labels, datasets: cats.map((c,i)=>({ label:c, data: DATA.dailyCategory.map(d=>d[c]), backgroundColor:palette[i], stack:\'s\' })) },\n    options:{ responsive:true, plugins:{legend:{position:\'top\', labels:{boxWidth:9, usePointStyle:true, color:\'#8592AE\', font:{size:10.5}}},\n      tooltip:{callbacks:{label:c=>`${c.dataset.label}: ${fmtShort(c.raw)}`}}},\n      scales:{ x:{...gridOpt, stacked:true}, y:{...gridOpt, stacked:true, ticks:{...gridOpt.ticks, callback:v=>fmtShort(v)}} } }\n  });\n}\n\nfunction renderDowChart(){\n  destroy(\'dow\');\n  const src = currentRegion===\'All\' ? DATA.dow : DATA.dowRegion[currentRegion];\n  charts.dow = new Chart(document.getElementById(\'dowChart\'), {\n    type:\'bar\',\n    data:{ labels: src.map(d=>d.DOW.slice(0,3)), datasets:[{ label:\'Avg-style Revenue\', data: src.map(d=>d.Revenue), backgroundColor:\'#7C9CFF\', borderRadius:6, maxBarThickness:52 }] },\n    options:{ responsive:true, plugins:{legend:{display:false}, tooltip:{callbacks:{label:c=>fmtShort(c.raw)}}},\n      scales:{ x:gridOpt, y:{...gridOpt, ticks:{...gridOpt.ticks, callback:v=>fmtShort(v)}} } }\n  });\n}\n\nfunction renderSlabChart(){\n  destroy(\'slab\');\n  charts.slab = new Chart(document.getElementById(\'slabChart\'), {\n    type:\'bar\',\n    data:{ labels: DATA.slab.labels, datasets:[{ data: DATA.slab.values, backgroundColor:[\'#F2495C\',\'#F5A524\',\'#F5A524\',\'#3ECF8E\',\'#2DD4BF\'], borderRadius:6, maxBarThickness:52 }] },\n    options:{ indexAxis:\'y\', responsive:true, plugins:{legend:{display:false}, tooltip:{callbacks:{label:c=>c.raw+\' reps\'}}},\n      scales:{ x:{...gridOpt, ticks:{...gridOpt.ticks}}, y:gridOpt } }\n  });\n}\n\n// ---------------- TABLES ----------------\nfunction renderTable(id, rows){\n  const tbody = document.querySelector(\'#\'+id+\' tbody\');\n  tbody.innerHTML = rows.map((r,i)=>{\n    const pct = r[\'Achievement in %\']*100;\n    return `<tr>\n      <td class="rankcell">${i+1}</td>\n      <td>${r.Name}</td>\n      <td style="color:${r.Region===\'North\'?\'var(--north)\':\'var(--south)\'}">${r.Region}</td>\n      <td class="num">${fmtShort(r[\'Revenue Target\'])}</td>\n      <td class="num">${fmtShort(r[\'Revenue Achived\'])}</td>\n      <td class="num"><span class="pct-tag ${pct>=100?\'hi\':\'lo\'}">${pct.toFixed(0)}%</span></td>\n    </tr>`;\n  }).join(\'\');\n}\n\nfunction renderTables(){\n  const top = currentRegion===\'All\' ? DATA.top10 : DATA.top10Region[currentRegion];\n  const bottom = currentRegion===\'All\' ? DATA.bottom10 : DATA.bottom10Region[currentRegion];\n  renderTable(\'topTable\', top);\n  renderTable(\'bottomTable\', bottom);\n}\n\nfunction renderFomTable(){\n  const tbody = document.querySelector(\'#fomTable tbody\');\n  tbody.innerHTML = DATA.fom.map(f=>{\n    const pct = f.AchPct;\n    return `<tr>\n      <td>${f[\'Field Operations Manager\']}</td>\n      <td class="num">${f.Reps}</td>\n      <td class="num">${fmtShort(f.Target)}</td>\n      <td class="num">${fmtShort(f.Achieved)}</td>\n      <td class="num"><span class="pct-tag ${pct>=100?\'hi\':\'lo\'}">${pct.toFixed(0)}%</span></td>\n    </tr>`;\n  }).join(\'\');\n}\n\n// ---------------- FILTER WIRING ----------------\ndocument.getElementById(\'regionFilters\').addEventListener(\'click\', e=>{\n  const btn = e.target.closest(\'.tab\');\n  if(!btn) return;\n  document.querySelectorAll(\'#regionFilters .tab\').forEach(b=>b.classList.remove(\'active\'));\n  btn.classList.add(\'active\');\n  currentRegion = btn.dataset.region;\n  renderAll();\n});\n\nfunction renderAll(){\n  renderKPIs();\n  renderDayStrip();\n  renderDayDetail();\n  renderTrendChart();\n  renderPaceChart();\n  renderRegionChart();\n  renderCatChart();\n  renderRegionCards();\n  renderCatStackChart();\n  renderDowChart();\n  renderSlabChart();\n  renderTables();\n  renderFomTable();\n}\n\nrenderAll();\n\n'



def render_html(data: dict, meta: dict) -> str:
    meta = {**meta, "source_file": meta.get("source_file", "")}
    head = HEAD_TOP
    body = BODY_HTML
    for k, v in meta.items():
        token = "{{" + k.upper() + "}}"
        head = head.replace(token, str(v))
        body = body.replace(token, str(v))

    data_json = json.dumps(data, default=str, separators=(",", ":"))
    html = (
        head + STYLE_CSS + "</head>" + body
        + "<script>const DATA = " + data_json + ";</script>"
        + APP_JS + "</html>"
    )
    return html


def generate(input_path, output_path=None):
    """Callable version for notebooks / Jupyter -- no command line needed.

    Example (in a Jupyter cell):
        from make_dashboard import generate
        generate("Sales_Representative_Target_vs_Achievement_Report_Denave_August.xlsm")
    """
    in_path = Path(input_path)
    if not in_path.exists():
        raise FileNotFoundError(f"File not found: {in_path}")

    out_path = Path(output_path) if output_path else in_path.with_name(in_path.stem + "_dashboard.html")

    print(f"Reading {in_path.name} ...")
    raw, tgt = load_workbook(in_path)

    print("Crunching daily / regional / rep-level breakdowns ...")
    data, meta = build_dataset(raw, tgt)
    meta["source_file"] = in_path.name

    print(f"Writing {out_path} ...")
    html = render_html(data, meta)
    out_path.write_text(html, encoding="utf-8")

    print(f"Done. Open {out_path} in a browser (needs internet once, to load Chart.js + fonts).")
    return out_path


def main():
    # Running inside Jupyter/IPython? sys.argv is the notebook kernel's own
    # args (e.g. "-f ...connection.json"), not yours -- argparse will fail on
    # those. Detect that case and point to the notebook-friendly generate().
    if "ipykernel_launcher" in sys.argv[0]:
        print(
            "It looks like this is running inside Jupyter/IPython, not a terminal.\n"
            "Command-line arguments don't work here. Instead, run in a cell:\n\n"
            "    from make_dashboard import generate\n"
            "    generate(r'Sales_Representative_Target_vs_Achievement_Report_Denave_August.xlsm')\n\n"
            "(Use the full path to your file if it's not in the current folder.)\n\n"
            "To use it from the command line instead, open a terminal / Anaconda Prompt and run:\n\n"
            "    python make_dashboard.py \"Sales_Representative_Target_vs_Achievement_Report_Denave_August.xlsm\"\n"
        )
        return

    parser = argparse.ArgumentParser(description="Generate the Denave x Canon daily performance dashboard from a raw .xlsm export.")
    parser.add_argument("input", help="Path to the Sales_Representative_Target_vs_Achievement_Report_Denave_<Month>.xlsm file")
    parser.add_argument("-o", "--output", help="Output HTML path (default: <input_name>_dashboard.html)")
    args = parser.parse_args()

    generate(args.input, args.output)


if __name__ == "__main__":
    main()
