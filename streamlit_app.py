"""
Streamlit front-end for the Denave x Canon CPP dashboard generator.

Lets you upload a Sales_Representative_Target_vs_Achievement_Report_Denave_<Month>.xlsm
file in a browser, builds the same HTML dashboard as make_dashboard.py, shows it
inline, and gives you a button to download the standalone HTML file.

RUN IT
------
    pip install streamlit pandas openpyxl
    streamlit run streamlit_app.py

Make sure make_dashboard.py sits in the SAME folder as this file -- this app
imports its data-processing functions directly instead of duplicating them.
"""

import tempfile
from pathlib import Path

import streamlit as st
import streamlit.components.v1 as components

from make_dashboard import load_workbook, build_dataset, render_html

st.set_page_config(
    page_title="Denave x Canon CPP -- Dashboard Generator",
    page_icon="📊",
    layout="wide",
)

st.title("📊 Denave x Canon CPP -- Daily Performance Dashboard Generator")
st.caption(
    "Upload the monthly Sales_Representative_Target_vs_Achievement_Report_Denave_*.xlsm "
    "export and get the full interactive dashboard -- Day Explorer, daily trend, pacing "
    "chart, category mix, region tabs, and leaderboards -- built automatically."
)

uploaded = st.file_uploader(
    "Upload the .xlsm report",
    type=["xlsm"],
    help="Must contain a 'Raw Data' sheet and a 'Target vs Achievement' sheet, same as the source reports.",
)

if uploaded is not None:
    with tempfile.NamedTemporaryFile(suffix=".xlsm", delete=False) as tmp:
        tmp.write(uploaded.getbuffer())
        tmp_path = Path(tmp.name)

    try:
        with st.spinner("Reading workbook and crunching daily / regional / rep-level breakdowns..."):
            raw, tgt = load_workbook(tmp_path)
            data, meta = build_dataset(raw, tgt)
            meta["source_file"] = uploaded.name
            html = render_html(data, meta)
    except Exception as e:
        st.error(
            "Couldn't process this file. Double check it has a 'Raw Data' sheet and a "
            "'Target vs Achievement' sheet in the same layout as the source reports.\n\n"
            f"Details: {e}"
        )
        st.stop()

    kpi = data["kpi"]
    st.success(f"Dashboard built for **{meta['month_label']}** ({meta['days_active']} of {meta['days_in_month']} days active).")

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Target", f"₹{kpi['totalTarget']/1e7:.2f} Cr")
    c2.metric("Achieved", f"₹{kpi['totalAchieved']/1e7:.2f} Cr", f"{kpi['achPct']:.1f}% attained")
    c3.metric("Transactions", f"{kpi['totalTransactions']:,}")
    c4.metric("Reps above 100%", f"{kpi['repsAbove100']} / {kpi['totalReps']}")

    st.download_button(
        "⬇️ Download standalone HTML dashboard",
        data=html,
        file_name=f"{Path(uploaded.name).stem}_dashboard.html",
        mime="text/html",
        use_container_width=True,
    )

    st.divider()
    components.html(html, height=3000, scrolling=True)

else:
    st.info("Waiting for a file upload to get started.")
