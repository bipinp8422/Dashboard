"""
Streamlit wrapper for the Denave x Canon CPP Daily Performance Cockpit.

Run locally:
    pip install streamlit pandas openpyxl
    streamlit run streamlit_app.py

Deploy on Streamlit Community Cloud:
    1. Put this file + make_dashboard.py + requirements.txt in a GitHub repo.
    2. Go to https://share.streamlit.io -> "New app" -> pick the repo/branch.
    3. Set "Main file path" to streamlit_app.py -> Deploy.
"""

import streamlit as st
import streamlit.components.v1 as components
from pathlib import Path
import tempfile

# These come straight from your existing script -- no changes needed there.
from make_dashboard import load_workbook, build_dataset, render_html

st.set_page_config(
    page_title="Denave x Canon CPP — Daily Performance Cockpit",
    page_icon="📊",
    layout="wide",
)

st.title("📊 Daily Performance Cockpit")
st.caption(
    "Upload the Sales_Representative_Target_vs_Achievement_Report_Denave_<Month>.xlsm "
    "export and get the interactive dashboard."
)

uploaded = st.file_uploader("Upload the .xlsm report", type=["xlsm", "xlsx"])

if uploaded is not None:
    # Save the upload to a temp file so pandas/openpyxl can read it by path.
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsm") as tmp:
        tmp.write(uploaded.getvalue())
        tmp_path = Path(tmp.name)

    with st.spinner("Reading workbook and building the dashboard…"):
        try:
            raw, tgt = load_workbook(tmp_path)
            data, meta = build_dataset(raw, tgt)
            meta["source_file"] = uploaded.name
            html = render_html(data, meta)
        except Exception as e:
            st.error(f"Couldn't build the dashboard: {e}")
            st.stop()

    st.success(f"Dashboard built for **{meta['month_label']}** — {meta['rep_count']} reps.")

    # Quick top-line numbers right in Streamlit (optional, no extra deps needed).
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Target", f"₹{data['kpi']['totalTarget']/1e7:.2f} Cr")
    c2.metric("Achieved", f"₹{data['kpi']['totalAchieved']/1e7:.2f} Cr")
    c3.metric("Achievement %", f"{data['kpi']['achPct']:.1f}%")
    c4.metric("Reps above 100%", f"{data['kpi']['repsAbove100']} / {data['kpi']['totalReps']}")

    st.download_button(
        label="⬇️ Download full dashboard (.html)",
        data=html,
        file_name=f"{Path(uploaded.name).stem}_dashboard.html",
        mime="text/html",
    )

    st.divider()
    st.subheader("Preview")
    # Embed the full interactive dashboard (charts, tabs, day explorer) inline.
    components.html(html, height=1400, scrolling=True)

else:
    st.info("Upload a workbook above to generate the dashboard.")
