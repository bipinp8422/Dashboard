"""
Streamlit front-end for the Denave x Canon CPP dashboard generator.

Lets you upload a Sales_Representative_Target_vs_Achievement_Report_Denave_<Month>.xlsm
file in a browser, builds the same HTML dashboard as make_dashboard.py, shows it
inline, and gives you a button to download the standalone HTML file.

It also lets you PREVIEW the North / South region emails (subject, recipients,
body, and the exact dashboard that would be attached) without sending anything.
Actual sending still only happens via send_region_dashboards.py --send.

RUN IT
------
    pip install streamlit pandas openpyxl
    streamlit run streamlit_app.py

Make sure make_dashboard.py and send_region_dashboards.py sit in the SAME
folder as this file -- this app imports functions directly from both instead
of duplicating them.
"""

import tempfile
from pathlib import Path

import streamlit as st

from make_dashboard import load_workbook, build_dataset, render_html
from send_region_dashboards import (
    render_region_only_dashboard,
    NORTH_TO,
    NORTH_CC,
    SOUTH_TO,
    SOUTH_CC,
)

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


def region_email_content(region: str, month_label: str):
    """Same subject/body text used in send_region_dashboards.py, kept here
    only for preview purposes -- no email is sent from this app."""
    to_addrs = NORTH_TO if region == "North" else SOUTH_TO
    cc_addrs = NORTH_CC if region == "North" else SOUTH_CC

    subject = f"Denave x Canon CPP - Daily Performance Dashboard ({region}) - {month_label}"
    body = (
        "Hi all,\n\n"
        f"Please find attached the daily performance dashboard for the {region} region for {month_label} "
        "(month-to-date).\n\n"
        f"The dashboard opens directly on the {region} view and includes:\n"
        "- Daily revenue trend and pacing vs target\n"
        "- Day-by-day breakdown (Day Explorer) - click any day for top reps, category mix, and top cities\n"
        "- Product category mix\n"
        "- Top and bottom performers\n"
        "- FOM-wise team rollup\n\n"
        "It's a standalone HTML file - just open it in any browser, no installation needed.\n\n"
        "Regards,"
    )
    return to_addrs, cc_addrs, subject, body


def show_email_preview(region: str, region_html: str, month_label: str, source_name: str):
    to_addrs, cc_addrs, subject, body = region_email_content(region, month_label)

    st.markdown(f"#### ✉️ Email preview -- {region}")
    st.text_input(f"To ({region})", value=", ".join(to_addrs), disabled=True, key=f"to_{region}")
    st.text_input(f"Cc ({region})", value=", ".join(cc_addrs), disabled=True, key=f"cc_{region}")
    st.text_input(f"Subject ({region})", value=subject, disabled=True, key=f"subj_{region}")
    st.text_area(f"Body ({region})", value=body, height=220, disabled=True, key=f"body_{region}")
    st.caption(f"📎 Attachment: {Path(source_name).stem}_{region.upper()}.html")

    with st.expander(f"Preview attached dashboard ({region} view)"):
        if hasattr(st, "iframe"):
            st.iframe(region_html, height=1000)
        else:
            import streamlit.components.v1 as components
            components.html(region_html, height=1000, scrolling=True)

    st.info("This is a preview only -- no email has been sent. To actually send it, "
            "run `python send_region_dashboards.py <file.xlsm> --send` with your SMTP "
            "environment variables set.")


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
    st.subheader("✉️ Preview region emails")
    st.caption(
        "See exactly what the North / South emails would look like -- recipients, subject, "
        "body, and the attached dashboard -- without sending anything. The attached dashboard "
        "is built only from that region's rows (the other region's data is never loaded into "
        "it), so there's nothing to click over and see."
    )

    pc1, pc2 = st.columns(2)
    preview_north = pc1.button("👁️ Preview North email", use_container_width=True)
    preview_south = pc2.button("👁️ Preview South email", use_container_width=True)

    if preview_north:
        north_html, north_meta = render_region_only_dashboard(raw, tgt, "North", uploaded.name)
        show_email_preview("North", north_html, north_meta["month_label"], uploaded.name)

    if preview_south:
        south_html, south_meta = render_region_only_dashboard(raw, tgt, "South", uploaded.name)
        show_email_preview("South", south_html, south_meta["month_label"], uploaded.name)

    st.divider()
    st.subheader("Full dashboard (All regions)")
    if hasattr(st, "iframe"):
        # Streamlit >= 1.56: newer, non-deprecated API. Accepts a raw HTML
        # string directly. A fixed height is used since the page container
        # itself has no defined height for "stretch" to fill against.
        st.iframe(html, height=3000)
    else:
        # Older Streamlit versions don't have st.iframe yet.
        import streamlit.components.v1 as components
        components.html(html, height=3000, scrolling=True)

else:
    st.info("Waiting for a file upload to get started.")
