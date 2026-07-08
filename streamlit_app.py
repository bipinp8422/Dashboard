"""
Streamlit front-end for the Denave x Canon CPP dashboard generator.

Lets you upload a Sales_Representative_Target_vs_Achievement_Report_Denave_<Month>.xlsm
file in a browser, builds the same HTML dashboard as make_dashboard.py, shows it
inline, and gives you a button to download the standalone HTML file.

It also lets you PREVIEW the North / South region emails (subject, recipients,
body, and the exact dashboard that would be attached), and -- if this app is
running on the SAME Windows machine as your Outlook desktop app -- open that
exact preview as a real Outlook draft with one click, so you can double-check
it and hit Send yourself. Nothing is ever sent automatically from this app.

RUN IT
------
    pip install streamlit pandas openpyxl
    # Only needed for the "Open in Outlook" button, and only works on Windows
    # with the Outlook desktop app installed and signed in:
    pip install pywin32
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


def open_in_outlook(to_addrs, cc_addrs, subject: str, body: str, attachment_path: Path):
    """Create a real Outlook draft with .Display() -- opens the compose
    window for the person to review and send themselves. Never calls
    .Send(), so nothing goes out on its own.

    Only works when this Streamlit app is running on the SAME Windows
    machine as a signed-in Outlook desktop install (COM automation can't
    reach across machines or into a browser-only/web Outlook session).
    """
    try:
        import win32com.client
    except ImportError:
        return False, (
            "The 'Open in Outlook' button needs the `pywin32` package, and only works when "
            "this app runs on Windows with the Outlook desktop app installed "
            "(`pip install pywin32`, then restart the app)."
        )

    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)  # 0 = olMailItem
        mail.To = "; ".join(to_addrs)
        mail.CC = "; ".join(cc_addrs)
        mail.Subject = subject
        mail.Body = body
        mail.Attachments.Add(str(attachment_path.resolve()))
        mail.Display()  # opens the draft window -- does NOT send it
        return True, None
    except Exception as e:
        return False, (
            "Couldn't open Outlook. Make sure the Outlook desktop app is installed, signed in, "
            f"and running on this machine.\n\nDetails: {e}"
        )


def show_email_preview(region: str, region_html: str, month_label: str, source_name: str):
    to_addrs, cc_addrs, subject, body = region_email_content(region, month_label)

    st.markdown(f"#### ✉️ Email preview -- {region}")
    st.text_input(f"To ({region})", value=", ".join(to_addrs), disabled=True, key=f"to_{region}")
    st.text_input(f"Cc ({region})", value=", ".join(cc_addrs), disabled=True, key=f"cc_{region}")
    st.text_input(f"Subject ({region})", value=subject, disabled=True, key=f"subj_{region}")
    st.text_area(f"Body ({region})", value=body, height=220, disabled=True, key=f"body_{region}")

    attachment_name = f"{Path(source_name).stem}_{region.upper()}.html"
    st.caption(f"📎 Attachment: {attachment_name}")

    with st.expander(f"Preview attached dashboard ({region} view)"):
        if hasattr(st, "iframe"):
            st.iframe(region_html, height=1000)
        else:
            import streamlit.components.v1 as components
            components.html(region_html, height=1000, scrolling=True)

    if st.button(f"📧 Open this in Outlook (draft, not sent)", key=f"outlook_{region}", use_container_width=True):
        # Outlook attaches from a real file on disk, so write the exact
        # previewed HTML out to a temp file first.
        tmp_dir = Path(tempfile.mkdtemp())
        attachment_path = tmp_dir / attachment_name
        attachment_path.write_text(region_html, encoding="utf-8")

        with st.spinner("Opening Outlook..."):
            ok, error = open_in_outlook(to_addrs, cc_addrs, subject, body, attachment_path)

        if ok:
            st.success(
                f"Draft opened in Outlook for {region} -- exactly what's shown above. "
                "Review it there and hit Send whenever you're ready."
            )
        else:
            st.error(error)

    st.info("Nothing is sent automatically. The button above only opens a draft in your "
            "own Outlook for you to review and send -- or use "
            "`python send_region_dashboards.py <file.xlsm> --send` for unattended SMTP sending.")


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
    if "show_preview" not in st.session_state:
        st.session_state.show_preview = {"North": False, "South": False}

    if pc1.button("👁️ Preview North email", use_container_width=True):
        st.session_state.show_preview["North"] = True
    if pc2.button("👁️ Preview South email", use_container_width=True):
        st.session_state.show_preview["South"] = True

    if st.session_state.show_preview["North"]:
        north_html, north_meta = render_region_only_dashboard(raw, tgt, "North", uploaded.name)
        show_email_preview("North", north_html, north_meta["month_label"], uploaded.name)

    if st.session_state.show_preview["South"]:
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
