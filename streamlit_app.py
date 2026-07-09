"""
Streamlit front-end for the Denave x Canon CPP dashboard generator.

Lets you upload a Sales_Representative_Target_vs_Achievement_Report_Denave_<Month>.xlsm
file in a browser, builds the same HTML dashboard as make_dashboard.py, shows it
inline, and gives you a button to download the standalone HTML file.

It also lets you PREVIEW the North / South region emails (subject, recipients,
body, and the exact dashboard that would be attached), and then actually SEND
that email -- with the dashboard attached automatically -- via the "✅ Send Now"
button. Nothing is ever sent until you click that button.

RUN IT
------
    pip install streamlit pandas openpyxl
    pip install pywin32          # only needed for the Outlook-sending option (Windows)
    streamlit run streamlit_app.py

Make sure make_dashboard.py and send_region_dashboards.py sit in the SAME
folder as this file -- this app imports functions directly from both instead
of duplicating them.

SENDING EMAIL
--------------
Two ways to send, chosen in the "Email sending settings" panel in the sidebar:

  - Outlook desktop (default, recommended if you're on Windows with Outlook
    installed): no password needed at all, since it drives the Outlook app
    you're already signed into. Choose whether to open each email as a
    normal Outlook draft (attachment already added, review then click Send
    yourself) or send it immediately with no window.

  - SMTP (advanced): needs real mail server credentials (host, port,
    username, password -- typically an app password, not your normal
    account password). Kept only in memory for this Streamlit session and
    never written to disk.

Either way, clicking the button on a region's preview attaches that
region's dashboard automatically -- no separate download-and-attach step.

Equivalent from a terminal, without opening Streamlit at all:
    python send_region_dashboards.py <file.xlsm> --send --via outlook
    python send_region_dashboards.py <file.xlsm> --send   (SMTP, needs SMTP_* env vars)
"""

import tempfile
from pathlib import Path

import streamlit as st

from make_dashboard import load_workbook, build_dataset, render_html
from send_region_dashboards import (
    render_region_only_dashboard,
    region_email_content,
    send_email,
    send_via_outlook,
    build_region_source_workbook,
    _convert_to_xlsx_if_needed,
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
    "Upload the report (.xlsm, .xlsb, or .xlsx)",
    type=["xlsm", "xlsb", "xlsx"],
    help="Must contain a 'Raw Data' sheet and a 'Target vs Achievement' sheet, same as the source reports. .xlsb files are converted automatically (needs LibreOffice available on this machine).",
)


# ---------------------------------------------------------------------------
# SMTP settings -- entered once per session, kept only in memory.
# ---------------------------------------------------------------------------
def sending_settings_panel():
    st.sidebar.header("✉️ Email sending settings")
    st.sidebar.caption("Using Outlook desktop by default -- no password needed, since it uses the Outlook app already signed in on this machine.")

    send_immediately = st.sidebar.checkbox(
        "Send immediately (skip review)",
        value=False,
        help="Off (default): opens each email as a normal Outlook draft, attachment already added, so you can look it over and click Send yourself. On: sends right away with no draft window.",
        key="outlook_send_immediately",
    )

    smtp_cfg = {"host": None, "port": None, "user": None, "password": None, "sender": None}
    use_smtp = False
    with st.sidebar.expander("Advanced: use SMTP instead of Outlook"):
        st.caption(
            "Only needed if you don't have Outlook desktop, or specifically want to send via a "
            "mail server instead. Kept in memory for this session only -- never written to disk."
        )
        use_smtp = st.checkbox("Use SMTP instead of Outlook for sending", value=False, key="use_smtp_toggle")
        host = st.text_input("SMTP host", value=st.session_state.get("smtp_host", ""), placeholder="smtp.office365.com").strip()
        port = st.text_input("SMTP port", value=st.session_state.get("smtp_port", "587")).strip()
        user = st.text_input("SMTP username (your email)", value=st.session_state.get("smtp_user", ""), placeholder="you@canon.co.in").strip()
        password = st.text_input("SMTP password (or app password)", value=st.session_state.get("smtp_password", ""), type="password").strip()
        sender = st.text_input("From address (optional, defaults to username)", value=st.session_state.get("smtp_sender", "")).strip()

        st.session_state["smtp_host"] = host
        st.session_state["smtp_port"] = port
        st.session_state["smtp_user"] = user
        st.session_state["smtp_password"] = password
        st.session_state["smtp_sender"] = sender

        smtp_cfg = {
            "host": host or None,
            "port": port or None,
            "user": user or None,
            "password": password or None,
            "sender": sender or None,
        }
        if use_smtp:
            configured = bool(host and user and password)
            if configured:
                st.success("SMTP settings entered -- will be used instead of Outlook.")
            else:
                st.warning("Fill in host, username, and password to actually send via SMTP.")

    if use_smtp:
        return {"method": "smtp", "smtp_config": smtp_cfg}
    return {"method": "outlook", "send_immediately": send_immediately}


def build_mailto_url(to_addrs, cc_addrs, subject: str, body: str) -> str:
    """Build a mailto: URL as a fallback for people who'd rather draft the
    email themselves in their own mail client. Browsers block mailto: links
    from auto-attaching files (a security restriction, not something this
    app can bypass) -- that's exactly why 'Send Now' below exists as the
    real one-click option."""
    import urllib.parse
    to = ",".join(to_addrs)
    params = {"subject": subject, "body": body}
    if cc_addrs:
        params["cc"] = ",".join(cc_addrs)
    query = urllib.parse.urlencode(params, quote_via=urllib.parse.quote)
    return f"mailto:{to}?{query}"


def show_email_preview(region: str, region_html: str, month_label: str, source_name: str, send_settings: dict, original_upload_path: Path):
    to_addrs, cc_addrs, subject, body = region_email_content(region, month_label)

    st.markdown(f"#### ✉️ Email preview -- {region}")
    st.text_input(f"To ({region})", value=", ".join(to_addrs), disabled=True, key=f"to_{region}")
    st.text_input(f"Cc ({region})", value=", ".join(cc_addrs), disabled=True, key=f"cc_{region}")
    st.text_input(f"Subject ({region})", value=subject, disabled=True, key=f"subj_{region}")
    st.text_area(f"Body ({region})", value=body, height=220, disabled=True, key=f"body_{region}")

    html_attachment_name = f"{Path(source_name).stem}_{region.upper()}.html"
    xlsx_attachment_name = f"{Path(source_name).stem}_{region.upper()}.xlsx"
    st.caption(
        f"📎 Attachments: {html_attachment_name} (dashboard) + {xlsx_attachment_name} "
        f"(region-filtered source data, all sheets) -- both included automatically when you send"
    )

    with st.expander(f"Preview attached dashboard ({region} view)"):
        if hasattr(st, "iframe"):
            st.iframe(region_html, height=1000)
        else:
            import streamlit.components.v1 as components
            components.html(region_html, height=1000, scrolling=True)

    # Build the region-filtered Excel workbook now so it's ready for both the
    # download button and the send buttons below -- built once per render,
    # cached in session_state keyed by region so re-renders don't rebuild it
    # unnecessarily while the user is just looking at the preview.
    xlsx_cache_key = f"region_xlsx_bytes_{region}"
    if xlsx_cache_key not in st.session_state:
        try:
            with st.spinner(f"Building {region} region data workbook..."):
                tmp_xlsx = Path(tempfile.mkstemp(suffix=".xlsx")[1])
                build_region_source_workbook(original_upload_path, region, tmp_xlsx)
                st.session_state[xlsx_cache_key] = tmp_xlsx.read_bytes()
                tmp_xlsx.unlink(missing_ok=True)
        except Exception as e:
            st.error(f"Couldn't build the {region} Excel attachment.\n\n{e}")
            st.stop()
    region_xlsx_bytes = st.session_state[xlsx_cache_key]

    method = send_settings["method"]
    if method == "outlook":
        configured = True  # no credentials needed, just needs Outlook installed
        button_label = "✅ Send Now via Outlook (both attachments included automatically)" if send_settings["send_immediately"] \
            else "✅ Open in Outlook (both attachments included, review before sending)"
    else:
        cfg = send_settings["smtp_config"]
        configured = bool(cfg["host"] and cfg["user"] and cfg["password"])
        button_label = "✅ Send Now (both attachments included automatically)"

    sc1, sc2, sc3 = st.columns(3)
    with sc1:
        send_clicked = st.button(
            button_label,
            use_container_width=True,
            type="primary",
            disabled=not configured,
            key=f"send_{region}",
        )
    with sc2:
        st.download_button(
            "⬇️ Download dashboard (.html)",
            data=region_html,
            file_name=html_attachment_name,
            mime="text/html",
            use_container_width=True,
            key=f"dl_html_{region}",
        )
    with sc3:
        st.download_button(
            "⬇️ Download region data (.xlsx)",
            data=region_xlsx_bytes,
            file_name=xlsx_attachment_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key=f"dl_xlsx_{region}",
        )

    if not configured:
        st.warning("Fill in the SMTP settings in the sidebar to enable 'Send Now'.")

    if send_clicked:
        spinner_msg = f"Preparing {region} email in Outlook..." if method == "outlook" else f"Sending {region} email to {len(to_addrs)} recipients..."
        with st.spinner(spinner_msg):
            try:
                with tempfile.NamedTemporaryFile(suffix=".html", delete=False) as tmp:
                    tmp.write(region_html.encode("utf-8"))
                    tmp_html_path = Path(tmp.name)
                # Give the temp files the same display names as the real attachments.
                real_html = tmp_html_path.with_name(html_attachment_name)
                tmp_html_path.replace(real_html)

                real_xlsx = real_html.with_name(xlsx_attachment_name)
                real_xlsx.write_bytes(region_xlsx_bytes)

                attachments = [real_html, real_xlsx]

                if method == "outlook":
                    send_via_outlook(
                        to_addrs, cc_addrs, subject, body, attachments,
                        send_immediately=send_settings["send_immediately"],
                    )
                    if send_settings["send_immediately"]:
                        st.success(f"Sent! {region} dashboard + region data emailed via Outlook to {len(to_addrs)} To + {len(cc_addrs)} Cc recipients, both attachments included.")
                    else:
                        st.success("Opened in Outlook -- both attachments already added. Review it and click Send whenever you're ready.")
                else:
                    send_email(to_addrs, cc_addrs, subject, body, attachments, smtp_config=send_settings["smtp_config"])
                    st.success(f"Sent! {region} dashboard + region data emailed to {len(to_addrs)} To + {len(cc_addrs)} Cc recipients, both attachments included.")
            except Exception as e:
                st.error(f"Couldn't send the {region} email: {e}")

    with st.expander("Prefer to send from your own mail app instead?"):
        mailto_url = build_mailto_url(to_addrs, cc_addrs, subject, body)
        if hasattr(st, "link_button"):
            st.link_button("📧 Open draft in your mail app", mailto_url, use_container_width=True)
        else:
            st.markdown(
                f'<a href="{mailto_url}" target="_blank" '
                f'style="display:block;text-align:center;padding:0.5rem;border:1px solid #999;'
                f'border-radius:0.5rem;text-decoration:none;">📧 Open draft in your mail app</a>',
                unsafe_allow_html=True,
            )
        st.caption(
            "Browsers block mailto: links from auto-attaching files (a security restriction, "
            "not something this app can get around), so you'd need to download both attachments "
            "above and attach them yourself in that draft. 'Send Now' above skips all of that."
        )


send_settings = sending_settings_panel()

if uploaded is not None:
    upload_suffix = Path(uploaded.name).suffix or ".xlsm"
    with tempfile.NamedTemporaryFile(suffix=upload_suffix, delete=False) as tmp:
        tmp.write(uploaded.getbuffer())
        tmp_path = Path(tmp.name)

    try:
        with st.spinner("Reading workbook and crunching daily / regional / rep-level breakdowns..."):
            working_path = _convert_to_xlsx_if_needed(tmp_path)
            raw, tgt = load_workbook(working_path)
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
    st.subheader("✉️ Region emails -- preview & send")
    st.caption(
        "See exactly what the North / South emails would look like -- recipients, subject, "
        "body, and the attached dashboard -- then send with one click, attachment included "
        "automatically. The attached dashboard is built only from that region's rows (the "
        "other region's data is never loaded into it)."
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
        show_email_preview("North", north_html, north_meta["month_label"], uploaded.name, send_settings, working_path)

    if st.session_state.show_preview["South"]:
        south_html, south_meta = render_region_only_dashboard(raw, tgt, "South", uploaded.name)
        show_email_preview("South", south_html, south_meta["month_label"], uploaded.name, send_settings, working_path)

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
