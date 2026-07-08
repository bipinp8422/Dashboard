#!/usr/bin/env python3
"""
send_region_dashboards.py
==========================

Extends make_dashboard.py to:
  1. Build the full dashboard dataset from a raw .xlsm export
  2. Generate a North-only and a South-only variant (each opens straight on
     that region's tab)
  3. Optionally email each variant to the right recipient list as an
     attachment

USAGE
-----
    pip install pandas openpyxl

    # Just generate the two files (safe, no email sent):
    python send_region_dashboards.py report.xlsm

    # Generate AND send both emails:
    python send_region_dashboards.py report.xlsm --send

Sending mail requires SMTP settings as environment variables (nothing is
hardcoded or stored in this file):

    SMTP_HOST       e.g. smtp.office365.com
    SMTP_PORT       e.g. 587
    SMTP_USER       your full email address
    SMTP_PASSWORD   your mailbox password or app password
    SMTP_FROM       (optional) defaults to SMTP_USER

Example (Windows PowerShell):
    $env:SMTP_HOST="smtp.office365.com"
    $env:SMTP_PORT="587"
    $env:SMTP_USER="you@canon.co.in"
    $env:SMTP_PASSWORD="your-app-password"
    python send_region_dashboards.py report.xlsm --send

Example (Mac/Linux):
    export SMTP_HOST=smtp.office365.com
    export SMTP_PORT=587
    export SMTP_USER=you@canon.co.in
    export SMTP_PASSWORD=your-app-password
    python send_region_dashboards.py report.xlsm --send

make_dashboard.py must sit in the same folder as this script -- it's
imported directly for the data-processing and HTML-rendering functions.
"""

import argparse
import os
import smtplib
import ssl
import sys
from email.message import EmailMessage
from pathlib import Path

from make_dashboard import load_workbook, build_dataset, render_html

# ---------------------------------------------------------------------------
# Recipient lists -- edit these if the distribution list changes.
# ---------------------------------------------------------------------------
NORTH_TO = [
    "rajender.singh@canon.co.in", "Vijay.Verma@canon.co.in", "rishabh.kapoor@canon.co.in",
    "shireesh.vaishnav@canon.co.in", "amit.nanchahal@canon.co.in", "abhishek.singh026@canon.co.in",
    "Satender.Yadav@canon.co.in", "Jaideep.SinghTEMP@canon.co.in",
]
NORTH_CC = [
    "KSHarihara.Prasad@canon.co.in", "Garima.Mohendroo@canon.co.in", "Ujjwal.JoshiTEMP@canon.co.in",
    "shashanks@denave.com", "ce-nikhil.verma@denave.com", "yunus.khan@canon.co.in",
]

SOUTH_TO = [
    "Mudari.satyanarayana@canon.co.in", "Al.Subramanian@canon.co.in", "sarath.nath@canon.co.in",
    "asiva.kumar@canon.co.in", "praveen.k@canon.co.in", "shivaraj.shettyk@canon.co.in",
    "Harikrishna.C@canon.co.in", "Shankar.d@canon.co.in",
]
SOUTH_CC = [
    "KSHarihara.Prasad@canon.co.in", "Garima.Mohendroo@canon.co.in", "sujesh.soman@canon.co.in",
    "Ujjwal.JoshiTEMP@canon.co.in", "shashanks@denave.com", "ce-nikhil.verma@denave.com",
]


def make_region_variant(html: str, region: str) -> str:
    """Return a copy of the dashboard HTML that opens directly on `region`
    ('North' or 'South') instead of 'All'."""
    out = html.replace("let currentRegion = 'All';", f"let currentRegion = '{region}';")
    out = out.replace(
        '<button class="tab active" data-region="All">All Regions</button>\n'
        '    <button class="tab" data-region="North">North</button>\n'
        '    <button class="tab" data-region="South">South</button>',
        '<button class="tab{all_active}" data-region="All">All Regions</button>\n'
        '    <button class="tab{north_active}" data-region="North">North</button>\n'
        '    <button class="tab{south_active}" data-region="South">South</button>'.format(
            all_active=" active" if region == "All" else "",
            north_active=" active" if region == "North" else "",
            south_active=" active" if region == "South" else "",
        ),
    )
    return out


def build_region_files(xlsm_path: Path, out_dir: Path, month_label: str):
    raw, tgt = load_workbook(xlsm_path)
    data, meta = build_dataset(raw, tgt)
    meta["source_file"] = xlsm_path.name
    base_html = render_html(data, meta)

    stem = xlsm_path.stem
    north_path = out_dir / f"{stem}_NORTH.html"
    south_path = out_dir / f"{stem}_SOUTH.html"

    north_path.write_text(make_region_variant(base_html, "North"), encoding="utf-8")
    south_path.write_text(make_region_variant(base_html, "South"), encoding="utf-8")

    return north_path, south_path, meta


def send_email(to_addrs, cc_addrs, subject, body, attachment_path: Path):
    host = os.environ.get("SMTP_HOST")
    port = int(os.environ.get("SMTP_PORT", "587"))
    user = os.environ.get("SMTP_USER")
    password = os.environ.get("SMTP_PASSWORD")
    sender = os.environ.get("SMTP_FROM", user)

    missing = [k for k, v in [("SMTP_HOST", host), ("SMTP_USER", user), ("SMTP_PASSWORD", password)] if not v]
    if missing:
        sys.exit(
            "Missing required environment variable(s): " + ", ".join(missing) +
            "\nSet these before running with --send. See the top of this script for examples."
        )

    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = sender
    msg["To"] = ", ".join(to_addrs)
    if cc_addrs:
        msg["Cc"] = ", ".join(cc_addrs)
    msg.set_content(body)

    with open(attachment_path, "rb") as f:
        msg.add_attachment(
            f.read(), maintype="text", subtype="html", filename=attachment_path.name
        )

    all_recipients = to_addrs + cc_addrs
    context = ssl.create_default_context()
    with smtplib.SMTP(host, port) as server:
        server.starttls(context=context)
        server.login(user, password)
        server.send_message(msg, from_addr=sender, to_addrs=all_recipients)

    print(f"Sent: {subject}  ->  {len(all_recipients)} recipients")


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("input", help="Path to the Sales_Representative_Target_vs_Achievement_Report_Denave_<Month>.xlsm file")
    parser.add_argument("--send", action="store_true", help="Actually send the emails (requires SMTP_* environment variables). Without this flag, only the two HTML files are generated.")
    parser.add_argument("-o", "--output-dir", default=None, help="Directory to write the North/South HTML files (default: same folder as input)")
    args = parser.parse_args()

    in_path = Path(args.input)
    if not in_path.exists():
        sys.exit(f"File not found: {in_path}")

    out_dir = Path(args.output_dir) if args.output_dir else in_path.parent
    out_dir.mkdir(parents=True, exist_ok=True)

    print(f"Reading {in_path.name} ...")
    north_path, south_path, meta = build_region_files(in_path, out_dir, None)
    print(f"Wrote {north_path}")
    print(f"Wrote {south_path}")

    if not args.send:
        print("\n--send not passed, so no email was sent. Attach these files manually, or re-run with --send.")
        return

    month_label = meta["month_label"]

    north_subject = f"Denave x Canon CPP - Daily Performance Dashboard (North) - {month_label}"
    north_body = (
        "Hi all,\n\n"
        f"Please find attached the daily performance dashboard for the North region for {month_label} "
        "(month-to-date).\n\n"
        "The dashboard opens directly on the North view and includes:\n"
        "- Daily revenue trend and pacing vs target\n"
        "- Day-by-day breakdown (Day Explorer) - click any day for top reps, category mix, and top cities\n"
        "- Product category mix\n"
        "- Top and bottom performers\n"
        "- FOM-wise team rollup\n\n"
        "It's a standalone HTML file - just open it in any browser, no installation needed.\n\n"
        "Regards,"
    )

    south_subject = f"Denave x Canon CPP - Daily Performance Dashboard (South) - {month_label}"
    south_body = (
        "Hi all,\n\n"
        f"Please find attached the daily performance dashboard for the South region for {month_label} "
        "(month-to-date).\n\n"
        "The dashboard opens directly on the South view and includes:\n"
        "- Daily revenue trend and pacing vs target\n"
        "- Day-by-day breakdown (Day Explorer) - click any day for top reps, category mix, and top cities\n"
        "- Product category mix\n"
        "- Top and bottom performers\n"
        "- FOM-wise team rollup\n\n"
        "It's a standalone HTML file - just open it in any browser, no installation needed.\n\n"
        "Regards,"
    )

    send_email(NORTH_TO, NORTH_CC, north_subject, north_body, north_path)
    send_email(SOUTH_TO, SOUTH_CC, south_subject, south_body, south_path)


if __name__ == "__main__":
    main()
