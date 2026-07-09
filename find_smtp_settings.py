#!/usr/bin/env python3
"""
find_smtp_settings.py
======================
Denave's webmail (https://webmail.denave.com/aurora/) is an Aurora/AfterLogic
front-end. Aurora is just an inbox UI -- the *actual* mail-sending server
(SMTP) it talks to behind the scenes can be on a different hostname, and
there's no universal way to guess it. This script tries the handful of
hostnames/ports that are almost always right for a setup like this, using
your real mailbox login, and tells you which one actually works.

USAGE
-----
    pip install --break-system-packages nothing-needed  # uses only stdlib
    python find_smtp_settings.py

It will prompt for your password so it's never left sitting in a file.
Once it prints a working HOST/PORT, plug those into send_region_dashboards.py
(via SMTP_HOST / SMTP_PORT env vars, or the Streamlit sidebar).
"""

import getpass
import smtplib
import socket
import ssl

USER = "canonreport@denave.com"

# Common candidates for a domain whose webmail lives at webmail.denave.com.
# (host, port, mode) -- mode is "starttls" or "ssl"
CANDIDATES = [
    ("webmail.denave.com", 587, "starttls"),
    ("webmail.denave.com", 465, "ssl"),
    ("webmail.denave.com", 25, "starttls"),
    ("mail.denave.com", 587, "starttls"),
    ("mail.denave.com", 465, "ssl"),
    ("smtp.denave.com", 587, "starttls"),
    ("smtp.denave.com", 465, "ssl"),
    ("denave.com", 587, "starttls"),
]


def try_one(host, port, mode, user, password, timeout=8):
    try:
        if mode == "ssl":
            context = ssl.create_default_context()
            server = smtplib.SMTP_SSL(host, port, timeout=timeout, context=context)
        else:
            server = smtplib.SMTP(host, port, timeout=timeout)
            server.ehlo()
            server.starttls(context=ssl.create_default_context())
            server.ehlo()
        server.login(user, password)
        server.quit()
        return True, "login OK"
    except smtplib.SMTPAuthenticationError as e:
        return False, f"connected, but login rejected: {e}"
    except (socket.timeout, socket.gaierror, ConnectionRefusedError, OSError) as e:
        return False, f"could not connect: {e}"
    except smtplib.SMTPException as e:
        return False, f"SMTP error: {e}"


def main():
    password = getpass.getpass(f"Password for {USER}: ")
    print(f"\nProbing {len(CANDIDATES)} host/port combinations for {USER} ...\n")

    working = []
    for host, port, mode in CANDIDATES:
        label = f"{host}:{port} ({mode})"
        ok, detail = try_one(host, port, mode, USER, password)
        status = "WORKS" if ok else "no"
        print(f"  [{status:5}] {label:35} -- {detail}")
        if ok:
            working.append((host, port))

    print()
    if working:
        host, port = working[0]
        print("Found a working combination:")
        print(f"  SMTP_HOST = {host}")
        print(f"  SMTP_PORT = {port}")
        print("\nUse these in send_region_dashboards.py (env vars or Streamlit sidebar).")
    else:
        print("None of the common candidates worked.")
        print("Next step: in the Aurora webmail UI, check Settings -> usually an")
        print("'Identities' or 'Show settings for mail clients' section lists the")
        print("real IMAP/SMTP server names -- or ask whoever administers")
        print("webmail.denave.com for the outgoing (SMTP) server + port.")


if __name__ == "__main__":
    main()
