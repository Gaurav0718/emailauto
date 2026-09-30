#!/usr/bin/env python3
"""
Send an HTML file as the body of an email via Office365 SMTP, taking the
recipient list from an Excel sheet.

The HTML file itself IS the email body (no plain-text wrapper content).
It renders fit-to-screen on any device because the HTML carries a fixed
640px table layout that mail clients scale down.

Recipient sheet contract (works for any similar .xlsx):
    Column 3  -> email address
    Column 4  -> how to send it: "to", "cc", or "bcc" (case-insensitive)
A header row is auto-detected and skipped; blank rows are ignored.

Usage:
    export SMTP_USER="nudge-app@indegene.com"
    export SMTP_PASS="your-app-password"
    python3 send_email.py                       # defaults below
    python3 send_email.py mail.html list.xlsx   # override HTML and/or sheet
"""

import os
import sys
import smtplib
from email.message import EmailMessage
from email.utils import make_msgid, formatdate

import openpyxl

# ---- Configuration -------------------------------------------------------

SMTP_HOST = os.environ.get("SMTP_HOST", "smtp.office365.com")
SMTP_PORT = int(os.environ.get("SMTP_PORT", "587"))

SENDER = os.environ.get("SENDER", "nudge-app@indegene.com")
SMTP_USER = os.environ.get("SMTP_USER", SENDER)
SMTP_PASS = os.environ.get("SMTP_PASS")  # REQUIRED — set via env var, never hard-code

SUBJECT = "Account Intelligence Update"

DEFAULT_HTML = "NUDGE_Account_RoundUp_Issue02_Sep2026_1.html"
DEFAULT_XLSX = "/Users/gaurav/Downloads/GK List-250926.xlsx"

# Which columns hold the data (1-based, matching the sheet).
EMAIL_COL = 3   # column 3 -> email address
DISPOSITION_COL = 4   # column 4 -> to / cc / bcc

VALID_DISPOSITIONS = {"to", "cc", "bcc"}


# ---- Recipient loading ---------------------------------------------------

def load_recipients(xlsx_path):
    """Return dict {'to': [...], 'cc': [...], 'bcc': [...]} from the sheet.

    Column 3 = email, column 4 = disposition. Header row auto-skipped,
    blanks ignored, disposition normalized to lowercase.
    """
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws = wb.active

    buckets = {"to": [], "cc": [], "bcc": []}

    for row in ws.iter_rows(min_row=1, values_only=True):
        # Guard short rows.
        if len(row) < max(EMAIL_COL, DISPOSITION_COL):
            continue

        email = row[EMAIL_COL - 1]
        disposition = row[DISPOSITION_COL - 1]

        if email is None or disposition is None:
            continue

        email = str(email).strip()
        disposition = str(disposition).strip().lower()

        # Skip header row / anything that isn't an address or a valid bucket.
        if "@" not in email or disposition not in VALID_DISPOSITIONS:
            continue

        buckets[disposition].append(email)

    return buckets


# ---- Message building ----------------------------------------------------

def load_html(path):
    with open(path, "r", encoding="utf-8") as f:
        return f.read()


def build_message(html, sender, buckets, subject):
    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = sender
    if buckets["to"]:
        msg["To"] = ", ".join(buckets["to"])
    if buckets["cc"]:
        msg["Cc"] = ", ".join(buckets["cc"])
    # Bcc intentionally NOT added as a header (that's what makes it blind).
    msg["Date"] = formatdate(localtime=True)
    msg["Message-ID"] = make_msgid(domain="indegene.com")

    msg.set_content(
        "This message is best viewed in an HTML-capable email client."
    )
    msg.add_alternative(html, subtype="html")

    return msg


def send(msg, host, port, user, password, all_recipients):
    with smtplib.SMTP(host, port, timeout=60) as server:
        server.ehlo()
        server.starttls()          # upgrade to TLS on 587
        server.ehlo()
        server.login(user, password)
        # Pass the full envelope explicitly so bcc addresses are delivered
        # without ever appearing in the headers.
        server.send_message(msg, to_addrs=all_recipients)


# ---- Main ----------------------------------------------------------------

def main():
    args = sys.argv[1:]
    html_path = args[0] if len(args) > 0 else DEFAULT_HTML
    xlsx_path = args[1] if len(args) > 1 else DEFAULT_XLSX

    if not os.path.isfile(html_path):
        sys.exit(f"HTML file not found: {html_path}")
    if not os.path.isfile(xlsx_path):
        sys.exit(f"Excel file not found: {xlsx_path}")
    if not SMTP_PASS:
        sys.exit(
            "SMTP_PASS not set. Export the mailbox app password first:\n"
            '  export SMTP_PASS="your-app-password"'
        )

    buckets = load_recipients(xlsx_path)
    all_recipients = buckets["to"] + buckets["cc"] + buckets["bcc"]

    if not all_recipients:
        sys.exit(f"No valid recipients found in {xlsx_path}")
    if not buckets["to"]:
        sys.exit("Sheet has no 'to' recipient. At least one is required.")

    html = load_html(html_path)
    msg = build_message(html, SENDER, buckets, SUBJECT)

    print(f"Sending '{SUBJECT}'")
    print(f"  From: {SENDER}")
    print(f"  To:   {', '.join(buckets['to'])}")
    print(f"  Cc:   {', '.join(buckets['cc']) or '-'}")
    print(f"  Bcc:  {', '.join(buckets['bcc']) or '-'}")
    print(f"  HTML: {html_path} ({len(html):,} chars)")
    print(f"  Sheet: {xlsx_path}")

    try:
        send(msg, SMTP_HOST, SMTP_PORT, SMTP_USER, SMTP_PASS, all_recipients)
    except smtplib.SMTPAuthenticationError as e:
        sys.exit(
            f"Auth failed: {e}\n"
            "Office365 usually needs an app password and SMTP AUTH enabled "
            "for the mailbox (Microsoft 365 admin > Active users > Mail)."
        )
    except smtplib.SMTPException as e:
        sys.exit(f"SMTP error: {e}")

    print(f"Sent to {len(all_recipients)} recipient(s).")


if __name__ == "__main__":
    main()
