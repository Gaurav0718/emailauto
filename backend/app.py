"""
Flask API for the email sender.

Endpoints:
  GET  /api/health          -> liveness + default config
  POST /api/analyze         -> transform uploaded HTML, parse recipients
  POST /api/send            -> transform + send (mode: test | real)

Test mode always uses the bundled default test list, so a "Test Run"
works before the user has uploaded any recipient sheet.
"""

import os
import smtplib

from flask import Flask, request, jsonify
from flask_cors import CORS

import email_service as svc

app = Flask(__name__)

# Cap uploads (HTML + xlsx are small; block oversized bodies).
app.config["MAX_CONTENT_LENGTH"] = int(
    os.environ.get("MAX_UPLOAD_MB", "25")
) * 1024 * 1024

# Restrict CORS to the deployed frontend in production.
# CORS_ORIGINS = comma-separated list, or "*" for any (dev default).
_origins = os.environ.get("CORS_ORIGINS", "*").strip()
if _origins == "*":
    CORS(app)
else:
    CORS(app, origins=[o.strip() for o in _origins.split(",") if o.strip()])

HERE = os.path.dirname(os.path.abspath(__file__))
DEFAULT_TEST_LIST = os.path.join(HERE, "default_test_list.xlsx")


def _read_html(file_storage):
    raw = file_storage.read()
    try:
        return raw.decode("utf-8")
    except UnicodeDecodeError:
        return raw.decode("latin-1")


@app.get("/api/health")
def health():
    return jsonify(
        status="ok",
        sender=svc.DEFAULT_SENDER,
        subject=svc.DEFAULT_SUBJECT,
        smtp_host=svc.SMTP_HOST,
        smtp_port=svc.SMTP_PORT,
        default_test_list=os.path.basename(DEFAULT_TEST_LIST),
        default_test_list_exists=os.path.isfile(DEFAULT_TEST_LIST),
    )


@app.post("/api/analyze")
def analyze():
    """Transform the uploaded HTML and (if given) parse the recipient sheet.
    No email is sent. Powers the live preview + recipient breakdown.
    """
    if "html" not in request.files:
        return jsonify(error="No HTML file uploaded (field 'html')."), 400

    html_in = _read_html(request.files["html"])
    tagline = request.form.get("tagline")
    html_out, changes = svc.prepare_html(html_in, tagline)

    recipients = None
    recipient_error = None
    if "excel" in request.files and request.files["excel"].filename:
        try:
            buckets = svc.load_recipients_from_bytes(request.files["excel"].read())
            recipients = _summarize(buckets)
        except Exception as e:
            recipient_error = f"Could not read recipient sheet: {e}"

    return jsonify(
        changes=changes,
        transformed_html=html_out,
        recipients=recipients,
        recipient_error=recipient_error,
    )


@app.post("/api/send")
def send():
    """Transform then send.

    Form fields:
      html   (file, required)
      excel  (file, required for mode=real; ignored for mode=test)
      mode   (test | real, default test)
      subject, sender, smtp_user, smtp_pass
    """
    if "html" not in request.files:
        return jsonify(error="No HTML file uploaded (field 'html')."), 400

    mode = (request.form.get("mode") or "test").lower()
    subject = request.form.get("subject") or svc.DEFAULT_SUBJECT
    sender = request.form.get("sender") or svc.DEFAULT_SENDER
    smtp_user = request.form.get("smtp_user") or sender
    smtp_pass = request.form.get("smtp_pass") or ""

    if not smtp_pass:
        return jsonify(error="SMTP password is required to send."), 400

    html_in = _read_html(request.files["html"])
    tagline = request.form.get("tagline")
    html_out, changes = svc.prepare_html(html_in, tagline)

    # Pick the recipient source.
    try:
        if mode == "real":
            if "excel" not in request.files or not request.files["excel"].filename:
                return jsonify(error="Recipient sheet required for a real send."), 400
            buckets = svc.load_recipients_from_bytes(request.files["excel"].read())
        else:
            if not os.path.isfile(DEFAULT_TEST_LIST):
                return jsonify(error="Default test list is missing on the server."), 500
            buckets = svc.load_recipients_from_path(DEFAULT_TEST_LIST)
    except Exception as e:
        return jsonify(error=f"Could not read recipient sheet: {e}"), 400

    try:
        delivered = svc.send_email(
            html_out, buckets, subject, sender, smtp_user, smtp_pass
        )
    except ValueError as e:
        return jsonify(error=str(e)), 400
    except smtplib.SMTPAuthenticationError as e:
        return jsonify(
            error="Authentication failed. Office365 usually needs an app "
                  "password and SMTP AUTH enabled for the mailbox. "
                  f"({e.smtp_code})"
        ), 401
    except smtplib.SMTPException as e:
        return jsonify(error=f"SMTP error: {e}"), 502

    return jsonify(
        ok=True,
        mode=mode,
        changes=changes,
        subject=subject,
        sender=sender,
        recipients=_summarize(buckets),
        delivered_count=len(delivered),
    )


def _summarize(buckets):
    return {
        "to": buckets["to"],
        "cc": buckets["cc"],
        "bcc": buckets["bcc"],
        "total": len(buckets["to"]) + len(buckets["cc"]) + len(buckets["bcc"]),
    }


if __name__ == "__main__":
    # Local dev. In production run via gunicorn (see Procfile / Dockerfile).
    port = int(os.environ.get("PORT", "5001"))
    debug = os.environ.get("FLASK_DEBUG", "1") == "1"
    app.run(host="0.0.0.0", port=port, debug=debug)
