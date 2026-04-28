"""Flask backend — reads Outlook via IMAP, stores pipeline in local Excel."""
from flask import Flask, jsonify, request, send_file, render_template
from flask_cors import CORS

import excel_manager as xl
from outlook_client import OutlookIMAPClient, group_messages_by_domain
from config.settings import FLASK_SECRET_KEY, MANDATE_STATUSES, OUTLOOK_EMAIL, EXCEL_FILE_PATH

app = Flask(__name__)
app.secret_key = FLASK_SECRET_KEY
CORS(app)

_imap = OutlookIMAPClient()

# ── Health / connection check ──────────────────────────────────────────────

@app.route("/api/status")
def api_status():
    return jsonify({
        "configured":  bool(OUTLOOK_EMAIL),
        "email":       OUTLOOK_EMAIL,
        "excel_path":  xl.get_excel_path(),
    })


@app.route("/api/test-connection")
def test_connection():
    ok, msg = _imap.test_connection()
    return jsonify({"ok": ok, "message": msg})

# ── Emails ─────────────────────────────────────────────────────────────────

@app.route("/api/emails")
def get_emails():
    try:
        folder = request.args.get("folder", "all")
        limit  = int(request.args.get("top", 150))
        if folder == "inbox":
            messages = _imap.get_inbox(limit)
        elif folder == "sent":
            messages = _imap.get_sent(limit)
        else:
            messages = _imap.get_all_messages(limit)
        grouped = group_messages_by_domain(messages)
        return jsonify({"messages": messages[:60], "grouped": grouped})
    except Exception as e:
        return jsonify({"error": str(e), "messages": [], "grouped": []}), 500


@app.route("/api/emails/search")
def search_emails():
    query = request.args.get("q", "").strip()
    if not query:
        return jsonify({"error": "query required"}), 400
    try:
        messages = _imap.search_messages(query)
        return jsonify({"messages": messages})
    except Exception as e:
        return jsonify({"error": str(e), "messages": []}), 500

# ── Companies ──────────────────────────────────────────────────────────────

@app.route("/api/companies", methods=["GET"])
def get_companies():
    try:
        return jsonify({"companies": xl.get_companies()})
    except Exception as e:
        return jsonify({"error": str(e), "companies": []}), 500


@app.route("/api/companies", methods=["POST"])
def create_or_update_company():
    data = request.get_json()
    if not data or not data.get("Company"):
        return jsonify({"error": "Company name required"}), 400
    try:
        xl.upsert_company(data)
        return jsonify({"success": True})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/companies/<company_name>/status", methods=["PUT"])
def update_status(company_name):
    data = request.get_json()
    new_status = data.get("status")
    if not new_status or new_status not in MANDATE_STATUSES:
        return jsonify({"error": "Invalid status"}), 400
    try:
        xl.update_company_status(company_name, new_status, data.get("notes"))
        return jsonify({"success": True})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/companies/<company_name>/log-emails", methods=["POST"])
def log_company_emails(company_name):
    data = request.get_json()
    emails = data.get("emails", [])
    try:
        added = xl.log_emails(emails, company_name)
        return jsonify({"added": added})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

# ── Excel download ─────────────────────────────────────────────────────────

@app.route("/api/excel/download")
def download_excel():
    try:
        path = xl.get_excel_path()
        return send_file(
            path,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name="FundraisingTracker.xlsx",
        )
    except Exception as e:
        return jsonify({"error": str(e)}), 500

# ── Meta ───────────────────────────────────────────────────────────────────

@app.route("/api/statuses")
def get_statuses():
    return jsonify({"statuses": MANDATE_STATUSES})


@app.route("/")
def index():
    return render_template("index.html")


if __name__ == "__main__":
    app.run(debug=True, port=5000)
