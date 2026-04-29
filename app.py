"""Flask backend — reads Outlook via macOS JXA, stores pipeline in local Excel."""
import json
import time
from flask import Flask, jsonify, request, send_file, render_template, Response, stream_with_context
from flask_cors import CORS

import excel_manager as xl
from outlook_mac_client import OutlookMacClient, group_messages_by_domain, parse_email_address
from config.settings import FLASK_SECRET_KEY, MANDATE_STATUSES, OUTLOOK_EMAIL, EXCEL_FILE_PATH, OLLAMA_HOST, OLLAMA_MODEL

app = Flask(__name__)
app.secret_key = FLASK_SECRET_KEY
CORS(app)

_outlook = OutlookMacClient()

# Simple in-memory email cache (5-minute TTL) to avoid re-running JXA on auto-classify
_email_cache: dict = {"messages": [], "grouped": [], "ts": 0}
_CACHE_TTL = 300  # seconds


def _get_emails_cached(limit: int = 200) -> tuple[list, list]:
    if time.time() - _email_cache["ts"] < _CACHE_TTL and _email_cache["messages"]:
        return _email_cache["messages"], _email_cache["grouped"]
    messages = _outlook.get_all_messages(limit)
    grouped  = group_messages_by_domain(messages)
    _email_cache.update({"messages": messages, "grouped": grouped, "ts": time.time()})
    return messages, grouped


def _sse(payload: dict) -> str:
    return f"data: {json.dumps(payload)}\n\n"


# ── Status ─────────────────────────────────────────────────────────────────

@app.route("/api/status")
def api_status():
    return jsonify({
        "configured":  True,
        "email":       OUTLOOK_EMAIL,
        "imap_server": _outlook._server,
        "excel_path":  xl.get_excel_path(),
    })


@app.route("/api/test-connection")
def test_connection():
    ok, msg = _outlook.test_connection()
    return jsonify({"ok": ok, "message": msg})


@app.route("/api/debug-outlook")
def debug_outlook():
    """Raw JXA diagnostic — visit this URL to see exactly what Outlook exposes."""
    return jsonify(_outlook.raw_diagnostic())


@app.route("/api/debug-emails")
def debug_emails():
    """First 5 inbox messages with all sender-access paths — confirms which path returns email addresses."""
    return jsonify(_outlook.debug_messages())


@app.route("/api/test-ollama")
def test_ollama():
    from ai_classifier import test_ollama as _test
    ok, msg = _test()
    return jsonify({"ok": ok, "message": msg})

# ── Emails ─────────────────────────────────────────────────────────────────

@app.route("/api/emails")
def get_emails():
    try:
        limit    = int(request.args.get("top", 150))
        messages, grouped = _get_emails_cached(limit)
        return jsonify({"messages": messages[:60], "grouped": grouped})
    except Exception as e:
        return jsonify({"error": str(e), "messages": [], "grouped": []}), 500


@app.route("/api/emails/search")
def search_emails():
    query = request.args.get("q", "").strip()
    if not query:
        return jsonify({"error": "query required"}), 400
    try:
        messages = _outlook.search_messages(query)
        return jsonify({"messages": messages})
    except Exception as e:
        return jsonify({"error": str(e), "messages": []}), 500


@app.route("/api/folders")
def list_folders():
    try:
        folders = _outlook.get_folders()
        return jsonify({"folders": folders})
    except Exception as e:
        return jsonify({"error": str(e), "folders": []}), 500

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
    data    = request.get_json()
    emails  = data.get("emails", [])
    try:
        added = xl.log_emails(emails, company_name)
        return jsonify({"added": added})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

# ── Auto-classify (SSE streaming) ──────────────────────────────────────────

@app.route("/api/auto-classify", methods=["POST"])
def auto_classify():
    from ai_classifier import classify_domain_emails

    def generate():
        try:
            yield _sse({"type": "status", "message": "Fetching emails from Outlook…"})

            messages, grouped = _get_emails_cached(200)

            if not grouped:
                yield _sse({"type": "done", "classified": 0, "errors": 0,
                            "message": "No email groups found. Make sure Outlook is open."})
                return

            # Build domain → messages lookup in one pass
            own_domain = OUTLOOK_EMAIL.split("@")[-1].lower() if "@" in OUTLOOK_EMAIL else ""
            domain_msgs: dict[str, list] = {}
            for msg in messages:
                _, from_addr = parse_email_address(msg.get("from"))
                d = from_addr.split("@")[-1].lower() if "@" in from_addr else ""
                if not d or d == own_domain:
                    continue
                domain_msgs.setdefault(d, []).append(msg)

            total = len(grouped)
            yield _sse({"type": "start", "total": total})

            classified = 0
            errors     = 0

            for i, group in enumerate(grouped):
                domain   = group["domain"]
                contacts = group.get("contacts", [])
                msgs     = domain_msgs.get(domain, [])

                yield _sse({
                    "type":   "progress",
                    "domain": domain,
                    "index":  i + 1,
                    "total":  total,
                })

                try:
                    result = classify_domain_emails(domain, msgs, contacts)

                    # Carry over last email date from grouped data
                    if group.get("last_email_date") and not result.get("Last Email Date"):
                        result["Last Email Date"] = group["last_email_date"][:10]

                    xl.upsert_company(result)

                    yield _sse({
                        "type":    "result",
                        "domain":  domain,
                        "company": result.get("Company", domain),
                        "status":  result.get("Status", ""),
                        "contact": result.get("Contact Name", ""),
                        "saved":   True,
                    })
                    classified += 1

                except Exception as e:
                    yield _sse({"type": "error", "domain": domain, "error": str(e)})
                    errors += 1

            yield _sse({"type": "done", "classified": classified, "errors": errors})

        except Exception as e:
            yield _sse({"type": "fatal", "error": str(e)})

    return Response(
        stream_with_context(generate()),
        mimetype="text/event-stream",
        headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"},
    )

# ── Excel download ─────────────────────────────────────────────────────────

@app.route("/api/excel/download")
def download_excel():
    try:
        return send_file(
            xl.get_excel_path(),
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


@app.route("/api/config")
def get_config():
    return jsonify({
        "ollama_host":  OLLAMA_HOST,
        "ollama_model": OLLAMA_MODEL,
    })


@app.route("/")
def index():
    return render_template("index.html")


if __name__ == "__main__":
    app.run(debug=True, port=5000)
