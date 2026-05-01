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

# In-memory email cache — 30-min TTL since fetching all emails takes time
_email_cache: dict = {"messages": [], "grouped": [], "ts": 0}
_CACHE_TTL = 1800  # seconds


def _get_emails_cached(limit: int = 10000) -> tuple[list, list]:
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


@app.route("/api/debug-sent")
def debug_sent():
    """Probe all recipient-access paths on sent items — tells us which JXA path works."""
    return jsonify(_outlook.debug_sent())


@app.route("/api/debug-groups")
def debug_groups():
    """Show from-address parsing for all messages — bypasses cache for fresh diagnostics."""
    _email_cache["ts"] = 0   # force fresh fetch
    messages, grouped = _get_emails_cached()
    own_domain = OUTLOOK_EMAIL.split("@")[-1].lower() if "@" in OUTLOOK_EMAIL else ""
    breakdown = {"empty_from": 0, "own_domain": 0, "external": 0, "domains": {}}
    for msg in messages:
        _, fa = parse_email_address(msg.get("from"))
        domain = fa.split("@")[-1].lower() if "@" in fa else ""
        if not domain:
            breakdown["empty_from"] += 1
        elif domain == own_domain:
            breakdown["own_domain"] += 1
        else:
            breakdown["external"] += 1
            breakdown["domains"][domain] = breakdown["domains"].get(domain, 0) + 1
    breakdown["domains"] = dict(sorted(breakdown["domains"].items(), key=lambda x: -x[1])[:20])
    return jsonify({
        "total_messages": len(messages),
        "grouped_count": len(grouped),
        "own_domain": own_domain,
        "breakdown": breakdown,
        "sample_from": [
            {"folder": m.get("folder"), "from": (m.get("from") or {}).get("emailAddress", {})}
            for m in messages[:5]
        ],
    })


@app.route("/api/test-ollama")
def test_ollama():
    from ai_classifier import test_ollama as _test
    ok, msg = _test()
    return jsonify({"ok": ok, "message": msg})

# ── Emails ─────────────────────────────────────────────────────────────────

@app.route("/api/emails")
def get_emails():
    try:
        messages, grouped = _get_emails_cached()
        return jsonify({
            "messages": messages[:60],   # display only
            "grouped":  grouped,
            "total":    len(messages),   # real count for the stat bar
        })
    except Exception as e:
        return jsonify({"error": str(e), "messages": [], "grouped": [], "total": 0}), 500


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
    from ai_classifier import classify_domain_emails, test_ollama

    def generate():
        try:
            # Pre-check: fail fast if Ollama is not reachable
            ok, ollama_msg = test_ollama()
            if not ok:
                yield _sse({"type": "fatal", "error": f"Ollama not ready — {ollama_msg}"})
                return

            yield _sse({"type": "status", "message": "Fetching emails from Outlook…"})

            # Force a fresh fetch so we always use current email data
            _email_cache["ts"] = 0
            messages, grouped = _get_emails_cached()

            print(f"[auto-classify] fetched {len(messages)} messages, {len(grouped)} groups")
            print(f"[auto-classify] own_domain filter: '{OUTLOOK_EMAIL.split('@')[-1].lower() if '@' in OUTLOOK_EMAIL else '<empty>'}'")
            # Sample first 10 from-addresses to diagnose grouping
            for m in messages[:10]:
                _, fa = parse_email_address(m.get("from"))
                print(f"[auto-classify]   sample from: '{fa}' folder={m.get('folder','')} subj={m.get('subject','')[:40]}")

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
            print(f"[auto-classify] classifying {total} domains with {OLLAMA_MODEL}")
            yield _sse({"type": "start", "total": total})

            classified = 0
            errors     = 0

            for i, group in enumerate(grouped):
                domain   = group["domain"]
                contacts = group.get("contacts", [])
                msgs     = domain_msgs.get(domain, [])

                print(f"[auto-classify] {i+1}/{total} {domain} ({len(msgs)} msgs)")
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
                    print(f"[auto-classify]   → {result.get('Company', domain)} | {result.get('Status', '?')}")

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
                    print(f"[auto-classify]   ✗ {domain}: {e}")
                    yield _sse({"type": "error", "domain": domain, "error": str(e)})
                    errors += 1

            print(f"[auto-classify] done — {classified} classified, {errors} errors")
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
