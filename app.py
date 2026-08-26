"""Flask backend — reads Outlook via macOS JXA, stores pipeline in local Excel."""
import json
import os
import time
from flask import Flask, jsonify, request, send_file, render_template, Response, stream_with_context
from flask_cors import CORS

import excel_manager as xl
from outlook_mac_client import (OutlookMacClient, external_domains_for_message,
                                external_participants_for_message,
                                group_messages_by_domain, parse_email_address)
from config.settings import (FLASK_SECRET_KEY, MANDATE_STATUSES, OUTLOOK_EMAIL,
                              EXCEL_FILE_PATH, OLLAMA_HOST, OLLAMA_MODEL,
                              EMAIL_CACHE_PATH, FOLLOW_UP_DAYS)

app = Flask(__name__)
app.secret_key = FLASK_SECRET_KEY
CORS(app)

_outlook = OutlookMacClient()

# ── Email cache (memory + disk) ────────────────────────────────────────────
# On startup: load from disk so Outlook doesn't need to be re-queried.
# On refresh: fetch fresh from Outlook and overwrite the disk file.
EMAIL_CACHE_VERSION = 1
_email_cache: dict = {"messages": [], "grouped": [], "ts": 0}


def _cache_meta() -> dict:
    ts = _email_cache["ts"]
    return {
        "cached_at": ts,
        "age_seconds": int(time.time() - ts) if ts else None,
        "total": len(_email_cache["messages"]),
        "groups": len(_email_cache["grouped"]),
        "cache_path": EMAIL_CACHE_PATH,
        "disk_exists": os.path.exists(EMAIL_CACHE_PATH),
        "needs_refresh": not bool(_email_cache["messages"]),
    }


def _save_disk_cache() -> None:
    tmp_path = f"{EMAIL_CACHE_PATH}.{os.getpid()}.tmp"
    try:
        cache_dir = os.path.dirname(os.path.abspath(EMAIL_CACHE_PATH))
        if cache_dir:
            os.makedirs(cache_dir, exist_ok=True)
        with open(tmp_path, "w", encoding="utf-8") as f:
            json.dump({
                "version":  EMAIL_CACHE_VERSION,
                "ts":       _email_cache["ts"],
                "messages": _email_cache["messages"],
            }, f)
            f.flush()
            os.fsync(f.fileno())
        os.replace(tmp_path, EMAIL_CACHE_PATH)
        print(f"[cache] saved {len(_email_cache['messages'])} messages → {EMAIL_CACHE_PATH}")
    except Exception as e:
        if os.path.exists(tmp_path):
            try:
                os.remove(tmp_path)
            except OSError:
                pass
        print(f"[cache] save failed: {e}")


def _load_disk_cache() -> bool:
    try:
        if not os.path.exists(EMAIL_CACHE_PATH):
            return False
        with open(EMAIL_CACHE_PATH, "r", encoding="utf-8") as f:
            data = json.load(f)
        if not isinstance(data, dict):
            raise ValueError("email cache is not a JSON object")
        messages = data.get("messages", [])
        if not isinstance(messages, list):
            raise ValueError("email cache messages must be a list")
        _email_cache.update({
            "messages": messages,
            "grouped":  group_messages_by_domain(messages),
            "ts":       data.get("ts", 0),
        })
        age_h = (time.time() - data.get("ts", 0)) / 3600 if data.get("ts") else 0
        print(f"[cache] loaded {len(messages)} messages from disk (age: {age_h:.1f}h)")
        return True
    except Exception as e:
        print(f"[cache] load failed: {e}")
        return False


def _get_emails_cached(force: bool = False, allow_fetch: bool = False) -> tuple[list, list]:
    if not force and _email_cache["messages"]:
        return _email_cache["messages"], _email_cache["grouped"]
    if not force and not allow_fetch:
        return _email_cache["messages"], _email_cache["grouped"]
    messages = _outlook.get_all_messages()
    grouped  = group_messages_by_domain(messages)
    _email_cache.update({"messages": messages, "grouped": grouped, "ts": time.time()})
    _save_disk_cache()
    return messages, grouped


def _domain_for_message(msg: dict) -> str:
    domains = external_domains_for_message(msg)
    if domains:
        return domains[0]
    _, addr = parse_email_address(msg.get("from"))
    return addr.split("@")[-1].lower() if "@" in addr else ""


def _cached_messages_for_domain(domain: str, allow_fetch: bool = False) -> tuple[list, dict]:
    domain = (domain or "").lower().strip()
    if not domain:
        return [], {}
    messages, grouped = _get_emails_cached(allow_fetch=allow_fetch)
    group = next((g for g in grouped if g.get("domain") == domain), {})
    return [m for m in messages if domain in external_domains_for_message(m)], group


def _normalize_key(value: str) -> str:
    return "".join(ch for ch in str(value or "").lower() if ch.isalnum())


def _company_tokens(value: str) -> list[str]:
    generic = {
        "asset", "assets", "management", "capital", "partners", "partner",
        "group", "holdings", "holding", "limited", "ltd", "llc", "inc", "plc",
        "fund", "funds", "ventures", "venture", "family", "office",
        "investment", "investments",
    }
    tokens = [
        token for token in "".join(
            ch.lower() if ch.isalnum() else " " for ch in str(value or "")
        ).split()
        if len(token) > 2 and token not in generic
    ]
    expanded = set(tokens)
    for token in tokens:
        if token.endswith("ical") and len(token) > 6:
            expanded.add(token[:-4])
        if token.endswith("medical") and len(token) > 8:
            expanded.add(token[:-7] + "med")
    return list(expanded)


def _cached_thread_messages(domain: str = "", email: str = "", company: str = "") -> tuple[list, dict]:
    domain = (domain or "").lower().strip()
    email = (email or "").lower().strip()
    if not domain and "@" in email:
        domain = email.split("@")[-1]

    if domain:
        return _cached_messages_for_domain(domain, allow_fetch=False)

    messages, _ = _get_emails_cached(allow_fetch=False)
    tokens = _company_tokens(company)
    if not tokens:
        return [], {}

    matched = []
    for msg in messages:
        people = external_participants_for_message(msg)
        haystack_parts = [_normalize_key(msg.get("subject", ""))]
        for person in people:
            name, addr = parse_email_address(person)
            haystack_parts.extend([
                _normalize_key(name),
                _normalize_key(addr),
                _normalize_key((addr.split("@")[-1] if "@" in addr else "").split(".")[0]),
            ])
        haystack = " ".join(haystack_parts)
        if any(token in haystack for token in tokens):
            matched.append(msg)
    return matched, {}


# Load disk cache immediately — no Outlook query needed on startup
_load_disk_cache()


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
        "follow_up_days": FOLLOW_UP_DAYS,
        "email_cache": _cache_meta(),
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
    """Show address parsing for cached messages without querying Outlook."""
    messages, grouped = _get_emails_cached()
    own_domain = OUTLOOK_EMAIL.split("@")[-1].lower() if "@" in OUTLOOK_EMAIL else ""
    breakdown = {"empty_from": 0, "own_domain": 0, "external": 0, "domains": {}}
    for msg in messages:
        domains = external_domains_for_message(msg)
        if not domains:
            breakdown["empty_from"] += 1
        else:
            breakdown["external"] += 1
            for domain in domains:
                if domain == own_domain:
                    breakdown["own_domain"] += 1
                    continue
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
        force = request.args.get("refresh") == "true"
        messages, grouped = _get_emails_cached(force=force, allow_fetch=force)
        meta = _cache_meta()
        return jsonify({
            "messages":  messages[:60],
            "grouped":   grouped,
            "total":     len(messages),
            "cached_at": meta["cached_at"],
            "cache_path": meta["cache_path"],
            "disk_exists": meta["disk_exists"],
            "needs_refresh": meta["needs_refresh"],
        })
    except Exception as e:
        return jsonify({"error": str(e), "messages": [], "grouped": [], "total": 0}), 500


@app.route("/api/cache-info")
def cache_info():
    return jsonify(_cache_meta())


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


@app.route("/api/emails/thread")
def get_cached_email_thread():
    domain = request.args.get("domain", "").strip()
    email = request.args.get("email", "").strip()
    company = request.args.get("company", "").strip()
    try:
        limit = min(max(int(request.args.get("limit", 120)), 1), 300)
    except ValueError:
        limit = 120

    if not domain and not email and not company:
        return jsonify({"error": "domain, email, or company required", "messages": []}), 400

    try:
        messages, group = _cached_thread_messages(domain=domain, email=email, company=company)
        messages = sorted(messages, key=lambda m: m.get("receivedDateTime", ""), reverse=True)
        return jsonify({
            "messages":  messages[:limit],
            "total":     len(messages),
            "group":     group,
            "source":    "cache",
            "cached_at": _email_cache["ts"],
            "cache_path": EMAIL_CACHE_PATH,
            "disk_exists": os.path.exists(EMAIL_CACHE_PATH),
            "needs_refresh": not bool(_email_cache["messages"]),
        })
    except Exception as e:
        return jsonify({"error": str(e), "messages": [], "total": 0}), 500


@app.route("/api/emails/message", methods=["POST"])
def get_email_message():
    data = request.get_json() or {}
    message_id = str(data.get("id") or "").strip()
    if not message_id:
        return jsonify({"error": "message id required"}), 400

    cached = next(
        (m for m in _email_cache.get("messages", []) if str(m.get("id") or "") == message_id),
        None,
    )
    if cached and cached.get("body"):
        return jsonify({"message": cached, "source": "cache"})

    try:
        message = _outlook.get_message(message_id)
        if not message:
            if cached:
                return jsonify({"message": cached, "source": "cache"})
            return jsonify({"error": "message not found"}), 404
        return jsonify({"message": message, "source": "outlook"})
    except Exception as e:
        if cached:
            cached = dict(cached)
            cached.setdefault("body", {"contentType": "text", "content": cached.get("bodyPreview", "")})
            return jsonify({"message": cached, "source": "cache", "warning": str(e)})
        return jsonify({"error": str(e)}), 500


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
    if data.get("Status") and data.get("Status") not in MANDATE_STATUSES:
        return jsonify({"error": "Invalid status"}), 400
    try:
        xl.upsert_company(data)
        return jsonify({"success": True})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/companies/<path:company_name>/details", methods=["PUT"])
def update_company(company_name):
    data = request.get_json()
    if not data or not data.get("Company"):
        return jsonify({"error": "Company name required"}), 400
    if data.get("Status") and data.get("Status") not in MANDATE_STATUSES:
        return jsonify({"error": "Invalid status"}), 400
    try:
        xl.upsert_company(data, match_name=company_name)
        return jsonify({"success": True})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/companies/<path:company_name>/status", methods=["PUT"])
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


@app.route("/api/companies/<path:company_name>/log-emails", methods=["POST"])
def log_company_emails(company_name):
    data    = request.get_json()
    emails  = data.get("emails", [])
    try:
        added = xl.log_emails(emails, company_name)
        return jsonify({"added": added})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/deals/classify", methods=["POST"])
def classify_deal():
    from ai_classifier import classify_domain_emails, test_ollama

    data = request.get_json() or {}
    domain = (data.get("domain") or "").strip().lower()
    emails = data.get("emails") if isinstance(data.get("emails"), list) else []
    contacts = data.get("contacts") if isinstance(data.get("contacts"), list) else []

    if not emails and domain:
        emails, group = _cached_messages_for_domain(domain, allow_fetch=False)
        contacts = contacts or group.get("contacts", [])

    if not domain:
        for msg in emails:
            domain = _domain_for_message(msg)
            if domain:
                break

    if not domain and not data.get("company_name"):
        return jsonify({"error": "Domain or company name required"}), 400
    if not emails:
        return jsonify({"error": "No cached thread emails found. Refresh emails or run Search Outlook first."}), 400

    ok, ollama_msg = test_ollama()
    if not ok:
        return jsonify({"error": f"Ollama not ready — {ollama_msg}"}), 503

    try:
        result = classify_domain_emails(domain or data.get("company_name"), emails, contacts)
        if domain:
            result["_Domain"] = domain
        if data.get("last_email_date") and not result.get("Last Email Date"):
            result["Last Email Date"] = str(data["last_email_date"])[:10]
        return jsonify({"result": result, "emails_used": len(emails)})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

# ── Auto-classify (SSE streaming) ──────────────────────────────────────────

@app.route("/api/auto-classify", methods=["POST"])
def auto_classify():
    from ai_classifier import classify_domain_emails, test_ollama

    def generate():
        try:
            yield _sse({"type": "status", "message": "Loading emails from cache…"})

            messages, grouped = _get_emails_cached()

            if not messages:
                yield _sse({
                    "type": "done",
                    "classified": 0,
                    "errors": 0,
                    "needs_refresh": True,
                    "message": "No local email cache found. Click Refresh Emails first, then run auto-classify again.",
                })
                return

            # Pre-check: fail fast if Ollama is not reachable after confirming cached emails exist.
            ok, ollama_msg = test_ollama()
            if not ok:
                yield _sse({"type": "fatal", "error": f"Ollama not ready — {ollama_msg}"})
                return

            print(f"[auto-classify] fetched {len(messages)} messages, {len(grouped)} groups")
            print(f"[auto-classify] own_domain filter: '{OUTLOOK_EMAIL.split('@')[-1].lower() if '@' in OUTLOOK_EMAIL else '<empty>'}'")
            # Sample first 10 from-addresses to diagnose grouping
            for m in messages[:10]:
                _, fa = parse_email_address(m.get("from"))
                print(f"[auto-classify]   sample from: '{fa}' folder={m.get('folder','')} subj={m.get('subject','')[:40]}")

            if not grouped:
                yield _sse({"type": "done", "classified": 0, "errors": 0,
                            "message": "No email groups found in the local cache. Click Refresh Emails to rebuild it."})
                return

            # Build domain → messages lookup in one pass
            own_domain = OUTLOOK_EMAIL.split("@")[-1].lower() if "@" in OUTLOOK_EMAIL else ""
            domain_msgs: dict[str, list] = {}
            for msg in messages:
                for d in external_domains_for_message(msg):
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
                    result["_Domain"] = domain

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
        "follow_up_days": FOLLOW_UP_DAYS,
    })


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/deal/company/<path:deal_key>")
@app.route("/deal/domain/<path:deal_key>")
def deal_page(deal_key):
    return render_template("index.html")


if __name__ == "__main__":
    app.run(debug=True, port=5000)
