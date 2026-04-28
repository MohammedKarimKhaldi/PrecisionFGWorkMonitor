"""Flask backend for Fundraising Email Monitor."""
from flask import Flask, jsonify, request, send_file, render_template, session, redirect, url_for
from flask_cors import CORS
import msal
import io
import requests as http_requests
from datetime import datetime

from config.settings import (
    AZURE_CLIENT_ID, AZURE_CLIENT_SECRET, AZURE_TENANT_ID,
    USER_EMAIL, FLASK_SECRET_KEY, MANDATE_STATUSES, GRAPH_API_BASE,
)
from outlook_client import OutlookClient, group_messages_by_domain, parse_email_address
from excel_manager import ExcelManager

app = Flask(__name__)
app.secret_key = FLASK_SECRET_KEY
CORS(app)

# OAuth scopes for delegated (user) auth flow
OAUTH_SCOPES = [
    "Mail.Read",
    "Mail.ReadBasic",
    "Files.ReadWrite",
    "offline_access",
    "User.Read",
]

REDIRECT_URI = "http://localhost:5000/auth/callback"

# ---------------------------------------------------------------------------
# Auth helpers
# ---------------------------------------------------------------------------

def _msal_app():
    return msal.ConfidentialClientApplication(
        AZURE_CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{AZURE_TENANT_ID}",
        client_credential=AZURE_CLIENT_SECRET,
    )


def _get_access_token():
    """Return a valid access token from session cache, refreshing if needed."""
    cache = msal.SerializableTokenCache()
    if "token_cache" in session:
        cache.deserialize(session["token_cache"])

    msal_app = msal.ConfidentialClientApplication(
        AZURE_CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{AZURE_TENANT_ID}",
        client_credential=AZURE_CLIENT_SECRET,
        token_cache=cache,
    )
    accounts = msal_app.get_accounts()
    result = None
    if accounts:
        result = msal_app.acquire_token_silent(OAUTH_SCOPES, account=accounts[0])

    if not result or "access_token" not in result:
        return None

    if cache.has_state_changed:
        session["token_cache"] = cache.serialize()
    return result["access_token"]


def _auth_headers(token):
    return {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}


def _graph_get(token, url, params=None):
    resp = http_requests.get(url, headers=_auth_headers(token), params=params, timeout=30)
    resp.raise_for_status()
    return resp.json()


def _require_auth(f):
    from functools import wraps
    @wraps(f)
    def decorated(*args, **kwargs):
        token = _get_access_token()
        if not token:
            return jsonify({"error": "not_authenticated", "login_url": url_for("login")}), 401
        return f(*args, token=token, **kwargs)
    return decorated

# ---------------------------------------------------------------------------
# OAuth Routes
# ---------------------------------------------------------------------------

@app.route("/auth/login")
def login():
    msal_app = _msal_app()
    auth_url = msal_app.get_authorization_request_url(
        OAUTH_SCOPES,
        redirect_uri=REDIRECT_URI,
        state="random-state",
    )
    return redirect(auth_url)


@app.route("/auth/callback")
def auth_callback():
    code = request.args.get("code")
    if not code:
        return "Authentication failed: no code returned.", 400

    cache = msal.SerializableTokenCache()
    msal_app = msal.ConfidentialClientApplication(
        AZURE_CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{AZURE_TENANT_ID}",
        client_credential=AZURE_CLIENT_SECRET,
        token_cache=cache,
    )
    result = msal_app.acquire_token_by_authorization_code(
        code, scopes=OAUTH_SCOPES, redirect_uri=REDIRECT_URI
    )
    if "error" in result:
        return f"Auth error: {result.get('error_description')}", 400

    session["token_cache"] = cache.serialize()
    session["user_name"] = result.get("id_token_claims", {}).get("name", "")
    session["user_email"] = result.get("id_token_claims", {}).get("preferred_username", USER_EMAIL)
    return redirect(url_for("index"))


@app.route("/auth/logout")
def logout():
    session.clear()
    return redirect(url_for("index"))


@app.route("/auth/me")
def me():
    token = _get_access_token()
    if not token:
        return jsonify({"authenticated": False})
    return jsonify({
        "authenticated": True,
        "name": session.get("user_name", ""),
        "email": session.get("user_email", USER_EMAIL),
    })

# ---------------------------------------------------------------------------
# Main UI
# ---------------------------------------------------------------------------

@app.route("/")
def index():
    return render_template("index.html")

# ---------------------------------------------------------------------------
# API: Emails
# ---------------------------------------------------------------------------

@app.route("/api/emails")
@_require_auth
def get_emails(token):
    folder = request.args.get("folder", "inbox")
    top = int(request.args.get("top", 200))

    params = {
        "$top": top,
        "$select": "id,subject,from,toRecipients,receivedDateTime,bodyPreview,conversationId,isRead",
        "$orderby": "receivedDateTime desc",
    }

    user_email = session.get("user_email", USER_EMAIL)

    if folder == "sent":
        url = f"{GRAPH_API_BASE}/users/{user_email}/mailFolders/sentitems/messages"
    else:
        url = f"{GRAPH_API_BASE}/users/{user_email}/mailFolders/inbox/messages"

    data = _graph_get(token, url, params)
    messages = data.get("value", [])

    grouped = group_messages_by_domain(messages)
    return jsonify({"messages": messages[:50], "grouped": grouped})


@app.route("/api/emails/search")
@_require_auth
def search_emails(token):
    query = request.args.get("q", "")
    if not query:
        return jsonify({"error": "query required"}), 400

    user_email = session.get("user_email", USER_EMAIL)
    url = f"{GRAPH_API_BASE}/users/{user_email}/messages"
    params = {
        "$search": f'"{query}"',
        "$top": 50,
        "$select": "id,subject,from,toRecipients,receivedDateTime,bodyPreview,conversationId,isRead",
    }
    data = _graph_get(token, url, params)
    return jsonify({"messages": data.get("value", [])})


@app.route("/api/emails/conversation/<conversation_id>")
@_require_auth
def get_conversation(conversation_id, token):
    user_email = session.get("user_email", USER_EMAIL)
    url = f"{GRAPH_API_BASE}/users/{user_email}/messages"
    params = {
        "$filter": f"conversationId eq '{conversation_id}'",
        "$select": "id,subject,from,toRecipients,receivedDateTime,bodyPreview,isRead",
        "$orderby": "receivedDateTime asc",
        "$top": 50,
    }
    data = _graph_get(token, url, params)
    return jsonify({"messages": data.get("value", [])})

# ---------------------------------------------------------------------------
# API: Companies (Excel-backed)
# ---------------------------------------------------------------------------

def _excel_manager(token):
    class _FakeClient:
        def _headers(self_inner):
            return _auth_headers(token)
    mgr = ExcelManager()
    mgr._graph_client = _FakeClient()
    return mgr


@app.route("/api/companies", methods=["GET"])
@_require_auth
def get_companies(token):
    try:
        mgr = _excel_manager(token)
        companies = mgr.get_companies()
        return jsonify({"companies": companies})
    except Exception as e:
        return jsonify({"error": str(e), "companies": []}), 500


@app.route("/api/companies", methods=["POST"])
@_require_auth
def create_company(token):
    data = request.get_json()
    if not data or not data.get("Company"):
        return jsonify({"error": "Company name required"}), 400
    try:
        mgr = _excel_manager(token)
        mgr.upsert_company(data)
        return jsonify({"success": True})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/companies/<company_name>/status", methods=["PUT"])
@_require_auth
def update_status(company_name, token):
    data = request.get_json()
    new_status = data.get("status")
    notes = data.get("notes")
    if not new_status or new_status not in MANDATE_STATUSES:
        return jsonify({"error": "Invalid status"}), 400
    try:
        mgr = _excel_manager(token)
        mgr.update_company_status(company_name, new_status, notes)
        return jsonify({"success": True})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/companies/<company_name>/log-emails", methods=["POST"])
@_require_auth
def log_company_emails(company_name, token):
    data = request.get_json()
    emails = data.get("emails", [])
    if not emails:
        return jsonify({"added": 0})
    try:
        mgr = _excel_manager(token)
        added = mgr.log_emails(emails, company_name)
        return jsonify({"added": added})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

# ---------------------------------------------------------------------------
# API: Excel download
# ---------------------------------------------------------------------------

@app.route("/api/excel/download")
@_require_auth
def download_excel(token):
    try:
        mgr = _excel_manager(token)
        content = mgr.get_workbook_bytes()
        return send_file(
            io.BytesIO(content),
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name="FundraisingTracker.xlsx",
        )
    except Exception as e:
        return jsonify({"error": str(e)}), 500

# ---------------------------------------------------------------------------
# API: Status options
# ---------------------------------------------------------------------------

@app.route("/api/statuses")
def get_statuses():
    return jsonify({"statuses": MANDATE_STATUSES})

# ---------------------------------------------------------------------------

if __name__ == "__main__":
    app.run(debug=True, port=5000)
