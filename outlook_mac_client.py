"""Read Outlook emails on macOS via JXA (JavaScript for Automation).
No IMAP, no passwords — talks directly to the running Outlook desktop app.
"""
import subprocess
import json

from config.settings import OUTLOOK_EMAIL

_FETCH_LIMIT = 150
_MONTHS_BACK = 6


def _run_jxa(script: str, timeout: int = 90) -> any:
    result = subprocess.run(
        ["osascript", "-l", "JavaScript", "-e", script],
        capture_output=True, text=True, timeout=timeout,
    )
    if result.returncode != 0:
        raise RuntimeError(result.stderr.strip() or "JXA script failed")
    out = result.stdout.strip()
    if not out or out == "null":
        return []
    return json.loads(out)


# ── JXA Scripts ───────────────────────────────────────────────────────────

# Fast header-only fetch from Inbox + Sent (no body = much faster)
_SCRIPT_GET_HEADERS = """\
(function() {
    var app = Application('Microsoft Outlook');
    var result = [];
    var limit = LIMIT;
    var cutoffMs = Date.now() - MONTHS * 30 * 24 * 60 * 60 * 1000;

    function addMsg(m, folderName) {
        try {
            var d = m.timeReceived();
            if (!d || d.getTime() < cutoffMs) return;
            var fromAddr = '', fromName = '';
            try { fromAddr = m.sender.emailAddress() || ''; } catch(e) {}
            try { fromName = m.sender.name() || ''; } catch(e) {}
            result.push({
                id:       String(m.id()),
                subject:  m.subject() || '',
                fromAddr: fromAddr,
                fromName: fromName,
                date:     d.toISOString(),
                preview:  '',
                folder:   folderName
            });
        } catch(e) {}
    }

    function fetchFolder(folder, name) {
        try {
            var msgs = folder.messages();
            var start = Math.max(0, msgs.length - limit);
            for (var i = msgs.length - 1; i >= start; i--) { addMsg(msgs[i], name); }
        } catch(e) {}
    }

    try { fetchFolder(app.inbox, 'Inbox'); } catch(e) {}

    // Find Sent folder by name (handles locale variants)
    try {
        var folders = app.mailFolders();
        for (var i = 0; i < folders.length; i++) {
            try {
                var n = folders[i].name().toLowerCase();
                if (n === 'sent items' || n === 'sent mail' || n === 'sent') {
                    fetchFolder(folders[i], 'Sent');
                    break;
                }
            } catch(e) {}
        }
    } catch(e) {}

    result.sort(function(a, b) { return b.date.localeCompare(a.date); });
    return JSON.stringify(result);
})()
"""

# Full search across ALL folders recursively (includes body preview)
_SCRIPT_SEARCH = """\
(function() {
    var app = Application('Microsoft Outlook');
    var query = QUERY_JSON;
    var result = [];
    var cutoffMs = Date.now() - 365 * 24 * 60 * 60 * 1000;

    function matchMsg(m, folderName) {
        try {
            var fromAddr = '', fromName = '';
            try { fromAddr = m.sender.emailAddress() || ''; } catch(e) {}
            try { fromName = m.sender.name() || ''; } catch(e) {}
            var subj = (m.subject() || '').toLowerCase();
            var q    = query.toLowerCase();
            if (subj.indexOf(q) < 0 && fromAddr.toLowerCase().indexOf(q) < 0
                && fromName.toLowerCase().indexOf(q) < 0) return;
            var d = m.timeReceived();
            if (d && d.getTime() < cutoffMs) return;
            var preview = '';
            try { preview = (m.plainTextContent() || '').replace(/\\s+/g,' ').substring(0, 280); } catch(e) {}
            result.push({
                id:       String(m.id()),
                subject:  m.subject() || '',
                fromAddr: fromAddr,
                fromName: fromName,
                date:     d ? d.toISOString() : '',
                preview:  preview,
                folder:   folderName
            });
        } catch(e) {}
    }

    function searchFolder(folder, name) {
        try {
            var msgs = folder.messages();
            for (var i = 0; i < msgs.length; i++) { matchMsg(msgs[i], name); }
        } catch(e) {}
        try {
            var sub = folder.mailFolders();
            for (var j = 0; j < sub.length; j++) {
                try { searchFolder(sub[j], sub[j].name()); } catch(e) {}
            }
        } catch(e) {}
    }

    try { searchFolder(app.inbox, 'Inbox'); } catch(e) {}
    try {
        var top = app.mailFolders();
        for (var k = 0; k < top.length; k++) {
            try { searchFolder(top[k], top[k].name()); } catch(e) {}
        }
    } catch(e) {}

    result.sort(function(a, b) { return b.date.localeCompare(a.date); });
    return JSON.stringify(result);
})()
"""

_SCRIPT_LIST_FOLDERS = """\
(function() {
    var app = Application('Microsoft Outlook');
    var result = [];

    function list(folder, depth) {
        try {
            var n    = folder.name();
            var msgs = folder.messages();
            result.push({ name: n, depth: depth, count: msgs.length });
            var sub = folder.mailFolders();
            for (var i = 0; i < sub.length; i++) {
                try { list(sub[i], depth + 1); } catch(e) {}
            }
        } catch(e) {}
    }

    try { result.push({ name: 'Inbox', depth: 0, count: app.inbox.messages.length }); } catch(e) {}
    try {
        var top = app.mailFolders();
        for (var i = 0; i < top.length; i++) { try { list(top[i], 0); } catch(e) {} }
    } catch(e) {}

    return JSON.stringify(result);
})()
"""

_SCRIPT_TEST = """\
(function() {
    try {
        var app = Application('Microsoft Outlook');
        var n = app.inbox.messages.length;
        return JSON.stringify({ ok: true, count: n });
    } catch(e) {
        return JSON.stringify({ ok: false, error: e.message });
    }
})()
"""


# ── Conversion ────────────────────────────────────────────────────────────

def _to_message_dict(raw: dict) -> dict:
    """Normalise JXA output to the format used by the rest of the app."""
    return {
        "id":                raw.get("id", ""),
        "conversationId":    raw.get("id", ""),
        "subject":           raw.get("subject", ""),
        "from": {
            "emailAddress": {
                "name":    raw.get("fromName", ""),
                "address": raw.get("fromAddr", ""),
            }
        },
        "receivedDateTime":  raw.get("date", ""),
        "bodyPreview":       raw.get("preview", ""),
        "isRead":            True,
        "folder":            raw.get("folder", ""),
    }


# ── Client ────────────────────────────────────────────────────────────────

class OutlookMacClient:
    def __init__(self):
        self._server = "Outlook desktop (macOS)"

    def test_connection(self) -> tuple[bool, str]:
        try:
            data = _run_jxa(_SCRIPT_TEST, timeout=15)
            if isinstance(data, dict) and data.get("ok"):
                return True, f"Outlook is running — {data.get('count', '?')} messages in inbox"
            return False, data.get("error", "Outlook not accessible") if isinstance(data, dict) else "Unexpected response"
        except subprocess.TimeoutExpired:
            return False, "Timed out — make sure Outlook is open"
        except Exception as e:
            return False, str(e)

    def get_all_messages(self, limit: int = _FETCH_LIMIT) -> list:
        script = (_SCRIPT_GET_HEADERS
                  .replace("LIMIT", str(limit))
                  .replace("MONTHS", str(_MONTHS_BACK)))
        raw = _run_jxa(script, timeout=120)
        seen, result = set(), []
        for r in raw:
            key = r.get("id") or r.get("subject", "")
            if key not in seen:
                seen.add(key)
                result.append(_to_message_dict(r))
        return result

    def get_inbox(self, limit: int = _FETCH_LIMIT) -> list:
        return [m for m in self.get_all_messages(limit) if m.get("folder") == "Inbox"]

    def get_sent(self, limit: int = _FETCH_LIMIT) -> list:
        return [m for m in self.get_all_messages(limit) if m.get("folder") == "Sent"]

    def search_messages(self, query: str, limit: int = 80) -> list:
        script = _SCRIPT_SEARCH.replace("QUERY_JSON", json.dumps(query))
        raw = _run_jxa(script, timeout=120)
        return [_to_message_dict(r) for r in raw[:limit]]

    def get_folders(self) -> list:
        return _run_jxa(_SCRIPT_LIST_FOLDERS, timeout=30)


# Keep IMAP client importable as fallback
def parse_email_address(address_obj):
    if not address_obj:
        return "", ""
    ea = address_obj.get("emailAddress", {})
    return ea.get("name", ""), ea.get("address", "")


def group_messages_by_domain(messages: list) -> list:
    own_domain = OUTLOOK_EMAIL.split("@")[-1].lower() if "@" in OUTLOOK_EMAIL else ""
    domain_map = {}
    for msg in messages:
        _, from_addr = parse_email_address(msg.get("from"))
        domain = from_addr.split("@")[-1].lower() if "@" in from_addr else ""
        if not domain or domain == own_domain:
            continue
        if domain not in domain_map:
            domain_map[domain] = {"domain": domain, "contacts": set(), "messages": []}
        name, _ = parse_email_address(msg.get("from"))
        domain_map[domain]["contacts"].add(f"{name} <{from_addr}>".strip())
        domain_map[domain]["messages"].append(msg)

    result = []
    for domain, info in domain_map.items():
        msgs = sorted(info["messages"], key=lambda m: m.get("receivedDateTime", ""), reverse=True)
        result.append({
            "domain":         domain,
            "contacts":       list(info["contacts"]),
            "message_count":  len(msgs),
            "last_email_date": msgs[0].get("receivedDateTime", "") if msgs else "",
            "latest_subject": msgs[0].get("subject", "") if msgs else "",
        })
    return sorted(result, key=lambda x: x["last_email_date"], reverse=True)
