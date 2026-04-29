"""Read Outlook emails on macOS via JXA (JavaScript for Automation).
No IMAP, no passwords — talks directly to the running Outlook desktop app.
"""
import subprocess
import json

from config.settings import OUTLOOK_EMAIL

_FETCH_LIMIT = 300
_MONTHS_BACK = 12


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

# Tries three strategies in order so it works for both single-account and
# multi-account Outlook setups (Exchange, IMAP, multiple profiles).
#
# Strategy 1: app.inbox  (classic single-account shortcut)
# Strategy 2: app.accounts() loop  (Exchange / multi-account)
# Strategy 3: app.mailFolders() recursive scan  (IMAP or custom folder layout)
#
# All three are tried; duplicates are removed by message id before returning.
_SCRIPT_GET_HEADERS = """\
(function() {
    var app = Application('Microsoft Outlook');
    var result = [];
    var seen   = {};
    var limit  = LIMIT;
    var cutoffMs = Date.now() - MONTHS * 30 * 24 * 60 * 60 * 1000;

    function addMsg(m, folderName) {
        try {
            var id = String(m.id());
            if (seen[id]) return;
            var d;
            try { d = m.timeReceived(); } catch(e) {}
            if (!d) { try { d = m.timeSent(); } catch(e) {} }
            if (!d) return;
            if (d.getTime() < cutoffMs) return;
            var fromAddr = '', fromName = '';
            try { fromAddr = m.sender.emailAddress() || ''; } catch(e) {}
            try { fromName = m.sender.name()         || ''; } catch(e) {}
            seen[id] = true;
            result.push({
                id:       id,
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
            var msgs  = folder.messages();
            var total = msgs.length;
            if (total === 0) return;
            var start = Math.max(0, total - limit);
            for (var i = total - 1; i >= start; i--) { addMsg(msgs[i], name); }
        } catch(e) {}
    }

    // Walk a folder tree; harvest Inbox and Sent wherever they live.
    function walkFolders(parent) {
        var sub;
        try { sub = parent.mailFolders(); } catch(e) { return; }
        for (var i = 0; i < sub.length; i++) {
            try {
                var n = sub[i].name().toLowerCase();
                if (n === 'inbox') {
                    fetchFolder(sub[i], 'Inbox');
                } else if (n === 'sent items' || n === 'sent mail' || n === 'sent') {
                    fetchFolder(sub[i], 'Sent');
                }
                walkFolders(sub[i]);
            } catch(e) {}
        }
    }

    // ── Strategy 1: app.inbox (works for simple single-account setups) ──
    try { fetchFolder(app.inbox, 'Inbox'); } catch(e) {}

    // ── Strategy 2: iterate accounts (Exchange / multi-account) ──
    try {
        var accounts = app.accounts();
        for (var a = 0; a < accounts.length; a++) {
            try { fetchFolder(accounts[a].inbox, 'Inbox'); } catch(e) {}
            walkFolders(accounts[a]);
        }
    } catch(e) {}

    // ── Strategy 3: app.mailFolders() recursive scan (IMAP / fallback) ──
    try {
        var top = app.mailFolders();
        for (var k = 0; k < top.length; k++) {
            try {
                var n = top[k].name().toLowerCase();
                if (n === 'inbox') {
                    fetchFolder(top[k], 'Inbox');
                } else if (n === 'sent items' || n === 'sent mail' || n === 'sent') {
                    fetchFolder(top[k], 'Sent');
                }
                walkFolders(top[k]);
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
    var app   = Application('Microsoft Outlook');
    var query = QUERY_JSON;
    var result = [];
    var seen   = {};
    var cutoffMs = Date.now() - 365 * 24 * 60 * 60 * 1000;

    function matchMsg(m, folderName) {
        try {
            var id = String(m.id());
            if (seen[id]) return;
            var fromAddr = '', fromName = '';
            try { fromAddr = m.sender.emailAddress() || ''; } catch(e) {}
            try { fromName = m.sender.name()         || ''; } catch(e) {}
            var subj = (m.subject() || '').toLowerCase();
            var q    = query.toLowerCase();
            if (subj.indexOf(q) < 0 && fromAddr.toLowerCase().indexOf(q) < 0
                && fromName.toLowerCase().indexOf(q) < 0) return;
            var d;
            try { d = m.timeReceived(); } catch(e) {}
            if (!d) { try { d = m.timeSent(); } catch(e) {} }
            if (!d) return;
            if (d.getTime() < cutoffMs) return;
            var preview = '';
            try { preview = (m.plainTextContent() || '').replace(/\\s+/g,' ').substring(0, 280); } catch(e) {}
            seen[id] = true;
            result.push({
                id:       id,
                subject:  m.subject() || '',
                fromAddr: fromAddr,
                fromName: fromName,
                date:     d.toISOString(),
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

    // Search all three entry points (same as header fetch)
    try { searchFolder(app.inbox, 'Inbox'); } catch(e) {}
    try {
        var accounts = app.accounts();
        for (var a = 0; a < accounts.length; a++) {
            try { searchFolder(accounts[a].inbox, 'Inbox'); } catch(e) {}
            try {
                var af = accounts[a].mailFolders();
                for (var f = 0; f < af.length; f++) {
                    try { searchFolder(af[f], af[f].name()); } catch(e) {}
                }
            } catch(e) {}
        }
    } catch(e) {}
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
    var app    = Application('Microsoft Outlook');
    var result = [];
    var seen   = {};

    function list(folder, depth) {
        var n = '';
        try { n = folder.name(); } catch(e) { return; }
        if (seen[n + depth]) return;
        seen[n + depth] = true;
        try {
            var count = folder.messages.length;
            result.push({ name: n, depth: depth, count: count });
        } catch(e) {
            result.push({ name: n, depth: depth, count: -1 });
        }
        try {
            var sub = folder.mailFolders();
            for (var i = 0; i < sub.length; i++) {
                try { list(sub[i], depth + 1); } catch(e) {}
            }
        } catch(e) {}
    }

    try { list(app.inbox, 0); } catch(e) {}
    try {
        var accounts = app.accounts();
        for (var a = 0; a < accounts.length; a++) {
            try {
                var name = accounts[a].name() || ('Account ' + a);
                result.push({ name: '── ' + name + ' ──', depth: 0, count: -1 });
                var folders = accounts[a].mailFolders();
                for (var f = 0; f < folders.length; f++) {
                    try { list(folders[f], 1); } catch(e) {}
                }
            } catch(e) {}
        }
    } catch(e) {}
    try {
        var top = app.mailFolders();
        for (var i = 0; i < top.length; i++) { try { list(top[i], 0); } catch(e) {} }
    } catch(e) {}

    return JSON.stringify(result);
})()
"""

# Diagnostic test — returns counts from every accessible entry point
# so we can tell the user exactly what is (and isn't) reachable.
_SCRIPT_TEST = """\
(function() {
    try {
        var app    = Application('Microsoft Outlook');
        var report = { ok: true, strategies: [] };

        // Strategy 1: app.inbox
        try {
            var n = app.inbox.messages.length;
            report.strategies.push({ name: 'app.inbox', count: n });
        } catch(e) {
            report.strategies.push({ name: 'app.inbox', error: e.message });
        }

        // Strategy 2: accounts
        try {
            var accounts = app.accounts();
            report.accountCount = accounts.length;
            for (var a = 0; a < accounts.length; a++) {
                try {
                    var aName = accounts[a].name() || ('account[' + a + ']');
                    try {
                        var n2 = accounts[a].inbox.messages.length;
                        report.strategies.push({ name: aName + '.inbox', count: n2 });
                    } catch(e) {
                        report.strategies.push({ name: aName + '.inbox', error: e.message });
                    }
                } catch(e) {}
            }
        } catch(e) {
            report.strategies.push({ name: 'app.accounts()', error: e.message });
        }

        // Strategy 3: app.mailFolders
        try {
            var top = app.mailFolders();
            report.topFolderCount = top.length;
            var folderNames = [];
            for (var i = 0; i < Math.min(top.length, 10); i++) {
                try { folderNames.push(top[i].name()); } catch(e) {}
            }
            report.topFolders = folderNames;
        } catch(e) {
            report.strategies.push({ name: 'app.mailFolders()', error: e.message });
        }

        return JSON.stringify(report);
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
            data = _run_jxa(_SCRIPT_TEST, timeout=20)
            if not isinstance(data, dict) or not data.get("ok"):
                err = data.get("error", "Outlook not accessible") if isinstance(data, dict) else "No response"
                return False, err

            total = 0
            parts = []
            for s in data.get("strategies", []):
                if "count" in s:
                    parts.append(f"{s['name']}: {s['count']}")
                    total += s["count"]
                elif "error" in s:
                    parts.append(f"{s['name']}: ✗ {s['error']}")

            top_folders = data.get("topFolders", [])
            summary = f"{total} messages visible"
            if parts:
                summary += " (" + ", ".join(parts) + ")"
            if top_folders:
                summary += f" | top folders: {', '.join(top_folders)}"
            return True, summary

        except subprocess.TimeoutExpired:
            return False, "Timed out — make sure Outlook is open"
        except Exception as e:
            return False, str(e)

    def get_all_messages(self, limit: int = _FETCH_LIMIT) -> list:
        script = (_SCRIPT_GET_HEADERS
                  .replace("LIMIT", str(limit))
                  .replace("MONTHS", str(_MONTHS_BACK)))
        raw = _run_jxa(script, timeout=180)
        # JXA already deduplicates by id; do a Python-side safety pass too
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
        raw = _run_jxa(script, timeout=180)
        return [_to_message_dict(r) for r in raw[:limit]]

    def get_folders(self) -> list:
        return _run_jxa(_SCRIPT_LIST_FOLDERS, timeout=30)


# ── Helpers ───────────────────────────────────────────────────────────────

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
            "domain":          domain,
            "contacts":        list(info["contacts"]),
            "message_count":   len(msgs),
            "last_email_date": msgs[0].get("receivedDateTime", "") if msgs else "",
            "latest_subject":  msgs[0].get("subject", "") if msgs else "",
        })
    return sorted(result, key=lambda x: x["last_email_date"], reverse=True)
