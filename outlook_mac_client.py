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
#
# Folder detection is language-aware: Outlook on a French Mac names folders
# "Boîte de réception" and "Éléments envoyés".  We match by substring so the
# same code works for English, French, Spanish, German, Italian and Portuguese.
#
# Fetch strategy:
#   1. app.inbox shortcut (may be empty on some versions — try anyway)
#   2. Recursive scan of app.mailFolders() — always works, language-aware
#   (app.accounts() is NOT used: it throws "Message incompréhensible" on
#    French Outlook for Mac and is therefore unreliable.)

_FOLDER_HELPERS = """\
    // Returns true for any localised "Inbox" folder name.
    function isInbox(name) {
        var n = name.toLowerCase();
        return n === 'inbox'
            || n.indexOf('réception') >= 0   // Boîte de réception (FR)
            || n.indexOf('reception') >= 0    // Boite de reception (FR no accent)
            || n === 'bandeja de entrada'     // ES
            || n === 'posteingang'            // DE
            || n === 'posta in arrivo'        // IT
            || n === 'caixa de entrada'       // PT
            || n === 'postvak in'             // NL
            || n === 'ontvangen';
    }

    // Returns true for any localised "Sent Items" folder name.
    function isSent(name) {
        var n = name.toLowerCase();
        return n === 'sent items' || n === 'sent mail' || n === 'sent'
            || n.indexOf('envoy')   >= 0   // Éléments envoyés / Envoyés (FR)
            || n === 'elementos enviados'  // ES
            || n === 'enviados'            // ES short
            || n === 'gesendete elemente'  // DE
            || n === 'gesendet'            // DE short
            || n === 'posta inviata'       // IT
            || n === 'itens enviados'      // PT
            || n === 'verzonden items'     // NL
            || n === 'verzonden';
    }
"""

_SCRIPT_GET_HEADERS = """\
(function() {
    var app = Application('Microsoft Outlook');
    var result  = [];
    var seen    = {};
    var limit   = LIMIT;
    var cutoffMs = Date.now() - MONTHS * 30 * 24 * 60 * 60 * 1000;

FOLDER_HELPERS

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

    function fetchFolder(folder, label) {
        try {
            var msgs  = folder.messages();
            var total = msgs.length;
            if (total === 0) return;
            var start = Math.max(0, total - limit);
            for (var i = total - 1; i >= start; i--) { addMsg(msgs[i], label); }
        } catch(e) {}
    }

    // Recurse into every sub-folder looking for more inbox/sent folders.
    function walkFolders(parent) {
        var sub;
        try { sub = parent.mailFolders(); } catch(e) { return; }
        for (var i = 0; i < sub.length; i++) {
            try {
                var n = sub[i].name();
                if (isInbox(n))     { fetchFolder(sub[i], 'Inbox'); }
                else if (isSent(n)) { fetchFolder(sub[i], 'Sent');  }
                walkFolders(sub[i]);
            } catch(e) {}
        }
    }

    // 1. app.inbox shortcut
    try { fetchFolder(app.inbox, 'Inbox'); } catch(e) {}

    // 2. Full recursive scan of all top-level folders (language-aware)
    try {
        var top = app.mailFolders();
        for (var k = 0; k < top.length; k++) {
            try {
                var n = top[k].name();
                if (isInbox(n))     { fetchFolder(top[k], 'Inbox'); }
                else if (isSent(n)) { fetchFolder(top[k], 'Sent');  }
                walkFolders(top[k]);
            } catch(e) {}
        }
    } catch(e) {}

    result.sort(function(a, b) { return b.date.localeCompare(a.date); });
    return JSON.stringify(result);
})()
""".replace("FOLDER_HELPERS", _FOLDER_HELPERS)

_SCRIPT_SEARCH = """\
(function() {
    var app   = Application('Microsoft Outlook');
    var query = QUERY_JSON;
    var result = [];
    var seen   = {};
    var cutoffMs = Date.now() - 365 * 24 * 60 * 60 * 1000;

FOLDER_HELPERS

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
""".replace("FOLDER_HELPERS", _FOLDER_HELPERS)

_SCRIPT_LIST_FOLDERS = """\
(function() {
    var app    = Application('Microsoft Outlook');
    var result = [];
    var seen   = {};

    function list(folder, depth) {
        var n = '';
        try { n = folder.name(); } catch(e) { return; }
        var key = n + '|' + depth;
        if (seen[key]) return;
        seen[key] = true;
        var count = -1;
        try { count = folder.messages.length; } catch(e) {}
        result.push({ name: n, depth: depth, count: count });
        try {
            var sub = folder.mailFolders();
            for (var i = 0; i < sub.length; i++) {
                try { list(sub[i], depth + 1); } catch(e) {}
            }
        } catch(e) {}
    }

    try { list(app.inbox, 0); } catch(e) {}
    try {
        var top = app.mailFolders();
        for (var i = 0; i < top.length; i++) { try { list(top[i], 0); } catch(e) {} }
    } catch(e) {}

    return JSON.stringify(result);
})()
"""

# Diagnostic: reports which top-level folders exist and whether each
# matches the inbox/sent patterns, so we can verify language detection.
_SCRIPT_TEST = """\
(function() {
    try {
        var app = Application('Microsoft Outlook');
        var report = { ok: true, folderMatches: [] };

FOLDER_HELPERS

        // Report what app.inbox sees
        try {
            report.appInboxCount = app.inbox.messages.length;
        } catch(e) {
            report.appInboxError = e.message;
        }

        // Show which top-level folders match inbox/sent patterns
        var totalInbox = 0, totalSent = 0;
        try {
            var top = app.mailFolders();
            report.topFolderCount = top.length;
            for (var i = 0; i < top.length; i++) {
                try {
                    var n     = top[i].name();
                    var count = -1;
                    try { count = top[i].messages.length; } catch(e) {}
                    var kind = isInbox(n) ? 'INBOX' : (isSent(n) ? 'SENT' : null);
                    if (kind) {
                        report.folderMatches.push({ name: n, kind: kind, count: count });
                        if (kind === 'INBOX') totalInbox += (count > 0 ? count : 0);
                        if (kind === 'SENT')  totalSent  += (count > 0 ? count : 0);
                    }
                } catch(e) {}
            }
        } catch(e) {
            report.topFoldersError = e.message;
        }

        report.totalInboxMessages = totalInbox;
        report.totalSentMessages  = totalSent;
        return JSON.stringify(report);
    } catch(e) {
        return JSON.stringify({ ok: false, error: e.message });
    }
})()
""".replace("FOLDER_HELPERS", _FOLDER_HELPERS)


# ── Conversion ────────────────────────────────────────────────────────────

def _to_message_dict(raw: dict) -> dict:
    return {
        "id":               raw.get("id", ""),
        "conversationId":   raw.get("id", ""),
        "subject":          raw.get("subject", ""),
        "from": {
            "emailAddress": {
                "name":    raw.get("fromName", ""),
                "address": raw.get("fromAddr", ""),
            }
        },
        "receivedDateTime": raw.get("date", ""),
        "bodyPreview":      raw.get("preview", ""),
        "isRead":           True,
        "folder":           raw.get("folder", ""),
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

            matches = data.get("folderMatches", [])
            inbox_total = data.get("totalInboxMessages", 0)
            sent_total  = data.get("totalSentMessages", 0)

            if not matches:
                top_n = data.get("topFolderCount", 0)
                app_n = data.get("appInboxCount", 0)
                hint  = f"app.inbox={app_n}, {top_n} top-level folders found but none matched inbox/sent patterns"
                return False, f"0 messages — {hint}"

            parts = [f"{m['name']} ({m['count']})" for m in matches]
            total = inbox_total + sent_total
            return True, (
                f"{total} messages visible — matched: {', '.join(parts)}"
            )

        except subprocess.TimeoutExpired:
            return False, "Timed out — make sure Outlook is open"
        except Exception as e:
            return False, str(e)

    def get_all_messages(self, limit: int = _FETCH_LIMIT) -> list:
        script = (_SCRIPT_GET_HEADERS
                  .replace("LIMIT", str(limit))
                  .replace("MONTHS", str(_MONTHS_BACK)))
        raw = _run_jxa(script, timeout=180)
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
