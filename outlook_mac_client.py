"""Read Outlook emails on macOS via JXA (JavaScript for Automation).
No IMAP, no passwords — talks directly to the running Outlook desktop app.
"""
import subprocess
import json

from config.settings import INTERNAL_EMAIL_DOMAINS, OUTLOOK_EMAIL

_FETCH_LIMIT = 10000   # effectively "all" — cap is the folder size itself
_MONTHS_BACK = 60      # 5 years of history


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


# ── Shared JXA helpers ────────────────────────────────────────────────────
#
# Folder name detection is language-aware.
# "Boîte de réception" → isInbox,  "Éléments envoyés" → isSent.

_FOLDER_HELPERS = """\
    function isInbox(name) {
        var n = name.toLowerCase();
        return n === 'inbox'
            || n.indexOf('réception') >= 0 || n.indexOf('reception') >= 0
            || n === 'bandeja de entrada' || n === 'posteingang'
            || n === 'posta in arrivo'    || n === 'caixa de entrada'
            || n === 'postvak in'         || n === 'ontvangen';
    }
    function isSent(name) {
        var n = name.toLowerCase();
        return n === 'sent items' || n === 'sent mail' || n === 'sent'
            || n.indexOf('envoy') >= 0
            || n === 'elementos enviados' || n === 'enviados'
            || n === 'gesendete elemente' || n === 'gesendet'
            || n === 'posta inviata'      || n === 'itens enviados'
            || n === 'verzonden items'    || n === 'verzonden';
    }
"""

# ── Diagnostic test ───────────────────────────────────────────────────────
#
# Three distinct message-access methods are tested for every inbox/sent folder:
#   A) .messages.length   – JXA specifier (often 0 even when messages exist)
#   B) .messages().length – materialised array (the correct count)
#   C) .messages[0].subject() – indexed access (works on some Outlook builds
#      where the array materialisation is broken)
#
# This tells us exactly which path works so we can fix the fetch accordingly.

_SCRIPT_TEST = """\
(function() {
    try {
        var app = Application('Microsoft Outlook');
        var report = { ok: true, folders: [] };

FOLDER_HELPERS

        function probe(folder) {
            var r = {};
            try { r.specLen  = folder.messages.length;    } catch(e) { r.specErr   = e.message; }
            try { r.arrayLen = folder.messages().length;  } catch(e) { r.arrayErr  = e.message; }
            try { folder.messages[0].subject(); r.idx0ok = true; } catch(e) { r.idx0err = e.message; }
            return r;
        }

        try { var ai = probe(app.inbox); ai.name = 'app.inbox'; report.folders.push(ai); } catch(e) {}

        try {
            var top = app.mailFolders();
            report.topCount = top.length;
            for (var i = 0; i < top.length; i++) {
                try {
                    var n = top[i].name();
                    if (!isInbox(n) && !isSent(n)) continue;
                    var p = probe(top[i]);
                    p.name = n;
                    p.kind = isInbox(n) ? 'INBOX' : 'SENT';
                    report.folders.push(p);
                } catch(e) {}
            }
        } catch(e) { report.topErr = e.message; }

        return JSON.stringify(report);
    } catch(e) {
        return JSON.stringify({ ok: false, error: e.message });
    }
})()
""".replace("FOLDER_HELPERS", _FOLDER_HELPERS)

# ── Header fetch ──────────────────────────────────────────────────────────
#
# fetchFolder tries three access methods in order:
#   1) messages() – normal array materialisation
#   2) index walk  – for Outlook builds where .messages() materialises to []
#                    but .messages[n] still works (iterate until out-of-bounds)

_SCRIPT_GET_HEADERS = """\
(function() {
    var app      = Application('Microsoft Outlook');
    var result   = [];
    var seen     = {};
    var limit    = LIMIT;
    var cutoffMs = Date.now() - MONTHS * 30 * 24 * 60 * 60 * 1000;

FOLDER_HELPERS

    // For recipient objects, emailAddress() materialises to a plain JS object
    // {address: "...", name: "..."} — accessing .address() as a method call fails
    // with "Il est impossible de convertir les types."
    function getAddrFromRecord(rec) {
        var addr = '';
        try { var ea = rec.emailAddress(); if (ea && ea.address) { addr = String(ea.address); } } catch(e) {}
        if (!addr) { try { addr = rec.emailAddress.address() || ''; } catch(e) {} }
        if (!addr) { try { addr = rec.emailAddress.address || ''; } catch(e) {} }
        return addr;
    }
    function getNameFromRecord(rec) {
        var name = '';
        try { var ea = rec.emailAddress(); if (ea && ea.name) { name = String(ea.name); } } catch(e) {}
        if (!name) { try { name = rec.emailAddress.name() || ''; } catch(e) {} }
        if (!name) { try { name = rec.emailAddress.name || ''; } catch(e) {} }
        return name;
    }
    function getSenderAddr(m) {
        var addr = '';
        try { addr = m.sender.emailAddress.address() || ''; } catch(e) {}
        if (!addr) { try { addr = m.sender.emailAddress.address || ''; } catch(e) {} }
        if (!addr) { try { var ea = m.sender.emailAddress(); addr = (typeof ea === 'string') ? ea : ''; } catch(e) {} }
        return addr;
    }
    function getSenderName(m) {
        var name = '';
        try { name = m.sender.emailAddress.name() || ''; } catch(e) {}
        if (!name) { try { name = m.sender.emailAddress.name || ''; } catch(e) {} }
        if (!name) { try { name = m.sender.name()        || ''; } catch(e) {} }
        if (!name) { try { name = m.sender.displayName() || ''; } catch(e) {} }
        return name;
    }
    function collectRecipients(m, fnName) {
        var out = [];
        try {
            var recips = [];
            if (fnName === 'toRecipients') recips = m.toRecipients();
            else if (fnName === 'ccRecipients') recips = m.ccRecipients();
            else if (fnName === 'bccRecipients') recips = m.bccRecipients();
            for (var i = 0; i < recips.length; i++) {
                var addr = getAddrFromRecord(recips[i]);
                var name = getNameFromRecord(recips[i]);
                if (addr || name) out.push({ address: addr, name: name });
            }
        } catch(e) {}
        return out;
    }

    function addMsg(m, folderName) {
        try {
            var id = String(m.id());
            if (seen[id]) return;
            var d;
            try { d = m.timeReceived(); } catch(e) {}
            if (!d) { try { d = m.timeSent(); } catch(e) {} }
            if (!d) return;
            if (d.getTime() < cutoffMs) return;

            var fromAddr = getSenderAddr(m);
            var fromName = getSenderName(m);
            var toRecipients = collectRecipients(m, 'toRecipients');
            var ccRecipients = collectRecipients(m, 'ccRecipients');
            var bccRecipients = collectRecipients(m, 'bccRecipients');

            seen[id] = true;
            result.push({
                id:       id,
                subject:  m.subject() || '',
                fromAddr: fromAddr,
                fromName: fromName,
                toRecipients: toRecipients,
                ccRecipients: ccRecipients,
                bccRecipients: bccRecipients,
                date:     d.toISOString(),
                preview:  '',
                folder:   folderName
            });
        } catch(e) {}
    }

    function fetchFolder(folder, label) {
        // Method 1: materialised array
        try {
            var msgs = folder.messages();
            if (msgs && msgs.length > 0) {
                var start = Math.max(0, msgs.length - limit);
                for (var i = msgs.length - 1; i >= start; i--) { addMsg(msgs[i], label); }
                return;
            }
        } catch(e) {}

        // Method 2: index walk (handles builds where .messages() returns [])
        var consec = 0;
        for (var idx = 0; idx < limit; idx++) {
            try {
                addMsg(folder.messages[idx], label);
                consec = 0;
            } catch(e) {
                if (++consec >= 5) break;   // 5 consecutive misses = end of folder
            }
        }
    }

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

    try { fetchFolder(app.inbox, 'Inbox'); } catch(e) {}
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

# ── Search ────────────────────────────────────────────────────────────────

_SCRIPT_SEARCH = """\
(function() {
    var app      = Application('Microsoft Outlook');
    var query    = QUERY_JSON;
    var result   = [];
    var seen     = {};
    var cutoffMs = Date.now() - 365 * 24 * 60 * 60 * 1000;

FOLDER_HELPERS

    function getAddrM(rec) {
        var a = '';
        try { var ea = rec.emailAddress(); if (ea && ea.address) { a = String(ea.address); } } catch(e) {}
        if (!a) { try { a = rec.emailAddress.address() || ''; } catch(e) {} }
        if (!a) { try { a = rec.emailAddress.address || ''; } catch(e) {} }
        return a;
    }
    function getNameM(rec) {
        var n = '';
        try { var ea = rec.emailAddress(); if (ea && ea.name) { n = String(ea.name); } } catch(e) {}
        if (!n) { try { n = rec.emailAddress.name() || ''; } catch(e) {} }
        if (!n) { try { n = rec.emailAddress.name || ''; } catch(e) {} }
        return n;
    }
    function getSenderAddr(m) {
        var addr = '';
        try { addr = m.sender.emailAddress.address() || ''; } catch(e) {}
        if (!addr) { try { addr = m.sender.emailAddress.address || ''; } catch(e) {} }
        if (!addr) { try { var ea = m.sender.emailAddress(); addr = (typeof ea === 'string') ? ea : ''; } catch(e) {} }
        return addr;
    }
    function getSenderName(m) {
        var name = '';
        try { name = m.sender.emailAddress.name() || ''; } catch(e) {}
        if (!name) { try { name = m.sender.emailAddress.name || ''; } catch(e) {} }
        if (!name) { try { name = m.sender.name() || ''; } catch(e) {} }
        if (!name) { try { name = m.sender.displayName() || ''; } catch(e) {} }
        return name;
    }
    function collectRecipients(m, fnName) {
        var out = [];
        try {
            var recips = [];
            if (fnName === 'toRecipients') recips = m.toRecipients();
            else if (fnName === 'ccRecipients') recips = m.ccRecipients();
            else if (fnName === 'bccRecipients') recips = m.bccRecipients();
            for (var i = 0; i < recips.length; i++) {
                var addr = getAddrM(recips[i]);
                var name = getNameM(recips[i]);
                if (addr || name) out.push({ address: addr, name: name });
            }
        } catch(e) {}
        return out;
    }
    function recipientText(recips) {
        var parts = [];
        for (var i = 0; i < recips.length; i++) {
            parts.push((recips[i].name || '') + ' ' + (recips[i].address || ''));
        }
        return parts.join(' ').toLowerCase();
    }

    function matchMsg(m, folderName) {
        try {
            var id = String(m.id());
            if (seen[id]) return;
            var fromAddr = getSenderAddr(m);
            var fromName = getSenderName(m);
            var toRecipients = collectRecipients(m, 'toRecipients');
            var ccRecipients = collectRecipients(m, 'ccRecipients');
            var bccRecipients = collectRecipients(m, 'bccRecipients');
            var subj = (m.subject() || '').toLowerCase();
            var q    = query.toLowerCase();
            if (subj.indexOf(q) < 0 && fromAddr.toLowerCase().indexOf(q) < 0
                && fromName.toLowerCase().indexOf(q) < 0
                && recipientText(toRecipients).indexOf(q) < 0
                && recipientText(ccRecipients).indexOf(q) < 0
                && recipientText(bccRecipients).indexOf(q) < 0) return;
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
                toRecipients: toRecipients,
                ccRecipients: ccRecipients,
                bccRecipients: bccRecipients,
                date:     d.toISOString(),
                preview:  preview,
                folder:   folderName
            });
        } catch(e) {}
    }

    function searchFolder(folder, name) {
        // Array access
        try {
            var msgs = folder.messages();
            for (var i = 0; i < msgs.length; i++) { matchMsg(msgs[i], name); }
        } catch(e) {}
        // Index fallback
        if (!result.length) {
            var consec = 0;
            for (var idx = 0; idx < 500; idx++) {
                try { matchMsg(folder.messages[idx], name); consec = 0; }
                catch(e) { if (++consec >= 5) break; }
            }
        }
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

# ── Full message fetch ────────────────────────────────────────────────────

_SCRIPT_GET_MESSAGE = """\
(function() {
    var app       = Application('Microsoft Outlook');
    var messageId = MESSAGE_ID_JSON;
    var folderHint = FOLDER_HINT_JSON;
    var result    = null;

FOLDER_HELPERS

    function getAddrFromRecord(rec) {
        var addr = '';
        try { var ea = rec.emailAddress(); if (ea && ea.address) { addr = String(ea.address); } } catch(e) {}
        if (!addr) { try { addr = rec.emailAddress.address() || ''; } catch(e) {} }
        if (!addr) { try { addr = rec.emailAddress.address || ''; } catch(e) {} }
        return addr;
    }
    function getNameFromRecord(rec) {
        var name = '';
        try { var ea = rec.emailAddress(); if (ea && ea.name) { name = String(ea.name); } } catch(e) {}
        if (!name) { try { name = rec.emailAddress.name() || ''; } catch(e) {} }
        if (!name) { try { name = rec.emailAddress.name || ''; } catch(e) {} }
        return name;
    }
    function getSenderAddr(m) {
        var addr = '';
        try { addr = m.sender.emailAddress.address() || ''; } catch(e) {}
        if (!addr) { try { addr = m.sender.emailAddress.address || ''; } catch(e) {} }
        if (!addr) { try { var ea = m.sender.emailAddress(); addr = (typeof ea === 'string') ? ea : ''; } catch(e) {} }
        return addr;
    }
    function getSenderName(m) {
        var name = '';
        try { name = m.sender.emailAddress.name() || ''; } catch(e) {}
        if (!name) { try { name = m.sender.emailAddress.name || ''; } catch(e) {} }
        if (!name) { try { name = m.sender.name()        || ''; } catch(e) {} }
        if (!name) { try { name = m.sender.displayName() || ''; } catch(e) {} }
        return name;
    }
    function collectRecipients(m, fnName) {
        var out = [];
        try {
            var recips = [];
            if (fnName === 'toRecipients') recips = m.toRecipients();
            else if (fnName === 'ccRecipients') recips = m.ccRecipients();
            else if (fnName === 'bccRecipients') recips = m.bccRecipients();
            for (var i = 0; i < recips.length; i++) {
                var addr = getAddrFromRecord(recips[i]);
                var name = getNameFromRecord(recips[i]);
                if (addr || name) out.push({ address: addr, name: name });
            }
        } catch(e) {}
        return out;
    }
    function bodyText(m) {
        var body = '';
        try { body = m.plainTextContent() || ''; } catch(e) {}
        if (!body) { try { body = m.content() || ''; } catch(e) {} }
        if (!body) { try { body = m.body() || ''; } catch(e) {} }
        return String(body || '');
    }
    function capture(m, folderName) {
        try {
            var id = String(m.id());
            if (id !== messageId) return false;
            var d;
            try { d = m.timeReceived(); } catch(e) {}
            if (!d) { try { d = m.timeSent(); } catch(e) {} }
            var body = bodyText(m);
            result = {
                id: id,
                subject: m.subject() || '',
                fromAddr: getSenderAddr(m),
                fromName: getSenderName(m),
                toRecipients: collectRecipients(m, 'toRecipients'),
                ccRecipients: collectRecipients(m, 'ccRecipients'),
                bccRecipients: collectRecipients(m, 'bccRecipients'),
                date: d ? d.toISOString() : '',
                preview: body.replace(/\\s+/g, ' ').substring(0, 280),
                body: body,
                folder: folderName
            };
            return true;
        } catch(e) {}
        return false;
    }
    function searchFolder(folder, label) {
        if (result) return;
        try {
            var msgs = folder.messages();
            for (var i = msgs.length - 1; i >= 0; i--) {
                if (capture(msgs[i], label)) return;
            }
        } catch(e) {}

        var consec = 0;
        for (var idx = 0; idx < 10000 && !result; idx++) {
            try {
                if (capture(folder.messages[idx], label)) return;
                consec = 0;
            } catch(e) {
                if (++consec >= 5) break;
            }
        }
    }
    function searchAllFolders() {
        try { searchFolder(app.inbox, 'Inbox'); } catch(e) {}
        try {
            var top = app.mailFolders();
            for (var k = 0; k < top.length && !result; k++) {
                try { searchFolder(top[k], top[k].name()); } catch(e) {}
            }
        } catch(e) {}
    }
    function findFolder(root, name) {
        try {
            var subs = root.mailFolders();
            for (var i = 0; i < subs.length; i++) {
                try {
                    var n = String(subs[i].name() || '').toLowerCase();
                    var t = name.toLowerCase();
                    if (n === t || n.indexOf(t) >= 0 || t.indexOf(n) >= 0) {
                        return subs[i];
                    }
                } catch(e) {}
            }
        } catch(e) {}
        return null;
    }

    if (folderHint) {
        var f = findFolder(app, folderHint);
        if (f) { searchFolder(f, folderHint); }
        if (result) { return JSON.stringify(result); }
    }

    searchAllFolders();
    return JSON.stringify(result || {});
})()
""".replace("FOLDER_HELPERS", _FOLDER_HELPERS)

# ── Sender probe (debug) ─────────────────────────────────────────────────
# Returns the first 5 inbox messages with every sender-access method tried,
# so we can see which path actually returns the email address.

_SCRIPT_DEBUG_MESSAGES = """\
(function() {
    var app = Application('Microsoft Outlook');
    var result = [];
    var msgs = app.inbox.messages();
    var n = Math.min(5, msgs.length);
    for (var i = msgs.length - 1; i >= Math.max(0, msgs.length - n); i--) {
        var m = msgs[i];
        var row = { subject: '' };
        try { row.subject = m.subject() || ''; } catch(e) {}

        // Path A: sender.emailAddress.address (record sub-field)
        try { row.pathA_addr = m.sender.emailAddress.address() || ''; } catch(e) { row.pathA_err = e.message; }
        try { row.pathA_name = m.sender.emailAddress.name() || ''; } catch(e) {}

        // Path B: sender.emailAddress() directly (might be string or object)
        try {
            var ea = m.sender.emailAddress();
            row.pathB_type = typeof ea;
            row.pathB_val  = String(ea);
        } catch(e) { row.pathB_err = e.message; }

        // Path C: sender.name()
        try { row.pathC_name = m.sender.name() || ''; } catch(e) { row.pathC_err = e.message; }

        // Path D: sender.displayName()
        try { row.pathD_name = m.sender.displayName() || ''; } catch(e) { row.pathD_err = e.message; }

        result.push(row);
    }
    return JSON.stringify(result);
})()
"""

# ── Sent-item recipient probe (debug) ────────────────────────────────────
# Tries every known JXA path to get recipient addresses from sent messages.

_SCRIPT_DEBUG_SENT = """\
(function() {
    var app = Application('Microsoft Outlook');
    var result = { sentFolder: null, msgs: [] };

FOLDER_HELPERS

    // Recursively find the first sent folder that actually has messages
    function findRealSent(folder, depth) {
        if (depth > 5) return null;
        var count = 0;
        try { count = folder.messages().length; } catch(e) {}
        if (count > 0) return folder;
        try {
            var sub = folder.mailFolders();
            for (var i = 0; i < sub.length; i++) {
                try {
                    if (isSent(sub[i].name())) {
                        var found = findRealSent(sub[i], depth + 1);
                        if (found) return found;
                    }
                } catch(e) {}
            }
        } catch(e) {}
        return null;
    }

    var sentFolderObj = null;
    try {
        var top = app.mailFolders();
        for (var k = 0; k < top.length && !sentFolderObj; k++) {
            try {
                var n = top[k].name();
                if (isSent(n)) {
                    sentFolderObj = findRealSent(top[k], 0);
                }
            } catch(e) {}
        }
    } catch(e) { result.err = e.message; }

    if (!sentFolderObj) {
        result.sentFolder = 'NOT FOUND (with messages)';
        return JSON.stringify(result);
    }

    var sentCount = -1;
    try { sentCount = sentFolderObj.messages().length; } catch(e) {}
    var sentName = '';
    try { sentName = sentFolderObj.name(); } catch(e) {}
    result.sentFolder = sentName + ' (' + sentCount + ' msgs)';

    var msgs;
    try { msgs = sentFolderObj.messages(); } catch(e) { result.msgsErr = e.message; return JSON.stringify(result); }

    // Probe the last 3 messages (most recent)
    var n = Math.min(3, msgs.length);
    for (var i = msgs.length - 1; i >= Math.max(0, msgs.length - n); i--) {
        var m = msgs[i];
        var row = {};
        try { row.subject = m.subject() || ''; } catch(e) { row.subject = '(err)'; }

        // Sender (should be the user)
        try { row.sender_addr = m.sender.emailAddress.address() || '(empty)'; } catch(e) { row.sender_err = e.message; }

        // TO recipients - method 1: toRecipients()
        try {
            var r1 = m.toRecipients();
            row.toRecipients_count = r1.length;
            if (r1.length > 0) {
                try { row.toR0_addr_call = r1[0].emailAddress.address() || '(empty)'; } catch(e) { row.toR0_addr_call_err = e.message; }
                try { row.toR0_addr_prop = r1[0].emailAddress.address || '(empty)'; } catch(e) {}
                try { row.toR0_name_call = r1[0].emailAddress.name() || '(empty)'; } catch(e) {}
                try { row.toR0_raw = String(r1[0].emailAddress()); } catch(e) { row.toR0_raw_err = e.message; }
            }
        } catch(e) { row.toRecipients_err = e.message; }

        // recipients() - alternative name
        try { var r2 = m.recipients(); row.recipients_count = r2.length; } catch(e) { row.recipients_err = e.message; }

        // ccRecipients()
        try { var r3 = m.ccRecipients(); row.ccRecipients_count = r3.length; } catch(e) { row.ccRecipients_err = e.message; }

        result.msgs.push(row);
    }
    return JSON.stringify(result);
})()
""".replace("FOLDER_HELPERS", _FOLDER_HELPERS)

# ── Folder list ───────────────────────────────────────────────────────────

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
        try { count = folder.messages().length; } catch(e) {
            try { count = folder.messages.length; } catch(e2) {}
        }
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


# ── Conversion ────────────────────────────────────────────────────────────

def _email_obj(name: str = "", address: str = "") -> dict:
    return {"emailAddress": {"name": name or "", "address": address or ""}}


def _email_domain(address: str) -> str:
    address = (address or "").lower().strip()
    return address.split("@")[-1] if "@" in address else ""


def _internal_domains() -> set[str]:
    domains = {d.lower() for d in INTERNAL_EMAIL_DOMAINS}
    if "@" in OUTLOOK_EMAIL:
        domains.add(OUTLOOK_EMAIL.split("@")[-1].lower())
    return domains


def _is_internal_address(address: str) -> bool:
    domain = _email_domain(address)
    return bool(domain and domain in _internal_domains())


def _recipient_list(raw: dict, key: str) -> list[dict]:
    recipients = []
    for rec in raw.get(key, []) or []:
        if isinstance(rec, dict):
            recipients.append(_email_obj(rec.get("name", ""), rec.get("address", "")))
    return recipients


def _dedupe_people(people: list[dict]) -> list[dict]:
    seen = set()
    out = []
    for person in people:
        name, addr = parse_email_address(person)
        key = (addr or name).lower()
        if not key or key in seen:
            continue
        seen.add(key)
        out.append(_email_obj(name, addr))
    return out


def external_participants_for_message(msg: dict) -> list[dict]:
    people = []
    for key in ("sender", "from"):
        if msg.get(key):
            people.append(msg[key])
    for key in ("toRecipients", "ccRecipients", "bccRecipients"):
        people.extend(msg.get(key, []) or [])
    return _dedupe_people([
        person for person in people
        if not _is_internal_address(parse_email_address(person)[1])
    ])


def external_domains_for_message(msg: dict) -> list[str]:
    domains = []
    seen = set()
    for person in external_participants_for_message(msg):
        _, addr = parse_email_address(person)
        domain = _email_domain(addr)
        if domain and domain not in seen:
            seen.add(domain)
            domains.append(domain)
    return domains


def _to_message_dict(raw: dict) -> dict:
    sender = _email_obj(raw.get("fromName", ""), raw.get("fromAddr", ""))
    to_recipients = _recipient_list(raw, "toRecipients")
    cc_recipients = _recipient_list(raw, "ccRecipients")
    bcc_recipients = _recipient_list(raw, "bccRecipients")
    msg = {
        "id":               raw.get("id", ""),
        "conversationId":   raw.get("id", ""),
        "subject":          raw.get("subject", ""),
        "sender":           sender,
        "from":             sender,
        "toRecipients":     to_recipients,
        "ccRecipients":     cc_recipients,
        "bccRecipients":    bcc_recipients,
        "receivedDateTime": raw.get("date", ""),
        "bodyPreview":      raw.get("preview", ""),
        "body":             {"contentType": "text", "content": raw.get("body", "")},
        "isRead":           True,
        "folder":           raw.get("folder", ""),
    }
    external = external_participants_for_message(msg)
    msg["external_participants"] = external
    msg["direction"] = _direction_for_message(msg)
    if external:
        # Keep the Graph-like "from" field useful for existing UI/code by
        # pointing it at the external counterparty, even when an internal
        # colleague sent the message to the client.
        msg["from"] = external[0]
    return msg


# ── Client ────────────────────────────────────────────────────────────────

class OutlookMacClient:
    def __init__(self):
        self._server = "Outlook desktop (macOS)"

    def test_connection(self) -> tuple[bool, str]:
        try:
            data = _run_jxa(_SCRIPT_TEST, timeout=25)
            if not isinstance(data, dict) or not data.get("ok"):
                err = data.get("error", "Outlook not accessible") if isinstance(data, dict) else "No response"
                return False, err

            folders = data.get("folders", [])
            if not folders:
                return False, (
                    f"Outlook is running but no inbox/sent folders matched. "
                    f"Top-level folder count: {data.get('topCount', '?')}"
                )

            parts = []
            any_accessible = False
            for f in folders:
                name = f.get("name", "?")
                array_len = f.get("arrayLen")
                spec_len  = f.get("specLen")
                idx_ok    = f.get("idx0ok", False)

                if array_len is not None and array_len > 0:
                    parts.append(f"{name}: {array_len} msgs (array ✓)")
                    any_accessible = True
                elif idx_ok:
                    parts.append(f"{name}: array=0 but index access ✓")
                    any_accessible = True
                elif spec_len is not None and spec_len > 0:
                    parts.append(f"{name}: {spec_len} msgs (specifier only)")
                    any_accessible = True
                else:
                    arr_err = f.get("arrayErr", "")
                    idx_err = f.get("idx0err", "")
                    parts.append(f"{name}: 0 — array err: {arr_err or 'empty'}, idx err: {idx_err or 'empty'}")

            if any_accessible:
                return True, "Outlook accessible — " + " | ".join(parts)
            else:
                return False, (
                    "Folders found but messages are not accessible via JXA. "
                    "This is a known limitation of New Outlook for Mac. "
                    "Fix: open Outlook → View → Switch to Legacy Outlook. "
                    "Details: " + " | ".join(parts)
                )

        except subprocess.TimeoutExpired:
            return False, "Timed out — make sure Outlook is open"
        except Exception as e:
            return False, str(e)

    def raw_diagnostic(self) -> dict:
        """Return the full diagnostic dict for the /api/debug-outlook endpoint."""
        try:
            return _run_jxa(_SCRIPT_TEST, timeout=25)
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def debug_messages(self) -> list:
        """Return first 5 inbox messages with all sender-access paths probed."""
        try:
            return _run_jxa(_SCRIPT_DEBUG_MESSAGES, timeout=30)
        except Exception as e:
            return [{"error": str(e)}]

    def debug_sent(self) -> dict:
        """Probe all recipient-access paths on the first 3 sent messages."""
        try:
            return _run_jxa(_SCRIPT_DEBUG_SENT, timeout=30)
        except Exception as e:
            return {"error": str(e)}

    def get_all_messages(self, limit: int = _FETCH_LIMIT) -> list:
        script = (_SCRIPT_GET_HEADERS
                  .replace("LIMIT", str(limit))
                  .replace("MONTHS", str(_MONTHS_BACK)))
        raw = _run_jxa(script, timeout=600)  # up to 10 min for large mailboxes
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

    def get_message(self, message_id: str, folder_hint: str = "") -> dict:
        script = (_SCRIPT_GET_MESSAGE
                  .replace("MESSAGE_ID_JSON", json.dumps(message_id))
                  .replace("FOLDER_HINT_JSON", json.dumps(folder_hint)))
        raw = _run_jxa(script, timeout=180)
        if not raw:
            return {}
        return _to_message_dict(raw)

    def get_folders(self) -> list:
        return _run_jxa(_SCRIPT_LIST_FOLDERS, timeout=30)


# ── Helpers ───────────────────────────────────────────────────────────────

def parse_email_address(address_obj):
    if not address_obj:
        return "", ""
    ea = address_obj.get("emailAddress", {})
    return ea.get("name", ""), ea.get("address", "")


def _direction_for_message(msg: dict) -> str:
    if not msg.get("sender") and (msg.get("folder") or "").lower() == "sent":
        return "OUT"
    _, sender_addr = parse_email_address(msg.get("sender") or msg.get("from"))
    if sender_addr and not _is_internal_address(sender_addr):
        return "IN"
    external_recipients = []
    for key in ("toRecipients", "ccRecipients", "bccRecipients"):
        external_recipients.extend([
            person for person in msg.get(key, []) or []
            if not _is_internal_address(parse_email_address(person)[1])
        ])
    if external_recipients:
        return "OUT"
    return "OUT" if (msg.get("folder") or "").lower() == "sent" else "IN"


def _meeting_signal(subject: str) -> str:
    """Return a compact meeting status inferred from an email subject."""
    s = (subject or "").lower()
    if not s:
        return ""
    if any(k in s for k in ("cancelled", "canceled", "annulé", "annule")):
        return "cancelled"
    if any(k in s for k in ("accepted", "accepté", "accepte", "confirmed", "confirmation")):
        return "accepted"
    if any(k in s for k in ("tentative", "provisoire")):
        return "tentative"
    if any(k in s for k in ("zoom", "teams", "meeting", "call", "invite", "invitation", "calendar")):
        return "scheduled"
    return ""


def group_messages_by_domain(messages: list) -> list:
    domain_map = {}
    for msg in messages:
        external_people = external_participants_for_message(msg)
        if not external_people:
            continue
        seen_domains = set()
        for person in external_people:
            name, addr = parse_email_address(person)
            domain = _email_domain(addr)
            if not domain or domain in seen_domains:
                continue
            seen_domains.add(domain)
            if domain not in domain_map:
                domain_map[domain] = {"domain": domain, "contacts": set(), "messages": []}
            domain_map[domain]["contacts"].add(f"{name} <{addr}>".strip())
            domain_map[domain]["messages"].append(msg)

    result = []
    for domain, info in domain_map.items():
        msgs = sorted(info["messages"], key=lambda m: m.get("receivedDateTime", ""), reverse=True)
        latest = msgs[0] if msgs else {}
        latest_people = external_participants_for_message(latest)
        latest_name, latest_addr = parse_email_address(latest_people[0] if latest_people else latest.get("from"))
        latest_contact = (
            f"{latest_name} <{latest_addr}>".strip()
            if latest_name and latest_addr else latest_name or latest_addr
        )
        meeting_msgs = [
            (m, _meeting_signal(m.get("subject", "")))
            for m in msgs
            if _meeting_signal(m.get("subject", ""))
        ]
        latest_meeting = meeting_msgs[0] if meeting_msgs else ({}, "")
        result.append({
            "domain":          domain,
            "contacts":        list(info["contacts"]),
            "message_count":   len(msgs),
            "last_email_date": msgs[0].get("receivedDateTime", "") if msgs else "",
            "latest_subject":  msgs[0].get("subject", "") if msgs else "",
            "latest_folder":   latest.get("folder", ""),
            "latest_direction": _direction_for_message(latest) if latest else "",
            "latest_contact":  latest_contact,
            "inbox_count":     sum(1 for m in msgs if (m.get("folder") or "").lower() != "sent"),
            "sent_count":      sum(1 for m in msgs if (m.get("folder") or "").lower() == "sent"),
            "meeting_count":   len(meeting_msgs),
            "latest_meeting_date": latest_meeting[0].get("receivedDateTime", ""),
            "latest_meeting_subject": latest_meeting[0].get("subject", ""),
            "latest_meeting_folder": latest_meeting[0].get("folder", ""),
            "latest_meeting_signal": latest_meeting[1],
        })
    return sorted(result, key=lambda x: x["last_email_date"], reverse=True)
