"""Read Outlook emails via IMAP — no Azure App required."""
import imaplib
import email
from email.header import decode_header as _decode_header
from email.utils import parseaddr, parsedate_to_datetime
from datetime import datetime, timezone
import re

from config.settings import OUTLOOK_EMAIL, OUTLOOK_PASSWORD, IMAP_SERVER, IMAP_PORT

# Known Outlook IMAP servers keyed by email domain
_IMAP_SERVERS = {
    "outlook.com":  "imap-mail.outlook.com",
    "hotmail.com":  "imap-mail.outlook.com",
    "live.com":     "imap-mail.outlook.com",
    "msn.com":      "imap-mail.outlook.com",
}
_M365_DEFAULT = "outlook.office365.com"

# Candidate folder names for Sent Items (Outlook uses localised names)
_SENT_CANDIDATES = ['"Sent Items"', "Sent", '"Sent Mail"']


def _detect_server(email_addr):
    if IMAP_SERVER:
        return IMAP_SERVER
    domain = email_addr.split("@")[-1].lower()
    return _IMAP_SERVERS.get(domain, _M365_DEFAULT)


def _decode_str(value):
    """Decode a possibly encoded email header string."""
    if value is None:
        return ""
    parts = _decode_header(value)
    decoded = []
    for part, charset in parts:
        if isinstance(part, bytes):
            decoded.append(part.decode(charset or "utf-8", errors="replace"))
        else:
            decoded.append(part)
    return "".join(decoded)


def _parse_message(raw_bytes):
    """Parse a raw IMAP message into a dict."""
    msg = email.message_from_bytes(raw_bytes)
    from_name, from_addr = parseaddr(_decode_str(msg.get("From", "")))
    to_raw = _decode_str(msg.get("To", ""))
    subject = _decode_str(msg.get("Subject", "(no subject)"))

    # Date
    date_str = msg.get("Date", "")
    try:
        dt = parsedate_to_datetime(date_str)
        received_dt = dt.astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    except Exception:
        received_dt = ""

    # Body preview
    preview = ""
    if msg.is_multipart():
        for part in msg.walk():
            if part.get_content_type() == "text/plain" and not part.get("Content-Disposition"):
                try:
                    preview = part.get_payload(decode=True).decode(
                        part.get_content_charset() or "utf-8", errors="replace"
                    )[:300]
                except Exception:
                    pass
                break
    else:
        try:
            preview = msg.get_payload(decode=True).decode(
                msg.get_content_charset() or "utf-8", errors="replace"
            )[:300]
        except Exception:
            pass

    preview = re.sub(r"\s+", " ", preview).strip()

    return {
        "id": msg.get("Message-ID", ""),
        "conversationId": msg.get("Thread-Index", msg.get("In-Reply-To", msg.get("Message-ID", ""))),
        "subject": subject,
        "from": {"emailAddress": {"name": from_name, "address": from_addr}},
        "receivedDateTime": received_dt,
        "bodyPreview": preview,
        "isRead": True,
        "to": to_raw,
    }


class OutlookIMAPClient:
    def __init__(self):
        self._server = _detect_server(OUTLOOK_EMAIL)
        self._conn = None

    # ── Connection ────────────────────────────────────────────────────────

    def connect(self):
        self._conn = imaplib.IMAP4_SSL(self._server, IMAP_PORT)
        self._conn.login(OUTLOOK_EMAIL, OUTLOOK_PASSWORD)
        return self

    def disconnect(self):
        if self._conn:
            try:
                self._conn.logout()
            except Exception:
                pass
            self._conn = None

    def test_connection(self):
        """Return (ok, message)."""
        try:
            self.connect()
            self.disconnect()
            return True, "Connected successfully"
        except imaplib.IMAP4.error as e:
            return False, str(e)
        except Exception as e:
            return False, str(e)

    # ── Fetching ──────────────────────────────────────────────────────────

    def _select(self, folder):
        status, data = self._conn.select(folder)
        if status != "OK":
            raise ValueError(f"Cannot select folder {folder}: {data}")
        return int(data[0])

    def _fetch_folder(self, folder, limit=150):
        try:
            count = self._select(folder)
        except Exception:
            return []

        if count == 0:
            return []

        start = max(1, count - limit + 1)
        msg_range = f"{start}:{count}"
        typ, data = self._conn.fetch(msg_range, "(RFC822)")
        if typ != "OK":
            return []

        messages = []
        for item in data:
            if isinstance(item, tuple):
                try:
                    messages.append(_parse_message(item[1]))
                except Exception:
                    pass
        return list(reversed(messages))  # newest first

    def _sent_folder(self):
        typ, folders = self._conn.list()
        if typ != "OK":
            return None
        folder_names = []
        for f in folders:
            if isinstance(f, bytes):
                parts = f.decode().split('"/"')
                name = parts[-1].strip().strip('"')
                folder_names.append(name)
        for candidate in ["Sent Items", "Sent", "Sent Mail"]:
            if any(candidate.lower() in n.lower() for n in folder_names):
                return f'"{candidate}"' if " " in candidate else candidate
        return None

    def get_inbox(self, limit=150):
        self.connect()
        try:
            return self._fetch_folder("INBOX", limit)
        finally:
            self.disconnect()

    def get_sent(self, limit=150):
        self.connect()
        try:
            sent_folder = self._sent_folder() or '"Sent Items"'
            return self._fetch_folder(sent_folder, limit)
        finally:
            self.disconnect()

    def get_all_messages(self, limit=150):
        """Return inbox + sent, newest first, deduped by Message-ID."""
        self.connect()
        try:
            inbox = self._fetch_folder("INBOX", limit)
            sent_folder = self._sent_folder() or '"Sent Items"'
            sent = self._fetch_folder(sent_folder, limit)
        finally:
            self.disconnect()

        seen = set()
        merged = []
        for m in inbox + sent:
            key = m["id"] or m["subject"]
            if key not in seen:
                seen.add(key)
                merged.append(m)
        return sorted(merged, key=lambda m: m["receivedDateTime"], reverse=True)

    def search_messages(self, query, limit=80):
        """Search subject and body for a keyword."""
        self.connect()
        results = []
        try:
            for folder in ["INBOX", self._sent_folder() or '"Sent Items"']:
                try:
                    self._select(folder)
                    # Search subject or from
                    typ, data = self._conn.search(
                        None,
                        f'(OR SUBJECT "{query}" FROM "{query}")'
                    )
                    if typ != "OK" or not data[0]:
                        continue
                    ids = data[0].split()[-limit:]
                    if not ids:
                        continue
                    id_str = ",".join(i.decode() for i in ids)
                    typ2, msgs = self._conn.fetch(id_str, "(RFC822)")
                    if typ2 != "OK":
                        continue
                    for item in msgs:
                        if isinstance(item, tuple):
                            try:
                                results.append(_parse_message(item[1]))
                            except Exception:
                                pass
                except Exception:
                    continue
        finally:
            self.disconnect()
        return sorted(results, key=lambda m: m["receivedDateTime"], reverse=True)


# ── Grouping helpers ──────────────────────────────────────────────────────

def parse_email_address(address_obj):
    if not address_obj:
        return "", ""
    ea = address_obj.get("emailAddress", {})
    return ea.get("name", ""), ea.get("address", "")


def group_messages_by_domain(messages):
    """Group messages by sender domain, skipping the user's own domain."""
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
            "domain": domain,
            "contacts": list(info["contacts"]),
            "message_count": len(msgs),
            "last_email_date": msgs[0].get("receivedDateTime", "") if msgs else "",
            "latest_subject": msgs[0].get("subject", "") if msgs else "",
        })
    return sorted(result, key=lambda x: x["last_email_date"], reverse=True)
