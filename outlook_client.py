"""Microsoft Graph API client for reading Outlook emails."""
import msal
import requests
from datetime import datetime, timezone
from config.settings import (
    AZURE_CLIENT_ID, AZURE_CLIENT_SECRET, AZURE_TENANT_ID,
    USER_EMAIL, GRAPH_API_BASE, GRAPH_SCOPES
)


class OutlookClient:
    def __init__(self):
        self._app = msal.ConfidentialClientApplication(
            AZURE_CLIENT_ID,
            authority=f"https://login.microsoftonline.com/{AZURE_TENANT_ID}",
            client_credential=AZURE_CLIENT_SECRET,
        )
        self._token = None

    def _get_token(self):
        result = self._app.acquire_token_for_client(scopes=GRAPH_SCOPES)
        if "access_token" not in result:
            raise RuntimeError(f"Could not acquire token: {result.get('error_description')}")
        return result["access_token"]

    def _headers(self):
        if not self._token:
            self._token = self._get_token()
        return {"Authorization": f"Bearer {self._token}", "Content-Type": "application/json"}

    def _get(self, url, params=None):
        resp = requests.get(url, headers=self._headers(), params=params, timeout=30)
        if resp.status_code == 401:
            self._token = self._get_token()
            resp = requests.get(url, headers=self._headers(), params=params, timeout=30)
        resp.raise_for_status()
        return resp.json()

    def get_messages(self, top=200, folder="inbox"):
        """Fetch recent messages from inbox."""
        url = f"{GRAPH_API_BASE}/users/{USER_EMAIL}/mailFolders/{folder}/messages"
        params = {
            "$top": top,
            "$select": "id,subject,from,toRecipients,ccRecipients,receivedDateTime,bodyPreview,conversationId,isRead",
            "$orderby": "receivedDateTime desc",
        }
        data = self._get(url, params)
        return data.get("value", [])

    def get_sent_messages(self, top=200):
        """Fetch sent messages to capture outbound contacts."""
        return self.get_messages(top=top, folder="sentitems")

    def get_all_messages(self, top=200):
        """Combine inbox and sent items."""
        inbox = self.get_messages(top=top)
        sent = self.get_sent_messages(top=top)
        return inbox + sent

    def get_conversation_messages(self, conversation_id):
        """Fetch all messages in a specific conversation thread."""
        url = f"{GRAPH_API_BASE}/users/{USER_EMAIL}/messages"
        params = {
            "$filter": f"conversationId eq '{conversation_id}'",
            "$select": "id,subject,from,toRecipients,receivedDateTime,bodyPreview,isRead",
            "$orderby": "receivedDateTime asc",
        }
        data = self._get(url, params)
        return data.get("value", [])

    def search_messages(self, query, top=100):
        """Search messages by keyword (company name, contact name)."""
        url = f"{GRAPH_API_BASE}/users/{USER_EMAIL}/messages"
        params = {
            "$search": f'"{query}"',
            "$top": top,
            "$select": "id,subject,from,toRecipients,receivedDateTime,bodyPreview,conversationId,isRead",
        }
        data = self._get(url, params)
        return data.get("value", [])


def parse_email_address(address_obj):
    """Extract name and email from a Graph API address object."""
    if not address_obj:
        return "", ""
    ea = address_obj.get("emailAddress", {})
    return ea.get("name", ""), ea.get("address", "")


def group_messages_by_domain(messages):
    """Group messages by sender/recipient domain to identify companies."""
    domain_map = {}
    for msg in messages:
        from_name, from_addr = parse_email_address(msg.get("from"))
        domain = from_addr.split("@")[-1] if "@" in from_addr else ""
        if not domain:
            continue
        if domain not in domain_map:
            domain_map[domain] = {"domain": domain, "contacts": set(), "messages": []}
        domain_map[domain]["contacts"].add(f"{from_name} <{from_addr}>")
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
