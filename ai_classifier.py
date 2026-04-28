"""AI-powered deal classification using Claude API (claude-opus-4-7)."""
import json
import anthropic
from config.settings import ANTHROPIC_API_KEY, MANDATE_STATUSES

_client = None


def _get_client():
    global _client
    if _client is None:
        if not ANTHROPIC_API_KEY:
            raise RuntimeError(
                "ANTHROPIC_API_KEY is not set. Add it to your .env file."
            )
        _client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
    return _client


_STATUSES_LIST = "\n".join(f"  - {s}" for s in MANDATE_STATUSES)

_SYSTEM_PROMPT = f"""You are an expert fundraising analyst for a hedge fund.
You analyse email threads between a fundraiser and potential institutional investors/allocators.
Extract structured deal information from the email data provided.

Pipeline statuses (use exactly one):
{_STATUSES_LIST}

Guidelines:
- "Company" should be the full institution name, not just the email domain.
  Infer it from the sender name, subject lines, or domain (e.g. amundi.com → Amundi Asset Management).
- "Contact Name" / "Contact Email": primary point of contact (most frequent or most senior).
- "Status": pick the stage that best reflects the latest interaction.
  • No reply / single intro email → "Initial Contact"
  • Back-and-forth exchange → "In Discussion"
  • Proposal / deck sent → "Proposal Sent"
  • Requesting data, DDQ, references → "Due Diligence"
  • Mandate confirmed → "Mandate Received"
  • Awaiting next meeting / nudging → "Follow Up"
- "Mandate Type": e.g. "Equity Long/Short", "Global Macro", "Fixed Income Relative Value".
  Use null if unknown.
- "AUM (M€)": a number in millions of euros if mentioned, otherwise omit the field.
- "Notes": 1-2 sentence summary of where the relationship stands and any next steps.

Return ONLY a valid JSON object — no markdown, no explanation."""


def classify_domain_emails(domain: str, emails: list, contacts: list) -> dict:
    """Send an email group to Claude and return a structured deal dict."""
    lines = []
    for em in sorted(emails, key=lambda e: e.get("receivedDateTime", ""), reverse=True)[:25]:
        ea = (em.get("from") or {}).get("emailAddress", {})
        from_str = f"{ea.get('name', '')} <{ea.get('address', '')}>".strip(" <>")
        date = (em.get("receivedDateTime") or "")[:10]
        subject = em.get("subject") or "(no subject)"
        preview = (em.get("bodyPreview") or "").strip()
        folder = em.get("folder", "")
        line = f"[{date}] [{folder}] From: {from_str}\nSubject: {subject}"
        if preview:
            line += f"\nPreview: {preview[:180]}"
        lines.append(line)

    emails_text = "\n\n---\n\n".join(lines) if lines else "No emails available."
    contacts_text = ", ".join(contacts) if contacts else "unknown"

    user_msg = (
        f'Domain: {domain}\n'
        f'Known contacts: {contacts_text}\n'
        f'Email count: {len(emails)}\n\n'
        f'EMAIL THREAD (most recent first):\n{emails_text}\n\n'
        f'Return a JSON object classifying this deal.'
    )

    response = _get_client().messages.create(
        model="claude-opus-4-7",
        max_tokens=512,
        system=[
            {
                "type": "text",
                "text": _SYSTEM_PROMPT,
                "cache_control": {"type": "ephemeral"},
            }
        ],
        messages=[{"role": "user", "content": user_msg}],
    )

    raw = response.content[0].text.strip()
    if raw.startswith("```"):
        raw = raw.split("```")[1]
        if raw.startswith("json"):
            raw = raw[4:]
        raw = raw.rsplit("```", 1)[0]

    result = json.loads(raw.strip())

    if result.get("Status") not in MANDATE_STATUSES:
        result["Status"] = "Initial Contact"

    if "AUM (M€)" in result:
        try:
            result["AUM (M€)"] = float(result["AUM (M€)"])
        except (TypeError, ValueError):
            del result["AUM (M€)"]

    return result
