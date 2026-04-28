"""AI-powered deal classification using a local Ollama model.
Runs entirely on your machine — no data sent anywhere.
Install: https://ollama.ai  then  `ollama pull llama3.2`
"""
import json
import urllib.request
import urllib.error
from config.settings import OLLAMA_HOST, OLLAMA_MODEL, MANDATE_STATUSES

_STATUSES_LIST = "\n".join(f"  - {s}" for s in MANDATE_STATUSES)

_SYSTEM_PROMPT = f"""You are an expert fundraising analyst for a hedge fund.
Analyse email threads between a fundraiser and potential institutional investors/allocators.
Extract structured deal information and return ONLY a valid JSON object — no markdown, no explanation.

JSON fields:
  "Company"       : full institution name (infer from domain if needed, e.g. amundi.com → Amundi Asset Management)
  "Contact Name"  : primary contact's full name
  "Contact Email" : primary contact's email address
  "Status"        : exactly one of the statuses listed below
  "Mandate Type"  : e.g. "Equity Long/Short", "Global Macro", "Fixed Income" — omit if unknown
  "AUM (M€)"      : number in millions of euros — omit if not mentioned
  "Notes"         : 1-2 sentences on current relationship status and next steps

Pipeline statuses:
{_STATUSES_LIST}

Status selection guide:
  single intro / no reply yet → "Initial Contact"
  back-and-forth emails       → "In Discussion"
  pitch deck / proposal sent  → "Proposal Sent"
  DDQ / data request          → "Due Diligence"
  mandate confirmed           → "Mandate Received"
  waiting / nudging           → "Follow Up"

Return ONLY the JSON object. No prose, no markdown fences."""


def _call_ollama(prompt: str) -> str:
    payload = json.dumps({
        "model":  OLLAMA_MODEL,
        "format": "json",
        "stream": False,
        "messages": [
            {"role": "system", "content": _SYSTEM_PROMPT},
            {"role": "user",   "content": prompt},
        ],
    }).encode()

    url = f"{OLLAMA_HOST.rstrip('/')}/api/chat"
    req = urllib.request.Request(
        url, data=payload, method="POST",
        headers={"Content-Type": "application/json"},
    )
    try:
        with urllib.request.urlopen(req, timeout=120) as resp:
            body = json.loads(resp.read())
            return body["message"]["content"]
    except urllib.error.URLError as e:
        raise RuntimeError(
            f"Cannot reach Ollama at {OLLAMA_HOST}. "
            "Make sure it is running (`ollama serve`) and that you have pulled a model "
            f"(`ollama pull {OLLAMA_MODEL}`)."
        ) from e


def classify_domain_emails(domain: str, emails: list, contacts: list) -> dict:
    """Classify a domain's email group and return a structured deal dict."""
    lines = []
    for em in sorted(emails, key=lambda e: e.get("receivedDateTime", ""), reverse=True)[:25]:
        ea = (em.get("from") or {}).get("emailAddress", {})
        from_str = f"{ea.get('name', '')} <{ea.get('address', '')}>".strip(" <>")
        date    = (em.get("receivedDateTime") or "")[:10]
        subject = em.get("subject") or "(no subject)"
        preview = (em.get("bodyPreview") or "").strip()
        folder  = em.get("folder", "")
        line = f"[{date}] [{folder}] From: {from_str}\nSubject: {subject}"
        if preview:
            line += f"\nPreview: {preview[:180]}"
        lines.append(line)

    emails_text   = "\n\n---\n\n".join(lines) if lines else "No emails available."
    contacts_text = ", ".join(contacts) if contacts else "unknown"

    prompt = (
        f"Domain: {domain}\n"
        f"Known contacts: {contacts_text}\n"
        f"Email count: {len(emails)}\n\n"
        f"EMAIL THREAD (most recent first):\n{emails_text}\n\n"
        f"Return a JSON object classifying this deal."
    )

    raw = _call_ollama(prompt).strip()

    # Strip markdown fences if the model added them despite instructions
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


def test_ollama() -> tuple[bool, str]:
    """Quick connectivity check — returns (ok, message)."""
    try:
        url = f"{OLLAMA_HOST.rstrip('/')}/api/tags"
        with urllib.request.urlopen(url, timeout=5) as resp:
            data = json.loads(resp.read())
            models = [m["name"] for m in data.get("models", [])]
            if not models:
                return False, f"Ollama is running but no models pulled. Run: ollama pull {OLLAMA_MODEL}"
            if not any(OLLAMA_MODEL.split(":")[0] in m for m in models):
                return False, (f"Model '{OLLAMA_MODEL}' not found. "
                               f"Available: {', '.join(models)}. "
                               f"Run: ollama pull {OLLAMA_MODEL}")
            return True, f"Ollama ready — model: {OLLAMA_MODEL} | available: {', '.join(models)}"
    except Exception as e:
        return False, (f"Cannot reach Ollama at {OLLAMA_HOST}. "
                       f"Install from https://ollama.ai and run: ollama serve && ollama pull {OLLAMA_MODEL}")
