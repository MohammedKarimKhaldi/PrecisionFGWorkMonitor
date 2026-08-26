import os
from dotenv import load_dotenv

BASE_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), os.pardir))


def _project_path(path):
    path = os.path.expanduser(path)
    if os.path.isabs(path):
        return path
    return os.path.abspath(os.path.join(BASE_DIR, path))


load_dotenv(os.path.join(BASE_DIR, ".env"))

OUTLOOK_EMAIL    = os.environ.get("OUTLOOK_EMAIL", "")
OUTLOOK_PASSWORD = os.environ.get("OUTLOOK_PASSWORD", "")
IMAP_SERVER      = os.environ.get("IMAP_SERVER", "")  # auto-detected if blank
IMAP_PORT        = 993
INTERNAL_EMAIL_DOMAINS = [
    d.strip().lower()
    for d in os.environ.get(
        "INTERNAL_EMAIL_DOMAINS",
        "precision-familygroup.com,plutus-investment.com",
    ).split(",")
    if d.strip()
]

EXCEL_FILE_PATH    = os.environ.get("EXCEL_FILE_PATH",    "./FundraisingTracker.xlsx")
EMAIL_CACHE_PATH   = _project_path(os.environ.get("EMAIL_CACHE_PATH") or "./email_cache.json")
FLASK_SECRET_KEY   = os.environ.get("FLASK_SECRET_KEY",   "dev-secret-key")

OLLAMA_HOST  = os.environ.get("OLLAMA_HOST",  "http://localhost:11434")
OLLAMA_MODEL = os.environ.get("OLLAMA_MODEL", "llama3.2")
FOLLOW_UP_DAYS = int(os.environ.get("FOLLOW_UP_DAYS", "5"))

MANDATE_STATUSES = [
    "Initial Contact",
    "In Discussion",
    "Proposal Sent",
    "Due Diligence",
    "Mandate Received",
    "Follow Up",
    "On Hold",
    "Not Interested",
    "Closed – Won",
    "Closed – Lost",
]

EXCEL_COMPANIES_SHEET = "Companies"
EXCEL_EMAIL_LOG_SHEET = "Email Log"
