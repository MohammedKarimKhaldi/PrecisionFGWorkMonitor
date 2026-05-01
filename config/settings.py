import os
from dotenv import load_dotenv

load_dotenv()

OUTLOOK_EMAIL    = os.environ.get("OUTLOOK_EMAIL", "")
OUTLOOK_PASSWORD = os.environ.get("OUTLOOK_PASSWORD", "")
IMAP_SERVER      = os.environ.get("IMAP_SERVER", "")  # auto-detected if blank
IMAP_PORT        = 993

EXCEL_FILE_PATH    = os.environ.get("EXCEL_FILE_PATH",    "./FundraisingTracker.xlsx")
EMAIL_CACHE_PATH   = os.environ.get("EMAIL_CACHE_PATH",   "./email_cache.json")
FLASK_SECRET_KEY   = os.environ.get("FLASK_SECRET_KEY",   "dev-secret-key")

OLLAMA_HOST  = os.environ.get("OLLAMA_HOST",  "http://localhost:11434")
OLLAMA_MODEL = os.environ.get("OLLAMA_MODEL", "llama3.2")

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
