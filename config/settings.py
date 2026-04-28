import os
from dotenv import load_dotenv

load_dotenv()

AZURE_CLIENT_ID = os.environ.get("AZURE_CLIENT_ID", "")
AZURE_CLIENT_SECRET = os.environ.get("AZURE_CLIENT_SECRET", "")
AZURE_TENANT_ID = os.environ.get("AZURE_TENANT_ID", "")
USER_EMAIL = os.environ.get("USER_EMAIL", "")
EXCEL_FILE_PATH = os.environ.get("EXCEL_FILE_PATH", "/Documents/FundraisingTracker.xlsx")
FLASK_SECRET_KEY = os.environ.get("FLASK_SECRET_KEY", "dev-secret-key")

GRAPH_API_BASE = "https://graph.microsoft.com/v1.0"
GRAPH_SCOPES = ["https://graph.microsoft.com/.default"]

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
