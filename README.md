# PrecisionFG Fundraising Email Monitor

A web app that reads your Outlook inbox via Microsoft Graph API and tracks fundraising mandate discussions with asset managers in a shared Excel workbook on OneDrive.

## Features

- **Outlook integration** – reads inbox + sent items via Microsoft Graph API (OAuth2 delegated auth)
- **Pipeline dashboard** – table of all companies with status badges, contact info, mandate type, AUM
- **Status tracking** – 10-stage pipeline: Initial Contact → In Discussion → Proposal Sent → Due Diligence → Mandate Received / Closed
- **Email thread view** – click a company to see the full email exchange
- **Shared Excel workbook** – automatically created on your OneDrive; two sheets: *Companies* and *Email Log*
- **Live search & filter** – filter pipeline by status or search by company/contact name

## Setup

### 1. Register an Azure AD App

1. Go to [Azure Portal → App registrations](https://portal.azure.com/#blade/Microsoft_AAD_RegisteredApp/ApplicationsListBlade)
2. **New registration** → name it "PrecisionFG Monitor" → Redirect URI: `http://localhost:5000/auth/callback`
3. Copy the **Application (client) ID** and **Directory (tenant) ID**
4. Go to **Certificates & secrets** → New client secret → copy the value
5. Go to **API permissions** → Add: `Mail.Read`, `Mail.ReadBasic`, `Files.ReadWrite`, `User.Read` (all Delegated)
6. Click **Grant admin consent**

### 2. Configure

```bash
cp .env.example .env
# Fill in AZURE_CLIENT_ID, AZURE_CLIENT_SECRET, AZURE_TENANT_ID, USER_EMAIL
```

### 3. Install & run

```bash
python -m venv .venv
source .venv/bin/activate      # Windows: .venv\Scripts\activate
pip install -r requirements.txt
python app.py
```

Open [http://localhost:5000](http://localhost:5000) → Sign in with Microsoft.

## Excel Workbook

The workbook is created automatically at the path set in `EXCEL_FILE_PATH` (default: `/Documents/FundraisingTracker.xlsx` on your OneDrive).

| Sheet | Columns |
|---|---|
| **Companies** | Company, Contact Name, Contact Email, Status, Last Email Date, Mandate Type, AUM (M€), Notes, First Contact Date, Updated |
| **Email Log** | Date, Company, Contact, Direction (IN/OUT), Subject, Preview, Conversation ID |

Status colours are applied automatically (green = active, red = lost, etc.).

## Pipeline Statuses

| Status | Meaning |
|---|---|
| Initial Contact | First outreach sent |
| In Discussion | Active conversation ongoing |
| Proposal Sent | Mandate proposal / deck delivered |
| Due Diligence | Company is reviewing / doing DD |
| Mandate Received | Mandate confirmed |
| Follow Up | Waiting for response |
| On Hold | Paused by either side |
| Not Interested | Declined |
| Closed – Won | Mandate executed |
| Closed – Lost | Deal did not close |
