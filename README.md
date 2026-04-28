# PrecisionFG Fundraising Email Monitor

A web app (open in Chrome) that reads your Outlook inbox via **IMAP** and tracks
fundraising mandate discussions with asset managers in a **local Excel workbook**.

No Azure App registration required.

## Features

- **Outlook integration** — reads inbox + sent items via standard IMAP protocol
- **Pipeline dashboard** — companies with status badges, contact, mandate type, AUM
- **Status tracking** — 10-stage pipeline from Initial Contact to Closed
- **Email thread view** — click any company to see matching email exchange
- **Local Excel workbook** — auto-created as `FundraisingTracker.xlsx`; two sheets: *Companies* and *Email Log*
- **Log emails to Excel** — one click to append the current email thread to the log sheet
- **Search & filter** — filter by status or search by company/contact name

## Setup

### 1. Enable IMAP in Outlook

**Outlook.com / Hotmail / Live (personal account)**
1. Go to Outlook.com → Settings → Mail → Sync email → POP and IMAP
2. Enable IMAP access
3. If you have 2-factor authentication: go to [account.microsoft.com/security](https://account.microsoft.com/security) → Advanced security → App passwords → create one

**Microsoft 365 work account**
- IMAP server: `outlook.office365.com`
- Basic auth must be enabled for your mailbox by your IT admin, or use an app password
- Ask your IT department to enable "Authenticated SMTP" / IMAP for your account

### 2. Configure

```bash
cp .env.example .env
```

Edit `.env`:

```
OUTLOOK_EMAIL=you@company.com
OUTLOOK_PASSWORD=your-app-password
# IMAP_SERVER=outlook.office365.com   ← override if auto-detection fails
EXCEL_FILE_PATH=./FundraisingTracker.xlsx
```

### 3. Install & run

```bash
python -m venv .venv
source .venv/bin/activate      # Windows: .venv\Scripts\activate
pip install -r requirements.txt
python app.py
```

Open **http://localhost:5000** in Chrome.

## Excel Workbook

The file is created automatically at `EXCEL_FILE_PATH`.  
To share it with colleagues, point `EXCEL_FILE_PATH` to a shared network drive.

| Sheet | Columns |
|---|---|
| **Companies** | Company, Contact Name, Contact Email, Status, Last Email Date, Mandate Type, AUM (M€), Notes, First Contact Date, Updated |
| **Email Log** | Date, Company, Contact, Direction (IN/OUT), Subject, Preview, Message-ID |

Status colours are applied automatically.

## Pipeline Statuses

| Status | Meaning |
|---|---|
| Initial Contact | First outreach sent |
| In Discussion | Active conversation |
| Proposal Sent | Mandate proposal delivered |
| Due Diligence | Company is doing DD |
| Mandate Received | Mandate confirmed |
| Follow Up | Waiting for reply |
| On Hold | Paused |
| Not Interested | Declined |
| Closed – Won | Mandate executed |
| Closed – Lost | Deal did not close |
