# PrecisionFG Fundraising Email Monitor

A web app (open in Chrome) that reads your Outlook inbox via **Outlook for Mac** and tracks
fundraising mandate discussions with asset managers in a **local Excel workbook**.

No Azure App registration required.

## Features

- **Outlook integration** — reads inbox + sent items from the local Outlook desktop app
- **Pipeline dashboard** — companies with status badges, contact, mandate type, AUM
- **Workflow queue** — highlights who needs a reply, who has not answered, and recent meeting signals
- **Status tracking** — 10-stage pipeline from Initial Contact to Closed
- **Email thread view** — click any company to see matching email exchange
- **Local Excel workbook** — auto-created as `FundraisingTracker.xlsx`; two sheets: *Companies* and *Email Log*
- **Log emails to Excel** — one click to append the current email thread to the log sheet
- **Search & filter** — filter by status or search by company/contact name

## Setup

### 1. Open Outlook for Mac

The app reads local Outlook data through macOS automation. Keep Outlook open when
refreshing emails. If Outlook blocks automation, allow Terminal or your IDE in
macOS System Settings → Privacy & Security → Automation.

### 2. Configure

```bash
cp .env.example .env
```

Edit `.env`:

```
OUTLOOK_EMAIL=you@company.com
EXCEL_FILE_PATH=./FundraisingTracker.xlsx
EMAIL_CACHE_PATH=./email_cache.json
FOLLOW_UP_DAYS=5
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

## Email Cache

Email summaries are saved locally at `EMAIL_CACHE_PATH` after you click
**Refresh Emails**. Normal page loads use that file only, so restarting the app is
fast and does not re-query Outlook.

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

## Workflow Labels

These labels are computed from the latest matched Outlook thread and do not replace
the Excel deal stage.

| Label | Meaning |
|---|---|
| Needs reply | Their latest email is newer than yours |
| No answer yet | Your latest email is older than `FOLLOW_UP_DAYS` |
| Waiting | Your latest email is still within the follow-up window |
| Meeting signal | A recent call, meeting, Zoom, Teams, accepted, or tentative invite was found |
| No email found | The company has no matched email thread yet |
