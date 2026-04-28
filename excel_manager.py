"""Manage the shared Excel workbook via Microsoft Graph API (OneDrive)."""
import io
import requests
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from config.settings import (
    AZURE_CLIENT_ID, AZURE_CLIENT_SECRET, AZURE_TENANT_ID,
    USER_EMAIL, EXCEL_FILE_PATH, GRAPH_API_BASE,
    EXCEL_COMPANIES_SHEET, EXCEL_EMAIL_LOG_SHEET, MANDATE_STATUSES,
)


STATUS_COLORS = {
    "Initial Contact":  "FFF9C4",  # yellow
    "In Discussion":    "BBDEFB",  # blue
    "Proposal Sent":    "C8E6C9",  # light green
    "Due Diligence":    "E1BEE7",  # purple
    "Mandate Received": "A5D6A7",  # green
    "Follow Up":        "FFE0B2",  # orange
    "On Hold":          "F5F5F5",  # grey
    "Not Interested":   "FFCDD2",  # red
    "Closed – Won":     "1B5E20",  # dark green (white text)
    "Closed – Lost":    "B71C1C",  # dark red (white text)
}

HEADER_FILL = PatternFill(start_color="1F3864", end_color="1F3864", fill_type="solid")
HEADER_FONT = Font(bold=True, color="FFFFFF", size=11)


class ExcelManager:
    def __init__(self, outlook_client=None):
        self._graph_client = outlook_client

    # ------------------------------------------------------------------
    # OneDrive upload / download
    # ------------------------------------------------------------------

    def _drive_url(self):
        encoded = EXCEL_FILE_PATH.replace("/", "%2F").replace(" ", "%20")
        return f"{GRAPH_API_BASE}/users/{USER_EMAIL}/drive/root:{encoded}"

    def download_workbook(self):
        """Download the Excel file from OneDrive. Returns openpyxl Workbook."""
        url = f"{self._drive_url()}:/content"
        resp = requests.get(url, headers=self._graph_client._headers(), timeout=30)
        if resp.status_code == 404:
            return self._create_empty_workbook()
        resp.raise_for_status()
        return openpyxl.load_workbook(io.BytesIO(resp.content))

    def upload_workbook(self, wb):
        """Upload workbook back to OneDrive."""
        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        url = f"{self._drive_url()}:/content"
        headers = self._graph_client._headers()
        headers["Content-Type"] = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        resp = requests.put(url, headers=headers, data=buf.read(), timeout=60)
        resp.raise_for_status()
        return resp.json()

    def get_workbook_bytes(self):
        """Return raw bytes of the current workbook for download."""
        url = f"{self._drive_url()}:/content"
        resp = requests.get(url, headers=self._graph_client._headers(), timeout=30)
        if resp.status_code == 404:
            wb = self._create_empty_workbook()
            buf = io.BytesIO()
            wb.save(buf)
            return buf.getvalue()
        resp.raise_for_status()
        return resp.content

    # ------------------------------------------------------------------
    # Workbook initialisation
    # ------------------------------------------------------------------

    def _create_empty_workbook(self):
        wb = openpyxl.Workbook()
        ws_companies = wb.active
        ws_companies.title = EXCEL_COMPANIES_SHEET
        ws_log = wb.create_sheet(EXCEL_EMAIL_LOG_SHEET)
        self._init_companies_sheet(ws_companies)
        self._init_email_log_sheet(ws_log)
        return wb

    def _init_companies_sheet(self, ws):
        headers = [
            "Company", "Contact Name", "Contact Email",
            "Status", "Last Email Date", "Mandate Type",
            "AUM (M€)", "Notes", "First Contact Date", "Updated",
        ]
        self._write_headers(ws, headers)
        ws.column_dimensions["A"].width = 28
        ws.column_dimensions["B"].width = 22
        ws.column_dimensions["C"].width = 30
        ws.column_dimensions["D"].width = 20
        ws.column_dimensions["E"].width = 18
        ws.column_dimensions["F"].width = 22
        ws.column_dimensions["G"].width = 12
        ws.column_dimensions["H"].width = 40
        ws.column_dimensions["I"].width = 18
        ws.column_dimensions["J"].width = 18

    def _init_email_log_sheet(self, ws):
        headers = [
            "Date", "Company", "Contact", "Direction",
            "Subject", "Preview", "Conversation ID",
        ]
        self._write_headers(ws, headers)
        ws.column_dimensions["A"].width = 18
        ws.column_dimensions["B"].width = 25
        ws.column_dimensions["C"].width = 30
        ws.column_dimensions["D"].width = 10
        ws.column_dimensions["E"].width = 45
        ws.column_dimensions["F"].width = 60
        ws.column_dimensions["G"].width = 36

    def _write_headers(self, ws, headers):
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=header)
            cell.font = HEADER_FONT
            cell.fill = HEADER_FILL
            cell.alignment = Alignment(horizontal="center", vertical="center")
            thin = Side(style="thin", color="FFFFFF")
            cell.border = Border(left=thin, right=thin)
        ws.row_dimensions[1].height = 20

    # ------------------------------------------------------------------
    # Companies CRUD
    # ------------------------------------------------------------------

    def get_companies(self):
        wb = self.download_workbook()
        ws = wb[EXCEL_COMPANIES_SHEET]
        companies = []
        headers = [cell.value for cell in ws[1]]
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row[0]:
                continue
            companies.append(dict(zip(headers, row)))
        return companies

    def upsert_company(self, company_data):
        """Insert or update a company row. Match on 'Company' name."""
        wb = self.download_workbook()
        ws = wb[EXCEL_COMPANIES_SHEET]
        headers = [cell.value for cell in ws[1]]

        target_row = None
        for row in ws.iter_rows(min_row=2):
            if row[0].value and row[0].value.strip().lower() == company_data.get("Company", "").strip().lower():
                target_row = row[0].row
                break

        company_data["Updated"] = datetime.utcnow().strftime("%Y-%m-%d %H:%M")

        if target_row:
            for col, header in enumerate(headers, 1):
                if header in company_data:
                    ws.cell(row=target_row, column=col, value=company_data[header])
        else:
            if "First Contact Date" not in company_data:
                company_data["First Contact Date"] = datetime.utcnow().strftime("%Y-%m-%d")
            new_row = [company_data.get(h, "") for h in headers]
            ws.append(new_row)
            target_row = ws.max_row

        self._apply_status_color(ws, target_row, headers, company_data.get("Status", ""))
        self.upload_workbook(wb)

    def update_company_status(self, company_name, new_status, notes=None):
        data = {"Company": company_name, "Status": new_status}
        if notes is not None:
            data["Notes"] = notes
        self.upsert_company(data)

    def _apply_status_color(self, ws, row_num, headers, status):
        color = STATUS_COLORS.get(status, "FFFFFF")
        fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
        font_color = "FFFFFF" if color in ("1B5E20", "B71C1C") else "000000"
        for col in range(1, len(headers) + 1):
            cell = ws.cell(row=row_num, column=col)
            cell.fill = fill
            cell.font = Font(color=font_color)

    # ------------------------------------------------------------------
    # Email Log
    # ------------------------------------------------------------------

    def log_emails(self, emails, company_name):
        """Append emails to the Email Log sheet (skip duplicates by subject+date)."""
        wb = self.download_workbook()
        ws = wb[EXCEL_EMAIL_LOG_SHEET]

        existing = set()
        for row in ws.iter_rows(min_row=2, values_only=True):
            if row[0] and row[4]:
                existing.add((str(row[0])[:10], str(row[4])[:50]))

        added = 0
        for email in emails:
            date_str = email.get("receivedDateTime", "")[:10]
            subject = (email.get("subject") or "")[:50]
            if (date_str, subject) in existing:
                continue
            from_name, from_addr = _parse_addr(email.get("from"))
            direction = "IN" if USER_EMAIL.lower() not in from_addr.lower() else "OUT"
            ws.append([
                date_str,
                company_name,
                f"{from_name} <{from_addr}>",
                direction,
                email.get("subject", ""),
                (email.get("bodyPreview") or "")[:200],
                email.get("conversationId", ""),
            ])
            existing.add((date_str, subject))
            added += 1

        if added:
            self.upload_workbook(wb)
        return added


def _parse_addr(address_obj):
    if not address_obj:
        return "", ""
    ea = address_obj.get("emailAddress", {})
    return ea.get("name", ""), ea.get("address", "")
