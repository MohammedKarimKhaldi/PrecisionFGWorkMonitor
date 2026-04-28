"""Manage a local Excel workbook for the fundraising pipeline."""
import os
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

from config.settings import (
    EXCEL_FILE_PATH,
    EXCEL_COMPANIES_SHEET,
    EXCEL_EMAIL_LOG_SHEET,
    MANDATE_STATUSES,
    OUTLOOK_EMAIL,
)

STATUS_COLORS = {
    "Initial Contact":  "FFF9C4",
    "In Discussion":    "BBDEFB",
    "Proposal Sent":    "C8E6C9",
    "Due Diligence":    "E1BEE7",
    "Mandate Received": "A5D6A7",
    "Follow Up":        "FFE0B2",
    "On Hold":          "F5F5F5",
    "Not Interested":   "FFCDD2",
    "Closed – Won":     "1B5E20",
    "Closed – Lost":    "B71C1C",
}

_HEADER_FILL = PatternFill(start_color="1F3864", end_color="1F3864", fill_type="solid")
_HEADER_FONT = Font(bold=True, color="FFFFFF", size=11)
_THIN        = Side(style="thin", color="FFFFFF")
_BORDER      = Border(left=_THIN, right=_THIN)


def _write_headers(ws, headers):
    for col, h in enumerate(headers, 1):
        c = ws.cell(row=1, column=col, value=h)
        c.font = _HEADER_FONT
        c.fill = _HEADER_FILL
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = _BORDER
    ws.row_dimensions[1].height = 20


def _init_companies_sheet(ws):
    _write_headers(ws, [
        "Company", "Contact Name", "Contact Email",
        "Status", "Last Email Date", "Mandate Type",
        "AUM (M€)", "Notes", "First Contact Date", "Updated",
    ])
    widths = [28, 22, 30, 20, 18, 22, 12, 40, 18, 18]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[ws.cell(1, i).column_letter].width = w


def _init_email_log_sheet(ws):
    _write_headers(ws, [
        "Date", "Company", "Contact", "Direction",
        "Subject", "Preview", "Message-ID",
    ])
    widths = [18, 25, 30, 10, 45, 60, 36]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[ws.cell(1, i).column_letter].width = w


def _load_or_create():
    path = os.path.abspath(EXCEL_FILE_PATH)
    if os.path.exists(path):
        return openpyxl.load_workbook(path)
    wb = openpyxl.Workbook()
    ws_co = wb.active
    ws_co.title = EXCEL_COMPANIES_SHEET
    _init_companies_sheet(ws_co)
    ws_log = wb.create_sheet(EXCEL_EMAIL_LOG_SHEET)
    _init_email_log_sheet(ws_log)
    wb.save(path)
    return wb


def _save(wb):
    wb.save(os.path.abspath(EXCEL_FILE_PATH))


def _apply_row_color(ws, row_num, col_count, status):
    color = STATUS_COLORS.get(status, "FFFFFF")
    fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
    dark = color in ("1B5E20", "B71C1C")
    for col in range(1, col_count + 1):
        cell = ws.cell(row=row_num, column=col)
        cell.fill = fill
        cell.font = Font(color="FFFFFF" if dark else "000000")


# ── Public API ─────────────────────────────────────────────────────────────

def get_companies():
    wb = _load_or_create()
    ws = wb[EXCEL_COMPANIES_SHEET]
    headers = [c.value for c in ws[1]]
    rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0]:
            rows.append(dict(zip(headers, row)))
    return rows


def upsert_company(data):
    """Insert or update a company row matched by name."""
    wb = _load_or_create()
    ws = wb[EXCEL_COMPANIES_SHEET]
    headers = [c.value for c in ws[1]]
    data["Updated"] = datetime.now().strftime("%Y-%m-%d %H:%M")

    target = None
    for row in ws.iter_rows(min_row=2):
        if row[0].value and row[0].value.strip().lower() == data.get("Company", "").strip().lower():
            target = row[0].row
            break

    if target:
        for col, h in enumerate(headers, 1):
            if h in data:
                ws.cell(row=target, column=col, value=data[h])
    else:
        if "First Contact Date" not in data:
            data["First Contact Date"] = datetime.now().strftime("%Y-%m-%d")
        ws.append([data.get(h, "") for h in headers])
        target = ws.max_row

    _apply_row_color(ws, target, len(headers), data.get("Status", ""))
    _save(wb)


def update_company_status(company_name, new_status, notes=None):
    patch = {"Company": company_name, "Status": new_status}
    if notes:
        patch["Notes"] = notes
    upsert_company(patch)


def log_emails(emails, company_name):
    """Append emails to Email Log sheet, skipping duplicates by Message-ID."""
    wb = _load_or_create()
    ws = wb[EXCEL_EMAIL_LOG_SHEET]

    existing_ids = {
        str(row[6]) for row in ws.iter_rows(min_row=2, values_only=True) if row[6]
    }

    own = OUTLOOK_EMAIL.lower()
    added = 0
    for em in emails:
        msg_id = str(em.get("id") or "")
        if msg_id and msg_id in existing_ids:
            continue
        ea = (em.get("from") or {}).get("emailAddress", {})
        from_addr = ea.get("address", "")
        direction = "OUT" if own in from_addr.lower() else "IN"
        ws.append([
            (em.get("receivedDateTime") or "")[:10],
            company_name,
            f"{ea.get('name', '')} <{from_addr}>".strip(),
            direction,
            em.get("subject", ""),
            (em.get("bodyPreview") or "")[:200],
            msg_id,
        ])
        existing_ids.add(msg_id)
        added += 1

    if added:
        _save(wb)
    return added


def get_excel_path():
    return os.path.abspath(EXCEL_FILE_PATH)
