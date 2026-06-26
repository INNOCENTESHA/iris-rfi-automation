"""
Bolt Trips Tracker → Google Sheets Sync
========================================
Reads Bolt_Trips_Tracker.xlsx from C:\\AI-Projects\\ and creates/refreshes
a Google Spreadsheet named "Bolt Trips Tracker" in your Google Drive.

FIRST-TIME SETUP (run once):
  pip install gspread google-auth google-auth-oauthlib openpyxl

CREDENTIALS:
  Uses the existing C:\\AI-Projects\\google_credentials.json and
  google_token.json (same credentials as the RFI script).

Run:
  python C:\\AI-Projects\\sync_bolt_to_gsheets.py
"""

import os, sys, json
from pathlib import Path
from datetime import datetime

# ── Auto-install dependencies ─────────────────────────────────────────────────
def _ensure(pkg, import_name=None):
    import importlib, subprocess
    import_name = import_name or pkg
    try:
        importlib.import_module(import_name)
    except ImportError:
        print(f"Installing {pkg} …")
        subprocess.check_call([sys.executable, "-m", "pip", "install", pkg, "-q"])

for _p, _i in [("gspread", "gspread"), ("google-auth", "google.oauth2"),
               ("google-auth-oauthlib", "google_auth_oauthlib"),
               ("openpyxl", "openpyxl")]:
    _ensure(_p, _i)

import gspread
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import openpyxl

# ── Configuration ─────────────────────────────────────────────────────────────
CREDENTIALS_FILE  = r"C:\AI-Projects\google_credentials.json"
TOKEN_FILE        = r"C:\AI-Projects\google_token.json"
EXCEL_FILE        = r"C:\AI-Projects\Bolt_Trips_Tracker.xlsx"
SPREADSHEET_NAME  = "Bolt Trips Tracker"
SPREADSHEET_ID_FILE = r"C:\AI-Projects\bolt_gsheets_id.txt"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

# ── Auth ──────────────────────────────────────────────────────────────────────
def get_credentials():
    creds = None
    if os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(TOKEN_FILE, "w") as f:
            f.write(creds.to_json())
    return creds

# ── Read Excel data ───────────────────────────────────────────────────────────
def read_excel():
    wb = openpyxl.load_workbook(EXCEL_FILE, data_only=False)
    ws = wb.active

    # Header row is row 4
    headers = [ws.cell(4, c).value for c in range(1, 17)]

    rows = []
    for row in range(5, ws.max_row + 1):
        date_val = ws.cell(row, 2).value
        if date_val is None:
            break
        row_data = []
        for c in range(1, 17):
            val = ws.cell(row, c).value
            if hasattr(val, 'strftime'):
                val = val.strftime('%d-%b-%Y')
            elif val is None:
                val = ''
            else:
                val = str(val)
            row_data.append(val)
        rows.append(row_data)

    return headers, rows

# ── Find or create Google Sheet ───────────────────────────────────────────────
def get_or_create_sheet(gc):
    # Try cached ID first
    if os.path.exists(SPREADSHEET_ID_FILE):
        with open(SPREADSHEET_ID_FILE) as f:
            sid = f.read().strip()
        try:
            sh = gc.open_by_key(sid)
            print(f"Opened existing sheet: {sh.title} ({sid})")
            return sh
        except Exception:
            print("Cached sheet ID not found, searching Drive…")

    # Search by name
    try:
        sh = gc.open(SPREADSHEET_NAME)
        print(f"Found existing sheet by name: {sh.id}")
        with open(SPREADSHEET_ID_FILE, 'w') as f:
            f.write(sh.id)
        return sh
    except gspread.SpreadsheetNotFound:
        pass

    # Create new
    sh = gc.create(SPREADSHEET_NAME)
    sh.share(None, perm_type='anyone', role='reader')  # optional: make readable
    print(f"Created new sheet: {sh.id}")
    with open(SPREADSHEET_ID_FILE, 'w') as f:
        f.write(sh.id)
    return sh

# ── Format & Upload ───────────────────────────────────────────────────────────
def upload_to_sheets(sh, headers, rows):
    try:
        ws = sh.worksheet("Bolt Trips")
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title="Bolt Trips", rows=500, cols=20)
        # Remove default Sheet1 if it exists
        try:
            sh.del_worksheet(sh.worksheet("Sheet1"))
        except Exception:
            pass

    # Clear existing content
    ws.clear()

    # Title rows
    today = datetime.now().strftime('%d %B %Y')
    total = len(rows)
    title_data = [
        ["Bolt Trips Tracker - Comprehensive Report"] + [""] * 15,
        [f"Generated: {today} | Total Trips: {total}"] + [""] * 15,
        [""] * 16,
        headers,
    ]

    all_data = title_data + rows
    ws.update(all_data, value_input_option='USER_ENTERED')

    # ── Basic formatting ──────────────────────────────────────────────────────
    sheet_id = ws._properties['sheetId']
    requests_body = []

    # Merge title rows A1:P1 and A2:P2
    for row_idx in [0, 1]:
        requests_body.append({
            "mergeCells": {
                "range": {"sheetId": sheet_id, "startRowIndex": row_idx,
                          "endRowIndex": row_idx + 1, "startColumnIndex": 0,
                          "endColumnIndex": 16},
                "mergeType": "MERGE_ALL"
            }
        })

    # Header row (row 4, index 3) - green background
    requests_body.append({
        "repeatCell": {
            "range": {"sheetId": sheet_id, "startRowIndex": 3,
                      "endRowIndex": 4, "startColumnIndex": 0, "endColumnIndex": 16},
            "cell": {
                "userEnteredFormat": {
                    "backgroundColor": {"red": 0.18, "green": 0.49, "blue": 0.20},
                    "textFormat": {"foregroundColor": {"red": 1, "green": 1, "blue": 1},
                                   "bold": True, "fontFamily": "Arial", "fontSize": 10},
                    "horizontalAlignment": "CENTER",
                    "verticalAlignment": "MIDDLE",
                }
            },
            "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)"
        }
    })

    # Title row (row 1) formatting
    requests_body.append({
        "repeatCell": {
            "range": {"sheetId": sheet_id, "startRowIndex": 0, "endRowIndex": 1,
                      "startColumnIndex": 0, "endColumnIndex": 16},
            "cell": {
                "userEnteredFormat": {
                    "backgroundColor": {"red": 0.18, "green": 0.49, "blue": 0.20},
                    "textFormat": {"foregroundColor": {"red": 1, "green": 1, "blue": 1},
                                   "bold": True, "fontFamily": "Arial", "fontSize": 13},
                    "horizontalAlignment": "CENTER",
                }
            },
            "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)"
        }
    })

    # Row 2 formatting (subtitle)
    requests_body.append({
        "repeatCell": {
            "range": {"sheetId": sheet_id, "startRowIndex": 1, "endRowIndex": 2,
                      "startColumnIndex": 0, "endColumnIndex": 16},
            "cell": {
                "userEnteredFormat": {
                    "backgroundColor": {"red": 0.85, "green": 0.93, "blue": 0.83},
                    "textFormat": {"fontFamily": "Arial", "fontSize": 10, "italic": True},
                    "horizontalAlignment": "CENTER",
                }
            },
            "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)"
        }
    })

    # Data rows alternating color
    for i, _ in enumerate(rows):
        row_idx = 4 + i  # 0-indexed, data starts at row 5 (index 4)
        if i % 2 == 1:  # alternate rows
            requests_body.append({
                "repeatCell": {
                    "range": {"sheetId": sheet_id, "startRowIndex": row_idx,
                              "endRowIndex": row_idx + 1, "startColumnIndex": 0,
                              "endColumnIndex": 16},
                    "cell": {
                        "userEnteredFormat": {
                            "backgroundColor": {"red": 0.96, "green": 0.96, "blue": 0.96},
                        }
                    },
                    "fields": "userEnteredFormat(backgroundColor)"
                }
            })

    # Freeze top 4 rows and column A
    requests_body.append({
        "updateSheetProperties": {
            "properties": {
                "sheetId": sheet_id,
                "gridProperties": {"frozenRowCount": 4}
            },
            "fields": "gridProperties(frozenRowCount,frozenColumnCount)"
        }
    })

    # Auto-resize columns
    requests_body.append({
        "autoResizeDimensions": {
            "dimensions": {"sheetId": sheet_id, "dimension": "COLUMNS",
                           "startIndex": 0, "endIndex": 16}
        }
    })

    # Bold title row 1
    requests_body.append({
        "repeatCell": {
            "range": {"sheetId": sheet_id, "startRowIndex": 0, "endRowIndex": 1,
                      "startColumnIndex": 0, "endColumnIndex": 1},
            "cell": {"userEnteredFormat": {"textFormat": {"bold": True, "fontSize": 14}}},
            "fields": "userEnteredFormat.textFormat"
        }
    })

    # ── Summary section ───────────────────────────────────────────────────────
    summary_start_row = 4 + len(rows) + 2  # 2 blank rows after data
    data_start = 5   # 1-indexed row in sheet
    data_end   = 4 + len(rows)

    summary_items = [
        ("Total Trips",               f"=COUNTA(B{data_start}:B{data_end})"),
        ("Total Distance (km)",       f"=SUM(K{data_start}:K{data_end})"),
        ("Average Distance (km)",     f"=AVERAGE(K{data_start}:K{data_end})"),
        ("Total Duration (min)",      f"=SUM(J{data_start}:J{data_end})"),
        ("Average Duration (min)",    f"=AVERAGE(J{data_start}:J{data_end})"),
        ("Total Fares (TSh)",         f"=SUM(L{data_start}:L{data_end})"),
        ("Total Discounts (TSh)",     f"=SUM(M{data_start}:M{data_end})"),
        ("Total Booking Fees (TSh)",  f"=SUM(N{data_start}:N{data_end})"),
        ("Total Charged (TSh)",       f"=SUM(O{data_start}:O{data_end})"),
        ("Average Trip Cost (TSh)",   f"=AVERAGE(O{data_start}:O{data_end})"),
        ("Most Expensive Trip (TSh)", f"=MAX(O{data_start}:O{data_end})"),
        ("Cheapest Trip (TSh)",       f"=MIN(O{data_start}:O{data_end})"),
    ]

    summary_values = [[label, "", formula] for label, formula in summary_items]
    ws.update(
        summary_values,
        f"A{summary_start_row + 1}:C{summary_start_row + len(summary_items)}",
        value_input_option='USER_ENTERED'
    )

    # Format summary section
    requests_body.append({
        "repeatCell": {
            "range": {
                "sheetId": sheet_id,
                "startRowIndex": summary_start_row,
                "endRowIndex": summary_start_row + len(summary_items),
                "startColumnIndex": 0, "endColumnIndex": 3
            },
            "cell": {
                "userEnteredFormat": {
                    "backgroundColor": {"red": 0.91, "green": 0.96, "blue": 0.91},
                    "textFormat": {
                        "bold": True, "fontFamily": "Arial", "fontSize": 10,
                        "foregroundColor": {"red": 0.18, "green": 0.49, "blue": 0.20}
                    },
                }
            },
            "fields": "userEnteredFormat(backgroundColor,textFormat)"
        }
    })

    # Apply all formatting requests
    sh.batch_update({"requests": requests_body})

    print(f"✅ Sheet updated: {len(rows)} trips across {len(all_data)} rows.")
    print(f"🔗 URL: https://docs.google.com/spreadsheets/d/{sh.id}")
    return sh.id

# ── Main ──────────────────────────────────────────────────────────────────────
def main():
    print("Bolt Trips Tracker → Google Sheets Sync")
    print("=" * 45)

    print("Authenticating with Google…")
    creds = get_credentials()
    gc = gspread.authorize(creds)

    print(f"Reading {EXCEL_FILE}…")
    headers, rows = read_excel()
    print(f"Found {len(rows)} trips.")

    print("Opening/creating Google Sheet…")
    sh = get_or_create_sheet(gc)

    print("Uploading data…")
    sheet_id = upload_to_sheets(sh, headers, rows)

    print(f"\nDone! View at:")
    print(f"https://docs.google.com/spreadsheets/d/{sheet_id}")

if __name__ == "__main__":
    main()
