"""
IRIS — Push RFI Status + RE Remarks from Status-Register PDFs
==============================================================
Reads the two "RFI with status" PDF registers, extracts RFI No. / Status /
RE Remarks / Remark Date for every row, then updates matching rows in both
Google Sheets spreadsheets (65.5 KM and 163.22 KM).

Usage:
    python C:\\AI-Projects\\push_rfi_status_from_pdfs.py

PDF paths (edit if different):
    STATUS_PDFS list below — can be absolute Windows paths.

Rules applied when pushing:
  • OPENED → written only if the cell is currently blank
  • Any other status → always written (overrides OPENED or blank)
  • RE Remarks → written if the cell is blank  OR  the remark is non-empty
  • Remark Date → same rule as RE Remarks
  • Never deletes or modifies any other column

Run:
    python C:\\AI-Projects\\push_rfi_status_from_pdfs.py
"""

import os, re, io, sys, json, time
from pathlib import Path

# ─── UTF-8 safety ────────────────────────────────────────────────────────────
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

# ─── PDF paths — update if you move the files ─────────────────────────────────
STATUS_PDFS = [
    r"C:\AI-Projects\163.22 KM UNGUJA RFI with status.pdf",
    r"C:\AI-Projects\65.5 KM UNGUJA RFI with status.pdf",
]

# ─── Google config (same as update_rfi_google_sheets.py) ─────────────────────
CREDENTIALS_FILE = r"C:\AI-Projects\google_credentials.json"
TOKEN_FILE       = r"C:\AI-Projects\google_token.json"
SPREADSHEET_655  = "65.5 IRIS RFI Detailed Log"
SPREADSHEET_163  = "163.22 IRIS RFI Detailed Log"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

# ─── Auto-install missing packages ────────────────────────────────────────────
def _ensure(pkg, import_name=None):
    import importlib, subprocess
    try:
        importlib.import_module(import_name or pkg)
    except ImportError:
        subprocess.check_call([sys.executable, "-m", "pip", "install", pkg, "-q"])

for _pkg, _imp in [
    ("pdfplumber",              "pdfplumber"),
    ("google-auth",             "google.auth"),
    ("google-auth-oauthlib",    "google_auth_oauthlib"),
    ("google-api-python-client","googleapiclient"),
]:
    _ensure(_pkg, _imp)

import pdfplumber
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build

from project_sheets import RFI_ID_BY_NAME

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 1 — PDF PARSING
# ══════════════════════════════════════════════════════════════════════════════

# Tight RFI pattern: 4-digit serial, optional single-digit sub-number
_RFI_PAT = re.compile(
    r"((?:IRIS|IR)[-\s]?(?:QAQC|Q)[-\s]?RFI[-\s]?[UP]?[-\s]?\d{2}[-\s]?\d{4}"
    r"(?:[-\s]+\d(?!\d))?)",
    re.IGNORECASE,
)

_VALID_STATUSES = {
    "approved":               "APPROVED",
    "approved with comments": "APPROVED WITH COMMENTS",
    "not approved":           "NOT APPROVED",
    "cancelled":              "CANCELLED",
    "canceled":               "CANCELLED",
    "opened":                 "OPENED",
    "closed":                 "CLOSED",
}

def _norm_status(raw: str) -> str:
    if not raw:
        return ""
    u = raw.strip().lower()
    if u in _VALID_STATUSES:
        return _VALID_STATUSES[u]
    if "not approved" in u:    return "NOT APPROVED"
    if "with comment" in u:    return "APPROVED WITH COMMENTS"
    if "approved" in u:        return "APPROVED"
    if "cancel" in u:          return "CANCELLED"
    if "close" in u:           return "CLOSED"
    if "open" in u:            return "OPENED"
    return ""

def _norm_rfi(s: str) -> str:
    s = s.strip().upper()
    s = re.sub(r"\s+", "-", s)
    s = re.sub(r"-{2,}", "-", s)
    return s

def _split_date_from_remarks(remarks: str):
    """If a date is embedded at the end of remarks, split it out."""
    m = re.search(r"(\d{2}-\w{3}-\d{2,4}|\d{1,2}/\d{1,2}/\d{2,4})$",
                  remarks.strip())
    if m:
        return m.group(1), remarks[:m.start()].strip()
    return "", remarks


def extract_status_from_pdfs(pdf_paths: list) -> dict:
    """
    Returns dict: rfi_no_upper → {status, re_remarks, remark_date}
    """
    records: dict[str, dict] = {}
    for pdf_path in pdf_paths:
        if not os.path.exists(pdf_path):
            print(f"  ⚠  PDF not found — skipping: {pdf_path}")
            continue
        print(f"  Parsing: {os.path.basename(pdf_path)}")
        page_count = 0
        row_count  = 0
        try:
            with pdfplumber.open(pdf_path) as pdf:
                for page in pdf.pages:
                    page_count += 1
                    for tbl in (page.extract_tables() or []):
                        for row in tbl:
                            if not row:
                                continue
                            cells = [(c.strip() if c else "") for c in row]
                            row_text = " | ".join(cells)

                            rfi_matches = _RFI_PAT.findall(row_text)
                            if not rfi_matches:
                                continue

                            # Status + remarks are typically in columns 12-16
                            check = cells[12:] if len(cells) > 12 else cells
                            status_raw = ""
                            remarks    = ""
                            remark_date = ""
                            for cell in check:
                                ns = _norm_status(cell)
                                if ns and not status_raw:
                                    status_raw = ns
                                elif status_raw and cell and not remarks:
                                    remarks = cell
                                elif status_raw and remarks and cell and not remark_date:
                                    if re.search(r"\d{1,2}[-/]\w{3}", cell):
                                        remark_date = cell

                            if not status_raw:
                                continue

                            if remarks and not remark_date:
                                remark_date, remarks = _split_date_from_remarks(remarks)

                            for rfi_raw in rfi_matches:
                                key = _norm_rfi(rfi_raw)
                                existing = records.get(key)
                                # Don't overwrite a real status with OPENED
                                if (existing and
                                        existing["status"] not in ("OPENED", "") and
                                        status_raw == "OPENED"):
                                    continue
                                records[key] = {
                                    "status":      status_raw,
                                    "re_remarks":  remarks,
                                    "remark_date": remark_date,
                                }
                                row_count += 1
        except Exception as exc:
            print(f"  ❌  Error parsing {os.path.basename(pdf_path)}: {exc}")
            continue
        print(f"     → {page_count} pages, {len(records)} unique RFI entries so far")

    return records


# ══════════════════════════════════════════════════════════════════════════════
# SECTION 2 — GOOGLE SHEETS AUTH
# ══════════════════════════════════════════════════════════════════════════════

def get_google_creds():
    creds = None
    if os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
            except Exception:
                creds = None
        if not creds:
            if not os.path.exists(CREDENTIALS_FILE):
                print(f"❌  Missing credentials file: {CREDENTIALS_FILE}")
                sys.exit(1)
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(TOKEN_FILE, "w") as f:
            f.write(creds.to_json())
    return creds


def _api_retry(fn, **kwargs):
    for attempt in range(4):
        try:
            return fn(**kwargs).execute()
        except Exception as e:
            err = str(e).lower()
            if any(t in err for t in ("rate", "quota", "503", "timeout", "10053", "10054",
                                       "connection", "reset", "aborted")):
                wait = 4 ** attempt
                print(f"    ⏳  API retry {attempt+1}/4 in {wait}s ({e})")
                time.sleep(wait)
            else:
                raise
    raise RuntimeError("API call failed after 4 retries")


def _find_spreadsheet_id(drive, name: str) -> str | None:
    if name in RFI_ID_BY_NAME:
        return RFI_ID_BY_NAME[name]

    resp = _api_retry(
        drive.files().list,
        q=f"name='{name}' and mimeType='application/vnd.google-apps.spreadsheet'",
        spaces="drive", fields="files(id,name)",
    )
    files = resp.get("files", [])
    return files[0]["id"] if files else None


# ══════════════════════════════════════════════════════════════════════════════
# SECTION 3 — PUSH STATUS TO SHEETS
# ══════════════════════════════════════════════════════════════════════════════

# Column header names for status/remarks/date (must match exactly what's in
# the sheets — case-insensitive comparison is used)
_STATUS_COL_NAMES   = {"status"}
_REMARKS_COL_NAMES  = {"re's remarks", "re remarks", "re's remarks", "re remarks"}
_DATE_COL_NAMES     = {"remark date", "re's remark date", "remark_date", "date"}
_RFINO_COL_NAMES    = {"rfi no.", "rfi no", "rfi number", "rfino"}


def _col_letter(col_0based: int) -> str:
    result = ""
    n = col_0based + 1
    while n:
        n, r = divmod(n - 1, 26)
        result = chr(65 + r) + result
    return result


def push_statuses_to_spreadsheet(service, drive, ss_name: str,
                                  status_lookup: dict) -> int:
    """
    Find every named table in `ss_name`, scan its rows, and update Status /
    RE Remarks / Remark Date wherever the RFI No. matches an entry in
    `status_lookup`.
    Returns total number of cells updated.
    """
    ssid = _find_spreadsheet_id(drive, ss_name)
    if not ssid:
        print(f"  ❌  Spreadsheet not found: {ss_name}")
        return 0

    # Get full spreadsheet metadata (all sheets)
    meta = _api_retry(service.spreadsheets().get,
                      spreadsheetId=ssid, fields="sheets")
    total_updated = 0

    for sheet in meta.get("sheets", []):
        sheet_name = sheet["properties"]["title"]
        # Only process RFI_DATA tables (not Progress, Road Code, etc.)
        if "RFI_DATA" not in sheet_name.upper():
            continue

        print(f"  Sheet: {sheet_name}")

        # Read all data
        resp = _api_retry(
            service.spreadsheets().values().get,
            spreadsheetId=ssid,
            range=f"'{sheet_name}'",
            valueRenderOption="FORMATTED_VALUE",
        )
        rows = resp.get("values", [])
        if not rows:
            print(f"    (empty)")
            continue

        # Find header row — first row that has "RFI No." or "Status" etc.
        header_row_idx = None
        header_map: dict[str, int] = {}
        for i, row in enumerate(rows):
            row_lower = [str(c).strip().lower() for c in row]
            if any(n in row_lower for n in _RFINO_COL_NAMES):
                header_row_idx = i
                for j, cell in enumerate(row_lower):
                    if cell in _RFINO_COL_NAMES:   header_map["rfi_no"]   = j
                    if cell in _STATUS_COL_NAMES:   header_map["status"]   = j
                    if cell in _REMARKS_COL_NAMES:  header_map["remarks"]  = j
                    if cell in _DATE_COL_NAMES:     header_map["rem_date"] = j
                break

        if header_row_idx is None:
            print(f"    ⚠  No header row found — skipping")
            continue

        rfi_col = header_map.get("rfi_no")
        stat_col = header_map.get("status")
        rem_col  = header_map.get("remarks")
        date_col = header_map.get("rem_date")

        if rfi_col is None or stat_col is None:
            print(f"    ⚠  Missing RFI No. or Status column — skipping")
            continue

        # Build batch update requests
        data_requests = []

        for row_idx in range(header_row_idx + 1, len(rows)):
            row = rows[row_idx]
            if rfi_col >= len(row):
                continue
            rfi_val = str(row[rfi_col]).strip().upper()
            # Normalise: collapse spaces to hyphens
            rfi_key = re.sub(r"\s+", "-", rfi_val)
            rfi_key = re.sub(r"-{2,}", "-", rfi_key)

            entry = status_lookup.get(rfi_key)
            if not entry:
                # Try fuzzy: strip sub-number suffix and try again
                base = re.sub(r"-\d+$", "", rfi_key)
                entry = status_lookup.get(base)
            if not entry:
                continue

            new_status   = entry.get("status", "")
            new_remarks  = entry.get("re_remarks", "")
            new_date     = entry.get("remark_date", "")

            # Sheets row number (1-based)
            sheets_row = header_row_idx + 1 + (row_idx - header_row_idx)

            # Existing values
            cur_status  = str(row[stat_col]).strip()  if stat_col < len(row)  else ""
            cur_remarks = str(row[rem_col]).strip()   if rem_col  is not None and rem_col < len(row) else ""
            cur_date    = str(row[date_col]).strip()  if date_col is not None and date_col < len(row) else ""

            # Status: write if blank OR upgrading from OPENED to something real
            write_status = (
                new_status and
                (not cur_status or
                 cur_status.upper() == "OPENED" and new_status != "OPENED")
            )
            # Remarks/date: write if cell is blank or we have richer content
            write_remarks = new_remarks and (not cur_remarks or len(new_remarks) > len(cur_remarks))
            write_date    = new_date    and not cur_date

            if not (write_status or write_remarks or write_date):
                continue

            # Build individual cell update requests
            if write_status and stat_col is not None:
                data_requests.append({
                    "updateCells": {
                        "range": {
                            "sheetId": sheet["properties"]["sheetId"],
                            "startRowIndex": sheets_row - 1,
                            "endRowIndex":   sheets_row,
                            "startColumnIndex": stat_col,
                            "endColumnIndex":   stat_col + 1,
                        },
                        "rows": [{"values": [{"userEnteredValue": {"stringValue": new_status}}]}],
                        "fields": "userEnteredValue",
                    }
                })
            if write_remarks and rem_col is not None:
                data_requests.append({
                    "updateCells": {
                        "range": {
                            "sheetId": sheet["properties"]["sheetId"],
                            "startRowIndex": sheets_row - 1,
                            "endRowIndex":   sheets_row,
                            "startColumnIndex": rem_col,
                            "endColumnIndex":   rem_col + 1,
                        },
                        "rows": [{"values": [{"userEnteredValue": {"stringValue": new_remarks}}]}],
                        "fields": "userEnteredValue",
                    }
                })
            if write_date and date_col is not None:
                data_requests.append({
                    "updateCells": {
                        "range": {
                            "sheetId": sheet["properties"]["sheetId"],
                            "startRowIndex": sheets_row - 1,
                            "endRowIndex":   sheets_row,
                            "startColumnIndex": date_col,
                            "endColumnIndex":   date_col + 1,
                        },
                        "rows": [{"values": [{"userEnteredValue": {"stringValue": new_date}}]}],
                        "fields": "userEnteredValue",
                    }
                })

        if not data_requests:
            print(f"    → nothing to update")
            continue

        # Send in batches of 500
        BATCH = 500
        for i in range(0, len(data_requests), BATCH):
            chunk = data_requests[i: i + BATCH]
            _api_retry(service.spreadsheets().batchUpdate,
                       spreadsheetId=ssid, body={"requests": chunk})
        cells_written = len(data_requests)
        total_updated += cells_written
        print(f"    ✓  {cells_written} cell(s) updated")

    return total_updated


# ══════════════════════════════════════════════════════════════════════════════
# SECTION 4 — ATTACHMENT DEFICIT CHECK
# ══════════════════════════════════════════════════════════════════════════════

RFI_ZIPS_FOLDER = r"D:\HIGHWAY ENGINEERING CAREER\RFI-ZIPs"
_VALID_RFI_RE   = re.compile(
    r"^(?:IRIS|IR)[-\s]?(?:QAQC|Q)[-\s]?RFI[-\s]?[UP][-\s]?\d{2}[-\s]?\d{4}"
    r"(?:[-\s]?\d+)?(?:\.[a-zA-Z]{2,5})?$",
    re.IGNORECASE,
)


def check_attachment_deficits(service, drive, status_lookup: dict) -> None:
    """
    For each spreadsheet / table, compare the RFI No. column against files
    in RFI_ZIPS_FOLDER (ZIPs + RAR contents index).  Print any RFIs that
    have no matching attachment file.
    """
    print("\n" + "=" * 60)
    print("ATTACHMENT DEFICIT CHECK  (SITE & LAB — both projects)")
    print("=" * 60)

    # Build set of available filenames (ZIPs + RAR contents index)
    available: set[str] = set()
    if os.path.isdir(RFI_ZIPS_FOLDER):
        for f in os.listdir(RFI_ZIPS_FOLDER):
            available.add(f.upper())
            # Also check inside ZIPs
            fp = os.path.join(RFI_ZIPS_FOLDER, f)
            if f.lower().endswith(".zip"):
                try:
                    import zipfile
                    with zipfile.ZipFile(fp, "r") as z:
                        for member in z.namelist():
                            available.add(os.path.basename(member).upper())
                except Exception:
                    pass
        # RAR contents index
        rar_log = os.path.join(RFI_ZIPS_FOLDER, "_rar_contents_log.json")
        if os.path.exists(rar_log):
            try:
                with open(rar_log, "r", encoding="utf-8") as f:
                    rar_index = json.load(f)
                for members in rar_index.values():
                    for m in members:
                        available.add(os.path.basename(m).upper())
            except Exception:
                pass

    def _has_attachment(rfi_no: str) -> bool:
        """Return True if any file in available set starts with the RFI no."""
        key = rfi_no.upper().replace(" ", "-")
        return any(a.startswith(key) for a in available)

    deficit_total = 0
    for ss_name in [SPREADSHEET_655, SPREADSHEET_163]:
        ssid = _find_spreadsheet_id(drive, ss_name)
        if not ssid:
            continue
        meta = _api_retry(service.spreadsheets().get,
                          spreadsheetId=ssid, fields="sheets")
        print(f"\n  {ss_name}")
        for sheet in meta.get("sheets", []):
            sheet_name = sheet["properties"]["title"]
            if "RFI_DATA" not in sheet_name.upper():
                continue
            resp = _api_retry(
                service.spreadsheets().values().get,
                spreadsheetId=ssid,
                range=f"'{sheet_name}'",
                valueRenderOption="FORMATTED_VALUE",
            )
            rows = resp.get("values", [])
            # Find header
            rfi_col = None
            data_start = 0
            for i, row in enumerate(rows):
                row_lower = [str(c).strip().lower() for c in row]
                if any(n in row_lower for n in _RFINO_COL_NAMES):
                    rfi_col = next(j for j, c in enumerate(row_lower) if c in _RFINO_COL_NAMES)
                    data_start = i + 1
                    break
            if rfi_col is None:
                continue

            deficits = []
            for row in rows[data_start:]:
                if rfi_col >= len(row):
                    continue
                rfi_no = str(row[rfi_col]).strip()
                if not rfi_no or rfi_no.lower() in ("rfi no.", "rfi no", "total"):
                    continue
                if not _has_attachment(rfi_no):
                    deficits.append(rfi_no)

            if deficits:
                print(f"    ⚠  {sheet_name}: {len(deficits)} RFI(s) WITHOUT attachment:")
                for d in deficits[:20]:
                    print(f"       • {d}")
                if len(deficits) > 20:
                    print(f"       … and {len(deficits)-20} more")
                deficit_total += len(deficits)
            else:
                print(f"    ✓  {sheet_name}: all attachments present")

    if deficit_total == 0:
        print("\n  ✅  All RFIs across both projects have attachment files.")
    else:
        print(f"\n  ⚠  Total deficits: {deficit_total} RFI(s) missing attachment(s).")


# ══════════════════════════════════════════════════════════════════════════════
# MAIN
# ══════════════════════════════════════════════════════════════════════════════

def main():
    SEP  = "=" * 65
    SEP2 = "-" * 65

    print()
    print(SEP)
    print("  IRIS — Push RFI Status + RE Remarks from PDF Registers")
    print(SEP)

    # ── Step 1: Parse PDFs ────────────────────────────────────────────────────
    print()
    print(SEP2)
    print("STEP 1 — Parse Status PDFs")
    print(SEP2)
    status_lookup = extract_status_from_pdfs(STATUS_PDFS)
    if not status_lookup:
        print("❌  No status records extracted from PDFs. Check PDF paths above.")
        sys.exit(1)

    # Summary
    counts: dict[str, int] = {}
    for v in status_lookup.values():
        s = v["status"]
        counts[s] = counts.get(s, 0) + 1
    print(f"\n  Extracted {len(status_lookup)} unique RFI entries:")
    for s, n in sorted(counts.items(), key=lambda x: -x[1]):
        print(f"    {s:<28} {n:>4}")

    # ── Step 2: Google Auth ───────────────────────────────────────────────────
    print()
    print(SEP2)
    print("STEP 2 — Authenticate with Google")
    print(SEP2)
    creds   = get_google_creds()
    service = build("sheets", "v4", credentials=creds)
    drive   = build("drive",  "v3", credentials=creds)
    print("  ✓  Authenticated")

    # ── Step 3: Push to both spreadsheets ────────────────────────────────────
    print()
    print(SEP2)
    print("STEP 3 — Push statuses to Google Sheets")
    print(SEP2)
    grand_total = 0
    for ss_name in [SPREADSHEET_655, SPREADSHEET_163]:
        print(f"\n  Spreadsheet: {ss_name}")
        n = push_statuses_to_spreadsheet(service, drive, ss_name, status_lookup)
        grand_total += n

    print()
    print(SEP2)
    print(f"STEP 3 DONE — {grand_total} cell(s) updated across both spreadsheets")

    # ── Step 4: Attachment deficit check ─────────────────────────────────────
    print()
    print(SEP2)
    print("STEP 4 — Attachment deficit check")
    print(SEP2)
    check_attachment_deficits(service, drive, status_lookup)

    print()
    print(SEP)
    print("  ✅  push_rfi_status_from_pdfs.py  COMPLETE")
    print(SEP)
    print()


if __name__ == "__main__":
    main()
