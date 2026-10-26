"""
IRIS RFI PDF → Google Sheets Importer  (v1 — April 2026)
=========================================================
Reads "65.5 KM UNGUJA RFI with status.pdf" and
       "163.22 KM UNGUJA RFI with status.pdf"
from the RFI-ZIPs folder, extracts every RFI record, then inserts any
records that are NOT already present in the matching Google Sheets
named table — in chronological order.

Rules:
  • NEW RFIs  (not found in the sheet)   → inserted with Status = "OPENED"
  • EXISTING  (RFI No. already in sheet) → never overwritten / never touched
  • Inserts happen ABOVE the footer row, just like update_rfi_google_sheets.py
  • S/No. continues from the highest existing number in the table

Target tables:
  "65.5 IRIS RFI Detailed Log"
      KM 65.5 UNGUJA LAB RFI_DATA   ← QA/QC section of the 65.5 KM PDF
      KM 65.5 UNGUJA SITE RFI_DATA  ← Q RFI  section of the 65.5 KM PDF

  "163.22 IRIS RFI Detailed Log"
      KM 163.22 UNGUJA LAB RFI_DATA   ← QA/QC section of the 163.22 KM PDF
      KM 163.22 UNGUJA SITE RFI_DATA  ← Q RFI  section of the 163.22 KM PDF

Run:
  python C:\\AI-Projects\\import_rfi_from_pdf.py
"""

import os, re, sys, time, json, datetime

# ── Auto-install packages ──────────────────────────────────────────────────────
def _ensure(pkg, import_name=None):
    import importlib, subprocess
    try:
        importlib.import_module(import_name or pkg)
    except ImportError:
        print(f"Installing {pkg} …")
        subprocess.check_call([sys.executable, "-m", "pip", "install", pkg, "-q"])

for _p, _i in [
    ("pdfplumber",               "pdfplumber"),
    ("gspread",                  "gspread"),
    ("google-auth",              "google.oauth2"),
    ("google-auth-oauthlib",     "google_auth_oauthlib"),
    ("google-api-python-client", "googleapiclient"),
]:
    _ensure(_p, _i)

import pdfplumber
import gspread
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# ── UTF-8 safety ───────────────────────────────────────────────────────────────
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

# ══════════════════════════════════════════════════════════════════════════════
# CONFIGURATION
# ══════════════════════════════════════════════════════════════════════════════
CREDENTIALS_FILE = r"C:\AI-Projects\google_credentials.json"
TOKEN_FILE       = r"C:\AI-Projects\google_token.json"
ZIP_FOLDER       = r"D:\HIGHWAY ENGINEERING CAREER\RFI-ZIPs"

# Source PDFs (must be present in ZIP_FOLDER)
PDF_655  = os.path.join(ZIP_FOLDER, "65.5 KM UNGUJA RFI with status.pdf")
PDF_163  = os.path.join(ZIP_FOLDER, "163.22 KM UNGUJA RFI with status.pdf")

# Target spreadsheet names (must match exactly what's in Google Drive)
from project_sheets import RFI_ID

SPREADSHEET_655  = "65.5 IRIS RFI Detailed Log"
SPREADSHEET_163  = "163.22 IRIS RFI Detailed Log"

# Google Sheets API scopes
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

# The 20 DATA_COLUMNS — same layout as update_rfi_google_sheets.py
DATA_COLUMNS = [
    "S/No.",               # 0
    "RFI No.",             # 1
    "Date Received",       # 2
    "Inspection Date",     # 3
    "Road Code",           # 4
    "Road Name",           # 5
    "Island",              # 6
    "Contract Number",     # 7
    "Start Station",       # 8
    "End Station",         # 9
    "Length Executed (km)",# 10
    "Merged Length (km)",  # 11
    "Activity Name",       # 12
    "Work Description",    # 13
    "Station Range",       # 14
    "Status",              # 15
    "RE's Remarks",        # 16
    "Remark Date",         # 17
    "Attachment Path",     # 18
    "Project",             # 19
]
_CI = {col: i for i, col in enumerate(DATA_COLUMNS)}
MAX_COLS = len(DATA_COLUMNS)

APPROVED_STATUSES = {"APPROVED", "APPROVED WITH COMMENTS", "NOT APPROVED",
                     "CANCELLED", "OPENED", "CLOSED"}

# ── Retry helper ───────────────────────────────────────────────────────────────
def _api(fn, *args, retries=5, **kwargs):
    for attempt in range(1, retries + 1):
        try:
            return fn(*args, **kwargs)
        except gspread.exceptions.APIError as e:
            if attempt == retries:
                raise
            wait = min(30, 5 * attempt)
            print(f"    ⏳ API error (attempt {attempt}/{retries}): {e} — retrying in {wait}s …")
            time.sleep(wait)
    return fn(*args, **kwargs)


# ══════════════════════════════════════════════════════════════════════════════
# PDF PARSING
# ══════════════════════════════════════════════════════════════════════════════

# Column headers as they appear in the PDF tables
_PDF_COL_ALIASES = {
    "s/no":              "sno",
    "s/no.":             "sno",
    "rfi no.":           "rfi_no",
    "rfi no":            "rfi_no",
    "date received":     "date_received",
    "inspection date":   "inspection_date",
    "inspection\ndate":  "inspection_date",
    "location":          "road_code",
    "road name":         "road_name",
    "island":            "island",
    "contract number":   "contract_number",
    "contract\nnumber":  "contract_number",
    "start station":     "start_station",
    "start\nstation":    "start_station",
    "end station":       "end_station",
    "end\nstation":      "end_station",
    "length executed\n(km)": "length_km",
    "length executed (km)":  "length_km",
    "length\nexecuted\n(km)":"length_km",
    "work description":  "work_description",
    "station range":     "station_range",
    "status":            "status",
    "re's remarks":      "re_remarks",
    "re's\nremarks":     "re_remarks",
    "remark date":       "remark_date",
    "remark\ndate":      "remark_date",
    "attachment path":   "attachment_path",
}


def _norm_header(raw: str) -> str:
    """Normalise a PDF column header to a canonical key."""
    if raw is None:
        return ""
    return _PDF_COL_ALIASES.get(raw.strip().lower(), raw.strip().lower())


def _clean_cell(v) -> str:
    """Return a clean string from a cell value (handles None)."""
    if v is None:
        return ""
    return str(v).strip()


def _parse_station(s: str) -> str:
    """Normalise chainage like '0+000' or '0+1034' → keep as-is string."""
    s = _clean_cell(s)
    # Remove any newlines
    s = s.replace("\n", " ").strip()
    return s


def _looks_like_rfi_no(v: str) -> bool:
    """Return True if the cell looks like an RFI number."""
    v = v.upper()
    return bool(re.match(r"^IR(?:IS)?[-_]", v) or re.search(r"RFI[-_]", v))


def extract_rfi_records_from_pdf(pdf_path: str, km_key: str) -> dict[str, list[dict]]:
    """
    Parse a PDF that contains two sections: QA/QC RFI's and Q RFI's.

    Returns:
      {
        "lab":  [ {...}, {...}, … ],   # QA/QC records → LAB table
        "site": [ {...}, {...}, … ],   # Q RFI records  → SITE table
      }

    Each record dict has keys matching DATA_COLUMNS.
    """
    results = {"lab": [], "site": []}

    if not os.path.exists(pdf_path):
        print(f"  ⚠  PDF not found: {pdf_path}")
        return results

    print(f"\n  Parsing: {os.path.basename(pdf_path)}")

    section   = None   # "lab" or "site"
    col_map   = {}     # {column_index: canonical_key}
    header_found = False

    with pdfplumber.open(pdf_path) as pdf:
        for page_no, page in enumerate(pdf.pages, 1):
            # ── Extract tables from this page ──────────────────────────────
            tables = page.extract_tables({
                "vertical_strategy":   "lines",
                "horizontal_strategy": "lines",
                "snap_tolerance":       6,
                "intersection_tolerance": 6,
            })

            if not tables:
                # Fallback: try text-based table extraction
                tables = page.extract_tables()

            for table in tables:
                if not table:
                    continue
                for row_idx, row in enumerate(table):
                    if not row or all(c is None or _clean_cell(c) == "" for c in row):
                        continue

                    # Convert None cells to empty strings
                    cells = [_clean_cell(c) for c in row]
                    first = cells[0].upper() if cells[0] else ""
                    joined = " ".join(cells[:4]).upper()

                    # ── Section header detection ───────────────────────
                    if "QA/QC RFI" in joined or "QA/QC" in first:
                        section = "lab"
                        header_found = False
                        col_map = {}
                        continue

                    if (re.search(r"\bQ\s+RFI", joined) or
                            (first == "Q" and "RFI" in joined) or
                            "Q RFI'S" in joined):
                        section = "site"
                        header_found = False
                        col_map = {}
                        continue

                    if section is None:
                        continue

                    # ── Header row detection (contains "RFI No." or "S/No.") ──
                    if not header_found:
                        lower_joined = " ".join(cells).lower()
                        if "rfi no" in lower_joined or "s/no" in lower_joined:
                            col_map = {}
                            for ci, c in enumerate(cells):
                                key = _norm_header(c)
                                if key:
                                    col_map[ci] = key
                            header_found = True
                        continue

                    # ── Data row ───────────────────────────────────────────
                    # Skip aggregate / stat rows
                    if re.match(r"^(opened|approved|total|sum|count)", first, re.I):
                        continue

                    # Must contain something that looks like an RFI number
                    rfi_no = ""
                    for ci, key in col_map.items():
                        if key == "rfi_no" and ci < len(cells):
                            rfi_no = cells[ci].strip()
                            break
                    # Also try to find RFI No. in any cell if col_map missed it
                    if not rfi_no:
                        for c in cells:
                            if _looks_like_rfi_no(c):
                                rfi_no = c.strip()
                                break

                    if not rfi_no or not _looks_like_rfi_no(rfi_no):
                        continue

                    # Build record dict
                    rec: dict[str, str] = {}
                    for ci, key in col_map.items():
                        if ci < len(cells):
                            rec[key] = cells[ci]

                    rec["rfi_no"] = rfi_no  # always set from lookup above

                    # Normalise status
                    raw_status = rec.get("status", "").strip().upper()
                    # PDF may have "Not Approved" → "NOT APPROVED"
                    if raw_status in APPROVED_STATUSES:
                        rec["status"] = raw_status.title() if raw_status not in (
                            "APPROVED", "OPENED", "CLOSED") else raw_status.capitalize()
                        # Map to exact Sheets dropdown values
                        _status_map = {
                            "Approved": "APPROVED",
                            "Approved With Comments": "APPROVED WITH COMMENTS",
                            "Not Approved": "NOT APPROVED",
                            "Cancelled": "CANCELLED",
                            "Opened": "OPENED",
                            "Closed": "CLOSED",
                        }
                        for k, v in _status_map.items():
                            if raw_status == v:
                                rec["status"] = v
                                break
                    else:
                        rec["status"] = "OPENED"  # default for unrecognised

                    rec["km_key"]  = km_key
                    results[section].append(rec)

    # De-duplicate by RFI No. (keep first occurrence)
    for section_key in ("lab", "site"):
        seen = set()
        deduped = []
        for r in results[section_key]:
            rn = r.get("rfi_no", "").upper()
            if rn and rn not in seen:
                seen.add(rn)
                deduped.append(r)
        results[section_key] = deduped
        print(f"    [{section_key.upper():4s}] {len(deduped):3d} record(s) parsed")

    return results


# ══════════════════════════════════════════════════════════════════════════════
# ACTIVITY NAME INFERENCE  (reuses the same patterns as update_rfi_google_sheets)
# ══════════════════════════════════════════════════════════════════════════════
_ACT_PATTERNS = [
    (r"site\s*clear|clearing\s*site|vegetation\s*clear|bush\s*clear|strip\s*topsoil|demolish",
     "SITE CLEARANCE"),
    (r"roadbed|road\s*bed|formation\s*level|shaping\s*road|profile\s*grading",
     "ROADBED PREPARATION"),
    (r"sub[\s\-]?base",                   "SUBBASE LAYING"),
    (r"base[\s\-]?course|granular\s*base","BASECOURSE"),
    (r"prime\s*coat|priming",             "PRIME COAT"),
    (r"tack\s*coat|bond\s*coat",          "TACK COAT"),
    (r"asphalt|wearing\s*course|binder\s*course|hot\s*mix|surface\s*dressing",
     "ASPHALT LAYING"),
    (r"culvert|pipe\s*drain|headwall|wingwall",
     "CULVERT"),
    (r"sub[\s\-]?grade|subgrade|compaction\s*test|DCP\s*test",
     "SUBGRADE LAYING"),
    (r"road\s*mark|line\s*mark|thermoplastic",
     "ROAD MARKING"),
    (r"guard[\s\-]?rail",                 "GUARD RAIL REPARATION"),
    (r"DCP\s*test|material\s*sampl|aggregate|bitumen\s*test",
     "MATERIAL SAMPLING & TESTING"),
    (r"benchmark|level(?:l?ing)\s*(?:for\s+)?surv|survey\s*insp|setting\s*out",
     "SURVEY INSPECTION"),
    (r"ogsl|ogl|centre\s*line\s*survey",  "OGL SURVEY"),
    (r"excavat|cut\s*and\s*fill|earth",   "EARTHWORKS"),
]

def infer_activity(work_desc: str) -> str:
    """Infer IRIS Activity Name from a Work Description string."""
    if not work_desc:
        return ""
    wd = work_desc.lower()
    for pat, act in _ACT_PATTERNS:
        if re.search(pat, wd, re.I):
            return act
    return ""


# ══════════════════════════════════════════════════════════════════════════════
# GOOGLE SHEETS AUTH
# ══════════════════════════════════════════════════════════════════════════════
def get_gspread_client() -> gspread.Client:
    creds = None
    if os.path.exists(TOKEN_FILE):
        try:
            creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
        except Exception:
            creds = None

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
            except Exception:
                creds = None
        if not creds:
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(TOKEN_FILE, "w") as tok:
            tok.write(creds.to_json())

    return gspread.authorize(creds)


# ══════════════════════════════════════════════════════════════════════════════
# NAMED TABLE HELPERS  (re-uses the same logic as update_rfi_google_sheets.py)
# ══════════════════════════════════════════════════════════════════════════════
def build_table_name(km_key: str, island: str, rfi_type: str) -> str:
    return f"KM {km_key} {island.upper()} {rfi_type.upper()} RFI_DATA"


def find_named_table(ws: gspread.Worksheet, table_name: str) -> dict | None:
    """
    Find a Named Range on a worksheet that matches table_name.
    Returns a dict with keys: start_row, end_row, start_col, end_col (all 0-based).
    """
    try:
        spreadsheet = ws.spreadsheet
        named_ranges = _api(spreadsheet.fetch_sheet_metadata).get(
            "namedRanges", [])
        for nr in named_ranges:
            if nr.get("name", "").strip().upper() == table_name.upper():
                rng = nr["range"]
                return {
                    "name":      table_name,
                    "sheet_id":  rng.get("sheetId"),
                    "start_row": rng.get("startRowIndex", 0),
                    "end_row":   rng.get("endRowIndex", 0),
                    "start_col": rng.get("startColumnIndex", 0),
                    "end_col":   rng.get("endColumnIndex", MAX_COLS),
                }
    except Exception as e:
        print(f"    ⚠ Could not fetch named ranges: {e}")
    return None


def read_table_data(ws: gspread.Worksheet, tinfo: dict):
    """
    Read the rows of a named table.
    Returns (header_map, existing_rfi_nos, max_sno, data_rows)
      header_map      : {column_name: 0-based col index within the row}
      existing_rfi_nos: set of uppercase RFI No. strings
      max_sno         : highest S/No. integer in the table
      data_rows       : list of row lists (header + data; NOT footer)
    """
    sr = tinfo["start_row"] + 1   # 1-based
    er = tinfo["end_row"]         # 1-based inclusive
    sc = tinfo["start_col"] + 1   # 1-based
    ec = min(tinfo["end_col"], tinfo["start_col"] + MAX_COLS)  # 1-based exclusive

    # Build A1-notation range
    def col_letter(n):  # 0-based → "A", "B", …
        s = ""
        n += 1
        while n:
            n, r = divmod(n - 1, 26)
            s = chr(65 + r) + s
        return s

    range_a1 = (f"{col_letter(tinfo['start_col'])}{sr}"
                f":{col_letter(ec - 1)}{er}")
    try:
        rows = _api(ws.get, range_a1)
    except Exception as e:
        print(f"    ⚠  Could not read table range {range_a1}: {e}")
        return {}, set(), 0, []

    if not rows:
        return {}, set(), 0, []

    # Row 0 = header
    header = [str(h).strip() for h in rows[0]]
    header_map = {h: i for i, h in enumerate(header)}
    # Case-insensitive fallback
    _hlow = {k.lower(): v for k, v in header_map.items()}
    for col in DATA_COLUMNS:
        if col not in header_map and col.lower() in _hlow:
            header_map[col] = _hlow[col.lower()]

    rfi_col = header_map.get("RFI No.")
    sno_col = header_map.get("S/No.")

    existing_rfis = set()
    max_sno = 0

    # Skip header (rows[0]) and footer (rows[-1] if > 1 row)
    data_rows = rows[1:]
    if len(data_rows) > 1:
        last = data_rows[-1]
        # Footer heuristic: S/No. column is empty or non-integer
        if sno_col is not None and sno_col < len(last):
            try:
                int(last[sno_col])
            except (ValueError, TypeError):
                data_rows = data_rows[:-1]  # strip footer

    for row in data_rows:
        if rfi_col is not None and rfi_col < len(row):
            rn = str(row[rfi_col]).strip().upper()
            if rn:
                existing_rfis.add(rn)
        if sno_col is not None and sno_col < len(row):
            try:
                v = int(str(row[sno_col]).strip())
                max_sno = max(max_sno, v)
            except (ValueError, TypeError):
                pass

    return header_map, existing_rfis, max_sno, data_rows


# ══════════════════════════════════════════════════════════════════════════════
# ROW BUILDER
# ══════════════════════════════════════════════════════════════════════════════
def build_sheet_row(sno: int, rec: dict, km_key: str,
                    header_map: dict, tinfo: dict) -> list:
    """
    Build a 20-column row list from a parsed PDF record.
    New RFIs always get Status = "OPENED".
    """
    def col_letter(n):
        s = ""
        n += 1
        while n:
            n, r = divmod(n - 1, 26)
            s = chr(65 + r) + s
        return s

    # Start station / end station
    start_raw  = rec.get("start_station", "")
    end_raw    = rec.get("end_station",   "")

    # Station string like "0+000" → extract numeric km value
    def _km_num(s):
        m = re.match(r"(\d+)\+(\d+)", s.replace(" ", ""))
        if m:
            return int(m.group(1)) + int(m.group(2)) / 1000
        try:
            return float(s)
        except (ValueError, TypeError):
            return None

    s_km = _km_num(start_raw)
    e_km = _km_num(end_raw)

    # Length formula (if both ends known, use Sheets formula; else use PDF value)
    if s_km is not None and e_km is not None and e_km >= s_km:
        # Write as formula referencing Start/End Station columns if possible
        sc_idx = header_map.get("Start Station")
        ec_idx = header_map.get("End Station")
        if sc_idx is not None and ec_idx is not None and tinfo:
            base_col = tinfo["start_col"]
            def _col(ci): return col_letter(base_col + ci)
            # placeholder; will be set per row
            length_val = f"=({_col(ec_idx)}-{_col(sc_idx)})/1000"
        else:
            length_val = round(e_km - s_km, 3) if e_km >= s_km else ""
    else:
        length_val = rec.get("length_km", "")

    # Contract number formula
    proj_val = f"KM {km_key}"
    contract_val = rec.get("contract_number", "")
    if not contract_val:
        contract_val = ("SMD/F01/W/DP/2024-2025/27" if km_key == "65.5"
                        else "SMZ/P01/W/DP/2025-2026/30")

    # Activity name (infer if not present)
    act_val = infer_activity(rec.get("work_description", ""))

    # Station range
    sr_raw = rec.get("station_range", "").strip()
    if not sr_raw and start_raw and end_raw:
        sr_raw = f"{start_raw}-{end_raw}"

    row = [""] * MAX_COLS
    row[_CI["S/No."]]              = sno
    row[_CI["RFI No."]]            = rec.get("rfi_no", "")
    row[_CI["Date Received"]]      = rec.get("date_received", "")
    row[_CI["Inspection Date"]]    = rec.get("inspection_date", "")
    row[_CI["Road Code"]]          = rec.get("road_code", "")
    row[_CI["Road Name"]]          = rec.get("road_name", "")
    row[_CI["Island"]]             = rec.get("island", "UNGUJA")
    row[_CI["Contract Number"]]    = contract_val
    row[_CI["Start Station"]]      = start_raw
    row[_CI["End Station"]]        = end_raw
    row[_CI["Length Executed (km)"]]= length_val
    row[_CI["Merged Length (km)"]] = ""
    row[_CI["Activity Name"]]      = act_val
    row[_CI["Work Description"]]   = rec.get("work_description", "")
    row[_CI["Station Range"]]      = sr_raw
    row[_CI["Status"]]             = "OPENED"   # NEW RFIs always get OPENED
    row[_CI["RE's Remarks"]]       = rec.get("re_remarks", "")
    row[_CI["Remark Date"]]        = rec.get("remark_date", "")
    row[_CI["Attachment Path"]]    = rec.get("attachment_path", "")
    row[_CI["Project"]]            = proj_val

    return row


# ══════════════════════════════════════════════════════════════════════════════
# INSERT MISSING RFIs
# ══════════════════════════════════════════════════════════════════════════════
def insert_missing_rfis(gc: gspread.Client,
                        records: list[dict],
                        spreadsheet_name: str,
                        table_name: str,
                        km_key: str) -> int:
    """
    Open spreadsheet_name, find table_name, determine which records from
    `records` are NOT already present, then insert them above the footer.

    Returns: number of rows inserted.
    """
    print(f"\n  ── {table_name} ──")

    # Open spreadsheet
    try:
        ss = _api(gc.open_by_key, RFI_ID[km_key])
    except gspread.exceptions.SpreadsheetNotFound:
        print(f"    ❌  Spreadsheet '{spreadsheet_name}' not found in Google Drive.")
        return 0

    # Find the worksheet that contains the named table
    ws_found = None
    tinfo    = None
    for ws in _api(ss.worksheets):
        tinfo = find_named_table(ws, table_name)
        if tinfo:
            ws_found = ws
            break

    if ws_found is None or tinfo is None:
        print(f"    ❌  Named table '{table_name}' not found in '{spreadsheet_name}'.")
        return 0

    print(f"    Found in sheet: '{ws_found.title}'  rows {tinfo['start_row']+1}–{tinfo['end_row']}")

    # Read existing data
    header_map, existing_rfis, max_sno, data_rows = read_table_data(ws_found, tinfo)
    print(f"    Existing RFIs in table: {len(existing_rfis):,}")
    print(f"    Max S/No. in table:     {max_sno}")

    # Filter to new-only records
    new_records = [r for r in records
                   if r.get("rfi_no", "").upper() not in existing_rfis]
    print(f"    Records from PDF:       {len(records):,}")
    print(f"    New (not in sheet):     {len(new_records):,}")

    if not new_records:
        print("    ✓  Nothing to insert.")
        return 0

    # Sort chronologically: by RFI No. (numeric portion), then by date received
    def sort_key(r):
        rn = r.get("rfi_no", "")
        # Extract the numeric suffix from the RFI No.
        m = re.search(r"(\d+)(?:[.-]\d+)?$", rn)
        num = int(m.group(1)) if m else 9999999
        dr = r.get("date_received", "")
        return (num, dr)

    new_records.sort(key=sort_key)

    # Build rows and insert above footer
    # Footer row = tinfo["end_row"] (1-based)
    # Insert point = tinfo["end_row"] - 1 (insert above footer, i.e. before last row)
    inserted = 0
    sno_counter = max_sno

    # We insert row-by-row to keep S/No. correct
    # (Batch insert is faster but requires adjusting formulas — row-by-row is safer)
    for rec in new_records:
        sno_counter += 1
        row_data = build_sheet_row(sno_counter, rec, km_key, header_map, tinfo)

        # Insert a blank row above the footer (before end_row, which is the footer)
        footer_1based = tinfo["end_row"]   # this is the footer row (1-based)
        try:
            _api(ws_found.spreadsheet.batch_update, {
                "requests": [{
                    "insertDimension": {
                        "range": {
                            "sheetId":    tinfo["sheet_id"],
                            "dimension":  "ROWS",
                            "startIndex": footer_1based - 1,  # 0-based before footer
                            "endIndex":   footer_1based,
                        },
                        "inheritFromBefore": True,
                    }
                }]
            })
            # Update tinfo to reflect new footer position
            tinfo["end_row"] += 1

            # Write the data into the newly inserted row
            def col_letter(n):
                s = ""
                n += 1
                while n:
                    n, r = divmod(n - 1, 26)
                    s = chr(65 + r) + s
                return s

            start_col_letter = col_letter(tinfo["start_col"])
            end_col_letter   = col_letter(tinfo["start_col"] + MAX_COLS - 1)
            cell_range = f"{start_col_letter}{footer_1based}:{end_col_letter}{footer_1based}"
            _api(ws_found.update, cell_range, [row_data],
                 value_input_option="USER_ENTERED")

            rfi_no = rec.get("rfi_no", "?")
            print(f"    ✓  Inserted S/No.{sno_counter:>4}: {rfi_no}")
            inserted += 1

        except Exception as e:
            print(f"    ✗  Failed to insert S/No.{sno_counter}: {rec.get('rfi_no','?')} — {e}")

        # Small delay to avoid quota
        time.sleep(0.3)

    return inserted


# ══════════════════════════════════════════════════════════════════════════════
# MAIN
# ══════════════════════════════════════════════════════════════════════════════
def main():
    SEP = "=" * 65
    print()
    print(SEP)
    print("  IRIS RFI PDF → Google Sheets Importer  (April 2026)")
    print(SEP)
    print(f"  65.5 KM PDF  : {PDF_655}")
    print(f"  163.22 KM PDF: {PDF_163}")
    print()

    # ── Step 1: Parse PDFs ─────────────────────────────────────────────────
    print("─" * 65)
    print("STEP 1 — PARSING PDFs")
    print("─" * 65)

    records_655 = extract_rfi_records_from_pdf(PDF_655,  "65.5")
    records_163 = extract_rfi_records_from_pdf(PDF_163,  "163.22")

    total_from_pdfs = (len(records_655["lab"]) + len(records_655["site"]) +
                       len(records_163["lab"]) + len(records_163["site"]))
    print(f"\n  Total records extracted: {total_from_pdfs}")
    if total_from_pdfs == 0:
        print("\n  ⚠  No records extracted from PDFs.")
        print("     Possible reasons:")
        print("     1. PDFs are not present in the RFI-ZIPs folder")
        print("     2. PDF table structure could not be parsed")
        print("     3. Try opening the PDF and checking the table formatting")
        return

    # ── Step 2: Authenticate ──────────────────────────────────────────────
    print()
    print("─" * 65)
    print("STEP 2 — AUTHENTICATING WITH GOOGLE SHEETS")
    print("─" * 65)
    try:
        gc = get_gspread_client()
        print("  ✓  Authenticated successfully")
    except Exception as e:
        print(f"  ❌  Authentication failed: {e}")
        return

    # ── Step 3: Insert missing RFIs ───────────────────────────────────────
    print()
    print("─" * 65)
    print("STEP 3 — INSERTING MISSING RFIs INTO GOOGLE SHEETS")
    print("─" * 65)

    total_inserted = 0

    # 65.5 KM — LAB (QA/QC)
    if records_655["lab"]:
        n = insert_missing_rfis(gc, records_655["lab"],
                                SPREADSHEET_655,
                                build_table_name("65.5", "UNGUJA", "LAB"),
                                "65.5")
        total_inserted += n

    # 65.5 KM — SITE (Q RFIs)
    if records_655["site"]:
        n = insert_missing_rfis(gc, records_655["site"],
                                SPREADSHEET_655,
                                build_table_name("65.5", "UNGUJA", "SITE"),
                                "65.5")
        total_inserted += n

    # 163.22 KM — LAB (QA/QC)
    if records_163["lab"]:
        n = insert_missing_rfis(gc, records_163["lab"],
                                SPREADSHEET_163,
                                build_table_name("163.22", "UNGUJA", "LAB"),
                                "163.22")
        total_inserted += n

    # 163.22 KM — SITE (Q RFIs)
    if records_163["site"]:
        n = insert_missing_rfis(gc, records_163["site"],
                                SPREADSHEET_163,
                                build_table_name("163.22", "UNGUJA", "SITE"),
                                "163.22")
        total_inserted += n

    # ── Summary ───────────────────────────────────────────────────────────
    print()
    print(SEP)
    print("  IMPORT COMPLETE")
    print(SEP)
    print(f"  Records from PDFs : {total_from_pdfs}")
    print(f"  New rows inserted : {total_inserted}")
    print(f"  Already existed   : {total_from_pdfs - total_inserted} (skipped — no overwrite)")
    print(f"  Completed at      : {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(SEP)
    print()


if __name__ == "__main__":
    main()
