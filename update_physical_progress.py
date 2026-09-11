"""
IRIS Physical Progress Updater  (v4 — April 2026)
==================================================
Reads executed lengths from the IRIS RFI Detailed Log Google Sheets
and writes activity-level physical progress values into both
Physical Progress Google Sheets.

WHAT IT DOES
------------
1. Reads IRIS RFI tables from:
     • 163.22 IRIS RFI Detailed Log  (Pemba + Unguja)
     • 65.5   IRIS RFI Detailed Log  (Pemba + Unguja)
2. Aggregates APPROVED lengths per (Road Code, PP Activity)
   using the "Merged Length (km)" column to avoid double-counting
   overlapping station ranges.
3. Reads the Physical Progress sheets to find every activity row
   and its "Additional + Contract Length OR Actual" cap.
4. Writes min(executed, cap) into the column matching today's
   scheduled date:
     • 65.5  : every 15th of the month
     • 163.22: every 19th of the month

ACTIVITY MAPPING  (IRIS RFI canonical → Physical Progress row label)
---------------------------------------------------------------------
  SITE CLEARANCE              → Site Clearance
  ROADBED PREPARATION         → Excavation/General Filling
  SUBGRADE LAYING (all)       → Excavation/General Filling
  SUBBASE LAYING              → Subbase
  BASECOURSE                  → Base Course
  PRIME COAT                  → Prime Coat
  ASPHALT LAYING              → Asphalt
  ROAD MARKING                → Road Marking
  TRAFFIC SIGNS INSTALLATION  → Traffic Sign
  CULVERT                     → Culvert
  SERVICE DUCTS               → Service Duct
  SIDE DRAINS                 → Side Drain
  SIDE DRAINS CLEARANCE       → Side Drain   (combined)
  (no RFI match)              → Low Water Crossing, Drift, Design  (skipped)

RUN
---
  python update_physical_progress.py

  Optional flags:
    --dry-run        Show what would be written, but don't write anything
    --date 2026-04-15   Override today's date (for back-filling)
    --project 65.5   Only update 65.5 project (or 163.22)
"""

from __future__ import annotations

import argparse
import json
import os
import re
import sys
from collections import defaultdict
from datetime import date, datetime
from pathlib import Path
from typing import Optional

# ── Google auth ───────────────────────────────────────────────────────────────
try:
    from google.oauth2.credentials import Credentials
    from google.auth.transport.requests import Request
    from googleapiclient.discovery import build
    from google.oauth2 import service_account
    HAS_GOOGLE = True
except ImportError:
    HAS_GOOGLE = False
    print("ERROR: Google API libraries not found.")
    print("Run:  pip install google-auth google-auth-oauthlib google-api-python-client")
    sys.exit(1)

# ══════════════════════════════════════════════════════════════════════════════
# CONFIGURATION
# ══════════════════════════════════════════════════════════════════════════════

BASE_DIR = Path(os.environ.get("IRIS_BASE_DIR") or (Path.home() / ".iris"))
try:
    BASE_DIR.mkdir(parents=True, exist_ok=True)
except Exception:
    pass
TOKEN_FILE = Path(os.environ.get("GOOGLE_TOKEN_PATH", str(BASE_DIR / "google_token.json")))
CREDS_FILE = Path(os.environ.get("GOOGLE_CREDS_PATH", str(BASE_DIR / "google_credentials.json")))
SERVICE_ACCOUNT_FILE = Path(os.environ["GOOGLE_SERVICE_ACCOUNT_PATH"]) if os.environ.get("GOOGLE_SERVICE_ACCOUNT_PATH") else None

# Physical Progress sheet IDs
PP_ID = {
    "65.5":   "16RE08Ivv2nSltqxVk1X9G8cBYcZP82kIGhzGAwRtZ7s",
    "163.22": "1YbDCzePl76CqtO9dLl2YZjIZtVmJKSgVkmU6LKqXKUc",
}

# IRIS RFI Detailed Log sheet IDs
RFI_ID = {
    "65.5":   "1WV58zo-jYgqOTBGOtMAjjlHC3WKdQ83s-XdRFlxZLLU",
    "163.22": "1IjN4W79VEZrAjevG5zZVKyJPuCdGbm_xAbcttTNjfOA",
}

# RFI_SHEET_NAMES: preferred tab names to try FIRST inside each IRIS RFI Detailed Log.
# If none are found the script auto-discovers all tabs and reads any that contain
# both "Road Code" and "Activity Name" columns — so it works regardless of tab naming.
RFI_SHEET_NAMES = {
    "65.5": [
        "KM 65.5 UNGUJA SITE RFI_DATA",
        "KM 65.5 UNGUJA LAB RFI_DATA",
        "KM 65.5 PEMBA SITE RFI_DATA",
        "KM 65.5 PEMBA LAB RFI_DATA",
        "65.5 KM SITE",
        "65.5 KM LAB",
        "Unguja SITE",
        "Unguja LAB",
        "Pemba SITE",
        "Pemba LAB",
        "ALL RFIs",
        "All RFIs",
    ],
    "163.22": [
        "KM 163.22 UNGUJA SITE RFI_DATA",
        "KM 163.22 UNGUJA LAB RFI_DATA",
        "KM 163.22 PEMBA SITE RFI_DATA",
        "KM 163.22 PEMBA LAB RFI_DATA",
        "163.22 KM SITE",
        "163.22 KM LAB",
        "Unguja SITE",
        "Unguja LAB",
        "Pemba SITE",
        "Pemba LAB",
        "ALL RFIs",
        "All RFIs",
    ],
}

# Physical Progress sheet name (tab name inside each Physical Progress file)
PP_SHEET_NAME = "PHYSICAL PROGRESS"

# Statuses to count as executed
APPROVED_STATUSES = {"APPROVED", "APPROVED WITH COMMENTS", "CLOSED", "OPENED"}

# ── Activity mapping: IRIS canonical name → Physical Progress label ──────────
# All SUBGRADE variants map to Excavation/General Filling (the earthworks activity)
IRIS_TO_PP: dict[str, str] = {
    "SITE CLEARANCE":              "Site Clearance",
    "ROADBED PREPARATION":         "Excavation/General Filling",
    "SUBGRADE LAYING":             "Excavation/General Filling",
    "SUBGRADE LAYING — 1st LAYER": "Excavation/General Filling",
    "SUBGRADE LAYING — 2nd LAYER": "Excavation/General Filling",
    "SUBGRADE LAYING — 3rd LAYER": "Excavation/General Filling",
    "SUBGRADE LAYING — 4th LAYER": "Excavation/General Filling",
    "SUBGRADE LAYING — 5th LAYER": "Excavation/General Filling",
    "SUBGRADE LAYING — 6th LAYER": "Excavation/General Filling",
    "SUBGRADE LAYING — nth LAYER": "Excavation/General Filling",
    "SUBBASE LAYING":              "Subbase",
    "BASECOURSE":                  "Base Course",
    "PRIME COAT":                  "Prime Coat",
    "ASPHALT LAYING":              "Asphalt",
    "ROAD MARKING":                "Road Marking",
    "TRAFFIC SIGNS INSTALLATION":  "Traffic Sign",
    "CULVERT":                     "Culvert",
    "SERVICE DUCTS":               "Service Duct",
    "SIDE DRAINS":                 "Side Drain",
    "SIDE DRAINS CLEARANCE":       "Side Drain",
    # Not mapped (no RFI data source):
    # Low Water Crossing, Drift, Design/Detailed Design → skipped
}

# Canonical Physical Progress activity labels (in order as they appear in sheets)
PP_ACTIVITIES_65_5 = [
    "Traffic Sign", "Service Duct", "Culvert", "Low Water Crossing", "Drift",
    "Side Drain", "Asphalt", "Prime Coat", "Base Course", "Road Marking",
    "Subbase", "Excavation/General Filling", "Site Clearance",
    "Design", "Detailed Design",
]
PP_ACTIVITIES_163_22 = [
    "Road Marking", "Traffic Sign", "Service Duct", "Culvert",
    "Low Water Crossing", "Drift", "Side Drain", "Asphalt",
    "Prime Coat", "Base Course", "Subbase",
    "Excavation/General Filling", "Site Clearance",
    "Design", "Detailed Design",
]

# Date schedules — dates must match column headers EXACTLY as written in the sheet
DATE_FORMAT_65_5   = "%-d-%b-%y"   # "15-Apr-26"
DATE_FORMAT_163_22 = "%-d-%b-%y"   # "19-Apr-26"
DATE_DAY_65_5   = 15
DATE_DAY_163_22 = 19


# ══════════════════════════════════════════════════════════════════════════════
# HELPERS
# ══════════════════════════════════════════════════════════════════════════════

def auth() -> object:
    """Return authenticated Google Sheets service."""
    # Service account short-circuit -- preferred for unattended/cloud runs
    # (no browser, no token to ever expire). Falls through to the existing
    # OAuth flow below when GOOGLE_SERVICE_ACCOUNT_PATH is not set.
    if SERVICE_ACCOUNT_FILE:
        if not SERVICE_ACCOUNT_FILE.exists():
            sys.exit(f"CONFIGURATION ERROR: Service account file not found: {SERVICE_ACCOUNT_FILE}")
        creds = service_account.Credentials.from_service_account_file(
            str(SERVICE_ACCOUNT_FILE),
            scopes=["https://www.googleapis.com/auth/spreadsheets"])
        return build("sheets", "v4", credentials=creds)

    if not TOKEN_FILE.exists():
        sys.exit(
            f"CONFIGURATION ERROR: Google token file not found: {TOKEN_FILE}. "
            "Set GOOGLE_TOKEN_PATH, or restore it from the GOOGLE_TOKEN_B64 secret."
        )

    try:
        with open(TOKEN_FILE, encoding="utf-8") as f:
            token_data = json.load(f)
    except json.JSONDecodeError as exc:
        sys.exit(
            f"CONFIGURATION ERROR: {TOKEN_FILE} is not valid JSON ({exc}). "
            "It is usually empty because the GOOGLE_TOKEN_B64 secret is missing or "
            "was not produced with 'base64 -w0 token.json'."
        )
    except OSError as exc:
        sys.exit(f"CONFIGURATION ERROR: cannot read {TOKEN_FILE}: {exc}")

    if not token_data.get("refresh_token"):
        sys.exit(
            f"CONFIGURATION ERROR: {TOKEN_FILE} contains no refresh_token. "
            "Re-authorise locally, then regenerate the GOOGLE_TOKEN_B64 secret."
        )
    creds = Credentials(
        token=token_data.get("token"),
        refresh_token=token_data.get("refresh_token"),
        token_uri=token_data.get("token_uri", "https://oauth2.googleapis.com/token"),
        client_id=token_data.get("client_id"),
        client_secret=token_data.get("client_secret"),
        scopes=token_data.get("scopes", ["https://www.googleapis.com/auth/spreadsheets"]),
    )
    if creds.expired and creds.refresh_token:
        import socket as _socket
        _old_timeout = _socket.getdefaulttimeout()
        _socket.setdefaulttimeout(30)   # don't hang >30 s waiting for OAuth endpoint
        try:
            creds.refresh(Request())
            # Persist refreshed token
            token_data["token"] = creds.token
            with open(TOKEN_FILE, "w") as f:
                json.dump(token_data, f, indent=2)
        except Exception as refresh_err:
            err_str = str(refresh_err).lower()
            if any(kw in err_str for kw in ("invalid_grant", "token has been expired",
                                             "token has been revoked", "revoked")):
                print(f"❌  Google token revoked/expired: {refresh_err}")
                print(f"↳  Deleting {TOKEN_FILE}")
                print(f"↳  Run update_rfi_google_sheets.py MANUALLY once to re-authenticate.")
                try:
                    os.remove(TOKEN_FILE)
                except OSError:
                    pass
                sys.exit(1)
            raise   # re-raise unexpected errors
        finally:
            _socket.setdefaulttimeout(_old_timeout)
    return build("sheets", "v4", credentials=creds)


def col_letter(n: int) -> str:
    """0-based column index → A, B, …, Z, AA, …"""
    result = ""
    n += 1
    while n:
        n, rem = divmod(n - 1, 26)
        result = chr(65 + rem) + result
    return result


def safe_float(v) -> float:
    """Parse floats, handling % sign and formatting."""
    try:
        s = str(v).replace(",", "").replace(" ", "").strip()
        if s.endswith('%'):
            s = s[:-1]
        return float(s) if s else 0.0
    except (ValueError, TypeError):
        return 0.0


def normalise_label(s: str) -> str:
    """Lower-case, collapse spaces, strip punctuation for fuzzy matching."""
    return re.sub(r"[^a-z0-9]", "", str(s).lower())


def match_activity(cell_text: str, candidates: list[str]) -> Optional[str]:
    """
    Fuzzy-match a cell value to one of the PP_ACTIVITIES list.
    Returns the matched canonical name or None.
    """
    ct = normalise_label(cell_text)
    if not ct:
        return None
    # Try exact normalised match first
    for act in candidates:
        if normalise_label(act) == ct:
            return act
    # Substring match
    for act in candidates:
        if normalise_label(act) in ct or ct in normalise_label(act):
            return act
    # Special cases
    if "excavat" in ct or "generalfill" in ct or "earthwork" in ct:
        return "Excavation/General Filling"
    if "subgrade" in ct:
        return "Excavation/General Filling"
    return None


def find_date_column(header_row: list, target_date: date) -> Optional[int]:
    """
    Search a header row for a cell matching the target date.
    Accepts multiple formats: "15-Apr-26", "15/04/2026", "Apr 15, 2026", etc.
    Returns 0-based column index within the row, or None.
    """
    formats = [
        "%d-%b-%y", "%-d-%b-%y",
        "%d/%m/%Y", "%d-%m-%Y",
        "%b %d, %Y", "%B %d, %Y",
        "%Y-%m-%d",
    ]
    for i, cell in enumerate(header_row):
        s = str(cell).strip()
        for fmt in formats:
            try:
                d = datetime.strptime(s, fmt).date()
                if d == target_date:
                    return i
            except ValueError:
                continue
        # Try partial: "15-Apr" without year — supply the year explicitly
        try:
            d2 = datetime.strptime(f"{s}-{target_date.year}", "%d-%b-%Y").date()
            if d2 == target_date:
                return i
        except ValueError:
            pass
    return None


def scheduled_date_for(project: str, today: date) -> date:
    """
    Return the UPCOMING scheduled date: if today is between (MMM-1)-day and MMM-day,
    write to MMM-day.

    65.5  → 15th of the month  (use THIS month's 15th if today <= 15th, else next month's)
    163.22→ 19th of the month  (use THIS month's 19th if today <= 19th, else next month's)
    """
    import calendar
    day = DATE_DAY_65_5 if project == "65.5" else DATE_DAY_163_22

    # If today <= scheduled day of this month, use this month's day
    if today.day <= day:
        max_day = calendar.monthrange(today.year, today.month)[1]
        return today.replace(day=min(day, max_day))

    # today > scheduled day → next month's day
    if today.month == 12:
        return date(today.year + 1, 1, day)
    return date(today.year, today.month + 1, day)


# ══════════════════════════════════════════════════════════════════════════════
# STEP 1 — READ IRIS RFI EXECUTED LENGTHS
# ══════════════════════════════════════════════════════════════════════════════

def _is_rfi_table(header: list[str]) -> bool:
    """Return True if this sheet looks like an RFI data table."""
    hl = {h.strip().lower() for h in header}
    return ("road code" in hl or "road_code" in hl) and \
           ("activity name" in hl or "activity_name" in hl)


def _read_sheet_records(service, ssid: str, sheet_name: str,
                         project: str) -> list[dict]:
    """Read one sheet tab and return a list of raw record dicts."""
    try:
        result = service.spreadsheets().values().get(
            spreadsheetId=ssid,
            range=f"'{sheet_name}'!A:Z",
        ).execute()
        rows = result.get("values", [])
        if len(rows) < 2:
            return []

        # Find first non-empty header row (skip blank rows at top)
        header_row_idx = 0
        for i, row in enumerate(rows[:5]):
            if any(str(c).strip() for c in row):
                header_row_idx = i
                break

        header = [str(h).strip() for h in rows[header_row_idx]]
        if not _is_rfi_table(header):
            return []

        hmap    = {h: i for i, h in enumerate(header)}
        hmap_lc = {h.lower(): i for i, h in enumerate(header)}

        def gcol(name: str) -> Optional[int]:
            v = hmap.get(name)
            if v is not None: return v
            return hmap_lc.get(name.lower())

        col_road_code = gcol("Road Code")
        col_island    = gcol("Island")
        col_activity  = gcol("Activity Name")
        col_status    = gcol("Status")
        col_merged    = gcol("Merged Length (km)")
        col_length    = gcol("Length Executed (km)")
        col_start     = gcol("Start Station")
        col_end       = gcol("End Station")
        col_project   = gcol("Project")  # may contain "65.5 KM" or "163.22 KM"

        records = []
        for row in rows[header_row_idx + 1:]:
            def cell(c):
                return str(row[c]).strip() if c is not None and c < len(row) else ""

            road_code = cell(col_road_code).upper()
            island    = cell(col_island).upper()
            activity  = cell(col_activity).strip()
            status    = cell(col_status).upper()
            proj_cell = cell(col_project).upper() if col_project is not None else ""

            if not road_code or not activity:
                continue

            # Filter by project if the Project column is present
            if proj_cell:
                expected = project.replace(".", r"\.")
                if not re.search(expected, proj_cell):
                    continue

            if status and status not in APPROVED_STATUSES:
                continue

            pp_act = IRIS_TO_PP.get(activity)
            if pp_act is None:
                # Also try stripping layer suffixes: "SUBGRADE LAYING — 2nd LAYER" → "SUBGRADE LAYING"
                base_act = re.sub(r"\s*—.*$", "", activity).strip()
                pp_act = IRIS_TO_PP.get(base_act)
            if pp_act is None:
                continue

            merged_km = safe_float(cell(col_merged))
            exec_km   = safe_float(cell(col_length))
            km = merged_km if merged_km > 0 else exec_km

            if km <= 0:
                try:
                    s_m = _parse_station(cell(col_start)) if col_start is not None else 0.0
                    e_m = _parse_station(cell(col_end))   if col_end   is not None else 0.0
                    km  = abs(e_m - s_m) / 1000.0
                except Exception:
                    km = 0.0

            if km <= 0:
                continue

            records.append({
                "road_code": road_code,
                "island":    island if island else "UNGUJA",
                "pp_act":    pp_act,
                "km":        km,
                "start":     safe_float(cell(col_start)) if col_start is not None else 0.0,
                "end":       safe_float(cell(col_end))   if col_end   is not None else 0.0,
            })
        return records
    except Exception as e:
        print(f"      Warning reading '{sheet_name}': {e}")
        return []


def read_rfi_data(service, project: str) -> dict:
    """
    Read all APPROVED RFIs from the IRIS RFI Detailed Log for `project`.

    Strategy:
      1. Try the preferred sheet name list (RFI_SHEET_NAMES).
      2. If still 0 records, auto-discover ALL tabs and read any that look
         like RFI data tables (contain Road Code + Activity Name columns).
      3. Also try the combined spreadsheet (rfi_spreadsheet_id.txt).

    Returns combined dict with keys of length 3 AND length 2:
      (road_code, island, pp_activity) → km
      (road_code, pp_activity)         → km  (island-agnostic fallback)
    """
    ssid = RFI_ID[project]
    print(f"\n  Reading IRIS RFI data for {project} KM project …")

    # Get all available tab names
    try:
        meta = service.spreadsheets().get(spreadsheetId=ssid).execute()
        all_tabs = [s["properties"]["title"] for s in meta.get("sheets", [])]
    except Exception as e:
        print(f"    ERROR accessing RFI sheet {ssid}: {e}")
        all_tabs = []

    print(f"    Tabs in spreadsheet: {all_tabs or '(none reachable)'}")

    raw_records: list[dict] = []
    tabs_read: set[str] = set()

    # ── Pass 1: preferred tab names ──────────────────────────────────────────
    for sheet_name in RFI_SHEET_NAMES[project]:
        if sheet_name in all_tabs and sheet_name not in tabs_read:
            recs = _read_sheet_records(service, ssid, sheet_name, project)
            if recs:
                print(f"    ✓ {len(recs):4d} records from '{sheet_name}'")
                raw_records.extend(recs)
                tabs_read.add(sheet_name)

    # ── Pass 2: auto-discover any remaining tabs ─────────────────────────────
    for sheet_name in all_tabs:
        if sheet_name in tabs_read:
            continue
        recs = _read_sheet_records(service, ssid, sheet_name, project)
        if recs:
            print(f"    ✓ {len(recs):4d} records (auto) from '{sheet_name}'")
            raw_records.extend(recs)
            tabs_read.add(sheet_name)

    # ── Pass 3: try the combined RFI spreadsheet ─────────────────────────────
    combined_id_file = BASE_DIR / "rfi_spreadsheet_id.txt"
    if len(raw_records) == 0 and combined_id_file.exists():
        combined_ssid = combined_id_file.read_text().strip()
        if combined_ssid and combined_ssid != ssid:
            print(f"    Trying combined RFI spreadsheet …")
            try:
                meta2 = service.spreadsheets().get(spreadsheetId=combined_ssid).execute()
                for s in meta2.get("sheets", []):
                    tname = s["properties"]["title"]
                    recs = _read_sheet_records(service, combined_ssid, tname, project)
                    if recs:
                        print(f"    ✓ {len(recs):4d} records (combined) from '{tname}'")
                        raw_records.extend(recs)
            except Exception as e:
                print(f"    Warning on combined sheet: {e}")

    print(f"    Total raw records: {len(raw_records)}")

    # Aggregate: for each (road_code, island, pp_act) merge station ranges
    grouped: dict[tuple, list[tuple[float, float]]] = defaultdict(list)
    for rec in raw_records:
        key = (rec["road_code"], rec["island"], rec["pp_act"])
        if rec["start"] > 0 or rec["end"] > 0:
            start_m = min(rec["start"], rec["end"])
            end_m   = max(rec["start"], rec["end"])
            grouped[key].append((start_m, end_m))
        else:
            # No station data — use km directly as a single segment
            grouped[key].append((0.0, rec["km"] * 1000))

    result: dict[tuple, float] = {}
    for key, segs in grouped.items():
        result[key] = round(_merge_segments(segs) / 1000, 4)

    # Also build road_code+pp_act (without island) for fallback lookup
    road_act: dict[tuple[str, str], float] = defaultdict(float)
    for (rc, isl, act), km in result.items():
        road_act[(rc, act)] += km

    # Return combined dict with both (road, island, act) and (road, act) keys
    combined = {}
    for k, v in result.items():
        combined[k] = v
    for k, v in road_act.items():
        combined[k] = v   # tuple length 2 as fallback

    print(f"    Aggregated: {len(result)} (road, island, activity) combinations")
    return combined


def _parse_station(s: str) -> float:
    """Parse '12+500' → 12500.0 (metres). Also handles plain numbers."""
    s = str(s).strip()
    m = re.match(r"(\d+)\+(\d+(?:\.\d+)?)", s)
    if m:
        return float(m.group(1)) * 1000 + float(m.group(2))
    try:
        v = float(s.replace(",", ""))
        # If large (>1000) treat as metres, else as km
        return v if v > 1000 else v * 1000
    except ValueError:
        return 0.0


def _merge_segments(segs: list[tuple[float, float]]) -> float:
    """Merge overlapping [start, end] segments and return total length in metres."""
    if not segs:
        return 0.0
    merged = sorted(segs, key=lambda x: x[0])
    result = []
    cs, ce = merged[0]
    for s, e in merged[1:]:
        if s <= ce:
            ce = max(ce, e)
        else:
            result.append((cs, ce))
            cs, ce = s, e
    result.append((cs, ce))
    return sum(e - s for s, e in result)


# ══════════════════════════════════════════════════════════════════════════════
# STEP 2 — PARSE PHYSICAL PROGRESS SHEET
# ══════════════════════════════════════════════════════════════════════════════

def parse_pp_sheet(service, project: str) -> dict:
    """
    Read the Physical Progress sheet and return a structure describing
    every activity row, its contract length cap, and which column the
    target date falls in.

    Returns:
    {
      sheet_title: str,
      date_col_idx: int,       # 0-based within the row
      target_date: date,
      ipo_col: int or None,    # column of IPO formula
      ocw_col: int or None,    # column of OCW formula
      prev_date_col: int or None, # previous date column
      all_date_cols: [(date, col_idx), ...],  # all detected date columns
      roads: [
        {
          road_code: str,
          road_name: str,
          island: str,
          contract_length: float,
          header_row_idx: int,  # 0-based sheet row of the road header
          ipo_value: float,     # IPO value at road header
          ocw_value: float,     # OCW value at road header
          activities: [
            {
              pp_activity: str,     # canonical PP label
              row_idx: int,         # 0-based sheet row of this activity
              col_idx: int,         # 0-based column of the date cell
              contract_km: float,   # cap value for this activity
              current_value: float, # existing value in this cell
              prev_value: float,    # value from previous date column
              ipo_value: float,     # IPO value for this activity
              ocw_value: float,     # OCW value for this activity
            }
          ]
        }
      ]
    }
    """
    ssid = PP_ID[project]
    today = _today_override if _today_override else date.today()
    target = scheduled_date_for(project, today)
    activities_list = PP_ACTIVITIES_65_5 if project == "65.5" else PP_ACTIVITIES_163_22

    print(f"\n  Parsing Physical Progress sheet for {project} KM …")
    print(f"  Target date: {target.strftime('%d-%b-%y')}")

    try:
        meta = service.spreadsheets().get(spreadsheetId=ssid).execute()
        sheet_names = [s["properties"]["title"] for s in meta.get("sheets", [])]
    except Exception as e:
        print(f"    ERROR: {e}")
        return {}

    # Find the Physical Progress sheet (tab)
    pp_tab = None
    for name in sheet_names:
        if "physical progress" in name.lower():
            pp_tab = name
            break
    if pp_tab is None:
        pp_tab = sheet_names[0]
        print(f"    WARNING: 'Physical Progress' tab not found, using '{pp_tab}'")

    # Read all data
    result = service.spreadsheets().values().get(
        spreadsheetId=ssid,
        range=f"'{pp_tab}'!A:BZ",   # BZ = col 78; covers AJ(36)/AK(37) even with gaps
    ).execute()
    all_rows = result.get("values", [])
    print(f"    Total rows in sheet: {len(all_rows)}")

    # ── Find date header row and the target date column ──────────────────────
    date_col_idx   = None
    date_header_ri = None
    contract_col   = None   # "Additional + Contract Length OR Actual"
    road_code_col  = None
    road_name_col  = None
    ipo_col        = None
    ocw_col        = None
    all_date_cols  = []
    prev_date_col  = None

    for ri, row in enumerate(all_rows):
        # Look for a row containing the target date
        dc = find_date_column(row, target)
        if dc is not None:
            date_col_idx   = dc
            date_header_ri = ri
            print(f"    Found target date at row {ri+1}, col {col_letter(dc)}")

            # ── Narrow scan (±3 rows): structural & date columns ──────────────
            for offset in range(-3, 4):
                hri = ri + offset
                if 0 <= hri < len(all_rows):
                    hr = all_rows[hri]
                    for ci, cell in enumerate(hr):
                        ct       = str(cell).strip()
                        ct_lower = ct.lower()

                        # Detect contract/cap column (must come first to avoid
                        # being overwritten by IPO/OCW which also contain these words)
                        if contract_col is None and \
                           re.search(r"additional.*contract.*length|contract.*additional.*length", ct_lower):
                            contract_col = ci

                        # Detect road code and name columns
                        if road_code_col is None and ct_lower in ("road code", "code"):
                            road_code_col = ci
                        if road_name_col is None and (
                            "road name" in ct_lower or
                            ("road" in ct_lower and "name" in ct_lower)
                        ):
                            road_name_col = ci

                        # Collect date columns (dedup by col index later)
                        date_obj = _try_parse_date(ct)
                        if date_obj:
                            all_date_cols.append((date_obj, ci))

            # ── Wide scan (first 15 rows): IPO / OCW columns ──────────────────
            # IPO and OCW sit far to the right (col AJ/AK ≈ index 35/36).
            # Between the last date col and IPO there are often empty cells,
            # which can truncate API rows — so we scan ALL first 15 rows.
            for scan_ri in range(min(15, len(all_rows))):
                hr = all_rows[scan_ri]
                for ci, cell in enumerate(hr):
                    ct       = str(cell).strip()
                    ct_lower = ct.lower()
                    if ipo_col is None and (
                        re.search(r"individual progress acc\.?\s*to overall", ct_lower) or
                        re.search(r"acc\.?\s*to overall contract", ct_lower) or
                        re.search(r"overall contract.*lengths.*sum", ct_lower)
                    ):
                        ipo_col = ci
                    if ocw_col is None and ci != ipo_col and (
                        re.search(r"overall contract.*weight", ct_lower) or
                        re.search(r"contract.*additional.*weight", ct_lower) or
                        re.search(r"overall.*additional.*weight", ct_lower)
                    ):
                        ocw_col = ci
                    if ipo_col is not None and ocw_col is not None:
                        break
                if ipo_col is not None and ocw_col is not None:
                    break

            # Deduplicate all_date_cols by column index (keep earliest date per col)
            _seen_cols: dict[int, date] = {}
            for d, ci in all_date_cols:
                if ci not in _seen_cols:
                    _seen_cols[ci] = d
            all_date_cols = sorted(_seen_cols.items(), key=lambda x: x[1])  # (col,date) sorted by date
            all_date_cols = [(d, ci) for ci, d in all_date_cols]            # back to (date,col)

            break

    if date_col_idx is None:
        # Try to find the column by scanning ALL header-like rows
        print(f"    WARNING: Target date {target} not found in any header row.")
        print(f"    Will scan for nearest date column…")
        date_col_idx, date_header_ri = _find_nearest_date_col(all_rows, target)
        if date_col_idx is None:
            print(f"    ERROR: Cannot find a date column near {target}.")
            return {}

    # Find prev_date_col: the date column immediately before target date
    if all_date_cols:
        all_date_cols.sort(key=lambda x: x[0])
        for i, (d, ci) in enumerate(all_date_cols):
            if d == target and i > 0:
                prev_date_col = all_date_cols[i - 1][1]
                break

    print(f"    IPO col: {col_letter(ipo_col) if ipo_col else 'NOT FOUND'}, "
          f"OCW col: {col_letter(ocw_col) if ocw_col else 'NOT FOUND'}, "
          f"Prev date col: {col_letter(prev_date_col) if prev_date_col else 'NONE'}, "
          f"Date range: {len(all_date_cols)} columns")

    # ── Walk rows to find roads and their activity sub-rows ─────────────────
    roads = []
    current_road = None

    for ri, row in enumerate(all_rows):
        if ri <= (date_header_ri or 0):
            continue  # skip header rows

        # Pad row to at least date_col_idx + 1
        while len(row) <= date_col_idx:
            row.append("")

        # ── Detect ISLAND divider row ──────────────────────────────────────
        row_text = " ".join(str(c).strip() for c in row[:6]).upper()
        if re.search(r"\bUNGUJA\b|\bPEMBA\b", row_text) and len([c for c in row if str(c).strip()]) <= 3:
            island_hint = "UNGUJA" if "UNGUJA" in row_text else "PEMBA"
            continue  # section label row; island_hint set for next roads

        # ── Detect ROAD row ────────────────────────────────────────────────
        # Road rows have a Road Code in a dedicated column (or just a recognisable
        # alphanumeric code like "NA-200", "SO-201", "102", "103" …)
        potential_code = _extract_road_code(row, road_code_col)
        if potential_code:
            # Determine island from road name / code context
            island = _island_from_context(row, row_text, project)
            road_name = str(row[road_name_col]).strip() if road_name_col and road_name_col < len(row) else _best_road_name(row, road_code_col)
            contract_km = safe_float(row[contract_col]) if contract_col and contract_col < len(row) else 0.0
            road_ipo = safe_float(row[ipo_col]) if ipo_col and ipo_col < len(row) else 0.0
            road_ocw = safe_float(row[ocw_col]) if ocw_col and ocw_col < len(row) else 0.0
            current_road = {
                "road_code":       potential_code,
                "road_name":       road_name,
                "island":          island,
                "contract_length": contract_km,
                "header_row_idx":  ri,
                "ipo_value":       road_ipo,
                "ocw_value":       road_ocw,
                "activities":      [],
            }
            roads.append(current_road)
            continue

        # ── Detect ACTIVITY sub-row ───────────────────────────────────────
        if current_road is None:
            continue

        # Activity rows: first non-empty cell in first 4 cols likely holds
        # the activity name (or a sequence number + activity name)
        act_label = _extract_activity_label(row)
        matched = match_activity(act_label, activities_list) if act_label else None
        if matched:
            # Contract km for this activity: either from its own contract-length
            # cell, or derived from road contract * activity weight
            act_contract = safe_float(row[contract_col]) if contract_col and contract_col < len(row) else 0.0
            if act_contract <= 0:
                # Try to get weight and multiply by road contract length
                weight = _extract_weight(row)
                act_contract = current_road["contract_length"] * weight if weight else 0.0

            # Use contract_length from the "Additional + Contract Length OR Actual" col
            # (this supersedes any weight calculation if it's a direct km value)
            if contract_col and contract_col < len(row):
                direct = safe_float(row[contract_col])
                if direct > 0:
                    act_contract = direct

            current_value = safe_float(row[date_col_idx])
            prev_value = safe_float(row[prev_date_col]) if prev_date_col is not None else 0.0
            act_ipo = safe_float(row[ipo_col]) if ipo_col is not None else 0.0
            act_ocw = safe_float(row[ocw_col]) if ocw_col is not None else 0.0

            current_road["activities"].append({
                "pp_activity":   matched,
                "row_idx":       ri,
                "col_idx":       date_col_idx,
                "contract_km":   act_contract,
                "current_value": current_value,
                "prev_value":    prev_value,
                "ipo_value":     act_ipo,
                "ocw_value":     act_ocw,
                "raw_label":     act_label,
            })

    total_acts = sum(len(r["activities"]) for r in roads)
    print(f"    Found {len(roads)} road rows with {total_acts} activity rows")
    return {
        "sheet_title":   pp_tab,
        "ssid":          ssid,
        "date_col_idx":  date_col_idx,
        "target_date":   target,
        "ipo_col":       ipo_col,
        "ocw_col":       ocw_col,
        "prev_date_col": prev_date_col,
        "all_date_cols": all_date_cols,
        "roads":         roads,
    }


def _try_parse_date(cell_text: str) -> Optional[date]:
    """Try to parse a cell as a date object."""
    s = str(cell_text).strip()
    formats = [
        "%d-%b-%y", "%-d-%b-%y",
        "%d/%m/%Y", "%d-%m-%Y",
        "%b %d, %Y", "%B %d, %Y",
        "%Y-%m-%d",
    ]
    for fmt in formats:
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def _find_nearest_date_col(all_rows, target):
    """Find the column whose header is closest to target date."""
    import calendar
    best_col = None
    best_ri  = None
    best_diff = 999999
    for ri, row in enumerate(all_rows[:30]):  # scan first 30 rows
        for ci, cell in enumerate(row):
            s = str(cell).strip()
            for fmt in ["%d-%b-%y", "%-d-%b-%y", "%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d"]:
                try:
                    d = datetime.strptime(s, fmt).date()
                    diff = abs((d - target).days)
                    if diff < best_diff:
                        best_diff = diff
                        best_col = ci
                        best_ri  = ri
                except ValueError:
                    pass
    if best_col is not None:
        print(f"    Nearest date column: row {best_ri+1}, col {col_letter(best_col)} (diff {best_diff} days)")
    return best_col, best_ri


def _extract_road_code(row: list, code_col: Optional[int]) -> Optional[str]:
    """
    Try to extract a road code from a row.
    Valid road codes: "NA-200", "SO-201", "CE-102", "WB-218",
                     "102", "103", "200", "201" etc.
    """
    candidates = []
    if code_col is not None and code_col < len(row):
        candidates.append(str(row[code_col]).strip())
    # Also check first 5 cells
    for cell in row[:5]:
        candidates.append(str(cell).strip())

    for c in candidates:
        # Explicit letter-number code like NA-200, SO-201, WB-217
        if re.match(r"^[A-Z]{1,3}-\d{3}$", c):
            return c
        # Numeric road numbers 100-299
        if re.match(r"^\d{3}$", c) and 100 <= int(c) <= 299:
            return c
        # Project-code prefixes like "102", "103"
    return None


def _extract_activity_label(row: list) -> str:
    """Return the first meaningful non-numeric cell value in the first 5 cols."""
    for cell in row[:5]:
        s = str(cell).strip()
        if s and not re.match(r"^\d+$", s) and len(s) > 2:
            return s
    return ""


def _extract_weight(row: list) -> float:
    """Extract percentage weight from any cell in first 6 cols. Returns 0-1 float."""
    for cell in row[:6]:
        s = str(cell).strip()
        m = re.search(r"(\d+(?:\.\d+)?)\s*%", s)
        if m:
            return float(m.group(1)) / 100.0
    return 0.0


def _island_from_context(row: list, row_text: str, project: str) -> str:
    """Guess island (UNGUJA / PEMBA) from row content."""
    rt = row_text.upper()
    if "UNGUJA" in rt:
        return "UNGUJA"
    if "PEMBA" in rt:
        return "PEMBA"
    # Road codes: NA, SO, CE, WB → Unguja; others may be Pemba
    for cell in row[:6]:
        c = str(cell).strip().upper()
        if re.match(r"^(NA|SO|CE|WB|NB)-\d{3}$", c):
            return "UNGUJA"
    return "UNGUJA"  # default


def _best_road_name(row: list, code_col: Optional[int]) -> str:
    """Return the longest non-code cell in the first 8 columns."""
    skip_col = code_col if code_col is not None else -1
    best = ""
    for ci, cell in enumerate(row[:8]):
        if ci == skip_col:
            continue
        s = str(cell).strip()
        if len(s) > len(best) and not re.match(r"^\d+$", s):
            best = s
    return best


# ══════════════════════════════════════════════════════════════════════════════
# STEP 3 — COMPUTE AND WRITE UPDATES
# ══════════════════════════════════════════════════════════════════════════════

def _code_variants(rc: str) -> list[str]:
    """
    Generate all plausible variants of a road code for fuzzy matching.

    Examples:
      "CE-102"  -> ["CE-102", "CE102", "102", "SO-102", ...]
      "AR-227"  -> ["AR-227", "AR227", "227", "CE-227", ...]
      "102"     -> ["102", "CE-102", "CE102", "SO-102", "SO102", ...]
    """
    rc = rc.strip().upper()
    seen: set[str] = set()
    variants: list[str] = []

    def _add(v: str):
        v = v.strip().upper()
        if v and v not in seen:
            seen.add(v)
            variants.append(v)

    _add(rc)

    # All district/road prefixes used in Zanzibar
    PREFIXES = (
        "CE", "SO", "NA", "WB", "NB", "WA", "MK", "MI",
        "CH", "WE", "AR", "KT", "UC", "UN", "KB", "UM",
        "PEM", "UNG",
    )

    # Strip letter prefix: "CE-102" or "AR-227" or "CE102" -> numeric part
    m = re.match(r"^([A-Z]{1,3})[-_]?(\d+)$", rc)
    if m:
        pfx_found, num = m.group(1), m.group(2)
        _add(f"{pfx_found}{num}")   # without separator: "CE-102" -> "CE102"
        _add(num)                   # bare numeric: "102"
        # Cross-prefix variants (road may be re-coded across districts)
        for pfx in PREFIXES:
            if pfx != pfx_found:
                _add(f"{pfx}-{num}")
                _add(f"{pfx}{num}")

    # Purely numeric: try all known prefixes with and without hyphen
    elif re.match(r"^\d+$", rc):
        for pfx in PREFIXES:
            _add(f"{pfx}-{rc}")
            _add(f"{pfx}{rc}")

    return variants


def compute_updates(pp_data: dict, rfi_data: dict) -> list[dict]:
    """
    Match RFI executed lengths to PP activity rows and compute write values.
    Uses fuzzy road-code matching (strips/adds letter prefixes) so that
    "102" matches "CE-102" and vice versa.

    Returns list of:
    {road_code, road_name, island, pp_activity, row_idx, col_idx,
     contract_km, executed_km, write_value, current_value, prev_value, changed}
    """
    updates = []

    # Pre-compute: for each (road_code_variant, activity) pair that EXISTS in
    # rfi_data, build a fast-lookup set.  This lets us distinguish:
    #   A) road+activity genuinely present in RFI data but lookup returned 0
    #      → real mismatch, worth flagging
    #   B) road present in RFI for OTHER activities but not this one
    #      → uninspected activity, silence it (not a bug)
    rfi_road_acts: set[tuple[str, str]] = set()
    for k in rfi_data:
        # k is either (road_code, activity)  or  (road_code, island, activity)
        rfi_road_acts.add((k[0], k[-1]))   # k[0]=road, k[-1]=activity in both cases

    # Mismatches: {road_code: [activity, ...]} — only populated for case A
    mismatches: dict[str, list[str]] = {}

    for road in pp_data.get("roads", []):
        rc     = road["road_code"].upper()
        island = road["island"].upper()
        variants = _code_variants(rc)

        for act in road["activities"]:
            pp_act      = act["pp_activity"]
            contract_km = act["contract_km"]
            executed    = 0.0

            # Try all code variants × island combinations
            for variant in variants:
                if executed > 0:
                    break
                executed = rfi_data.get((variant, island, pp_act), 0.0)
                if executed == 0.0:
                    executed = rfi_data.get((variant, pp_act), 0.0)
                if executed == 0.0:
                    for isl in ("UNGUJA", "PEMBA"):
                        executed = rfi_data.get((variant, isl, pp_act), 0.0)
                        if executed > 0:
                            break

            # Only flag when rfi_data actually CONTAINS (some variant, this activity)
            # but the lookup still returned 0 — this is a genuine key/format mismatch.
            # Roads that simply haven't been inspected for this activity are silent.
            if executed == 0.0:
                real_mismatch = any(
                    (v, pp_act) in rfi_road_acts for v in variants
                )
                if real_mismatch:
                    mismatches.setdefault(rc, []).append(pp_act)

            write_val = round(min(executed, contract_km), 4) if contract_km > 0 else round(executed, 4)

            # Non-decreasing rule: never write less than previous period's value
            prev_val = act.get("prev_value", 0.0)
            write_val = max(prev_val, write_val)
            # Also never write less than what's already in the cell
            write_val = max(act["current_value"], write_val)
            changed = abs(write_val - act["current_value"]) > 0.0001

            updates.append({
                "road_code":     rc,
                "road_name":     road["road_name"],
                "island":        island,
                "pp_activity":   pp_act,
                "row_idx":       act["row_idx"],
                "col_idx":       act["col_idx"],
                "contract_km":   contract_km,
                "executed_km":   round(executed, 4),
                "write_value":   write_val,
                "current_value": act["current_value"],
                "prev_value":    prev_val,
                "changed":       changed,
            })

    if mismatches:
        # These are genuine lookup failures: rfi_data has data for this
        # road+activity pair (under some code variant) but the lookup returned 0.
        # Likely cause: road code uses a separator format not yet in _code_variants.
        total = sum(len(v) for v in mismatches.values())
        print(f"    [NOTE] {total} road/activity lookup(s) found in RFI index "
              f"but returned 0 km — possible code-format mismatch:")
        for rc_key in sorted(mismatches)[:10]:   # cap at 10 to avoid log flood
            acts = ", ".join(mismatches[rc_key])
            print(f"           {rc_key}: {acts}")
        if len(mismatches) > 10:
            print(f"           ... and {len(mismatches)-10} more road(s)")
    return updates


def write_updates(service, ssid: str, sheet_title: str,
                  updates: list[dict], dry_run: bool = False) -> int:
    """
    Batch-write all changed values to the Physical Progress sheet.
    Returns count of cells written.
    """
    batch = []
    written = 0
    for u in updates:
        if not u["changed"]:
            continue
        row_1b = u["row_idx"] + 1          # 1-based
        col_1b = u["col_idx"] + 1
        a1 = f"'{sheet_title}'!{col_letter(u['col_idx'])}{row_1b}"
        batch.append({
            "range": a1,
            "values": [[u["write_value"]]],
        })
        written += 1
        print(f"    {'DRY ' if dry_run else ''}WRITE  {u['road_code']:10s} "
              f"{u['pp_activity']:30s}  "
              f"{u['current_value']:.4f} → {u['write_value']:.4f} km  "
              f"[{a1}]")

    if not dry_run and batch:
        service.spreadsheets().values().batchUpdate(
            spreadsheetId=ssid,
            body={
                "valueInputOption": "USER_ENTERED",
                "data": batch,
            },
        ).execute()
        print(f"    ✓ {written} cells written to Google Sheets")
    elif dry_run:
        print(f"    DRY RUN: {written} cells would be written")
    return written


# ══════════════════════════════════════════════════════════════════════════════
# STEP 4 — READ BACK FORMULA PROGRESS ROWS FROM PHYSICAL PROGRESS SHEET
# ══════════════════════════════════════════════════════════════════════════════

def read_progress_pct(service, ssid: str, sheet_title: str,
                      pp_data: dict) -> dict[str, dict]:
    """
    After km values have been written, read back the formula rows using column
    indices (IPO and OCW columns detected in parse_pp_sheet).

    Reads UNFORMATTED_VALUE to get raw decimal numbers (not "1.19%" strings).
    For each road, progress_pct = IPO / OCW * 100  (clamped 0–100).

    Returns:
        {road_code: {
            "progress_pct":   float (0–100),
            "ipo_pct":        float,   # Individual Progress % (IPO × 100)
            "ocw_pct":        float,   # Overall Contract Weight % (OCW × 100)
            "island":         str,
            "road_name":      str,
            "contract_km":    float,
            "activities": {
                "activity_name": {
                    "progress_pct": float,
                    "ipo_pct": float,
                    "ocw_pct": float,
                }
            }
        }}
    """
    ipo_col = pp_data.get("ipo_col")
    ocw_col = pp_data.get("ocw_col")
    roads   = pp_data.get("roads", [])

    if ipo_col is None or ocw_col is None:
        print("    ⚠  IPO/OCW columns not found — progress % unavailable")
        return {}

    # Re-read sheet with UNFORMATTED_VALUE to get raw numbers (not "1.19%" strings)
    try:
        result = service.spreadsheets().values().get(
            spreadsheetId=ssid,
            range=f"'{sheet_title}'!A:BZ",   # wide enough to cover IPO/OCW at AJ/AK
            valueRenderOption="UNFORMATTED_VALUE",
        ).execute()
        all_rows = result.get("values", [])
    except Exception as e:
        print(f"    Warning reading progress rows: {e}")
        return {}

    def cell_num(ri, ci):
        if ri >= len(all_rows) or ci >= len(all_rows[ri]):
            return 0.0
        return safe_float(all_rows[ri][ci])

    out = {}
    for road in roads:
        rc = road["road_code"].upper()
        ri = road["header_row_idx"]

        # Read road-level IPO and OCW (raw decimal from UNFORMATTED_VALUE, e.g. 0.0119 for 1.19%)
        ipo_raw = cell_num(ri, ipo_col)   # e.g. 0.0119
        ocw_raw = cell_num(ri, ocw_col)   # e.g. 0.0119

        # Convert to percentage units (multiply by 100)
        ipo_pct = round(ipo_raw * 100, 6)
        ocw_pct = round(ocw_raw * 100, 6)

        if ocw_pct > 0:
            progress_pct = round(min(100.0, (ipo_pct / ocw_pct) * 100), 2)
        else:
            progress_pct = 0.0

        # Also read per-activity IPO and OCW
        act_progress = {}
        for act in road.get("activities", []):
            act_ri   = act["row_idx"]
            act_ipo  = round(cell_num(act_ri, ipo_col) * 100, 6)
            act_ocw  = round(cell_num(act_ri, ocw_col) * 100, 6)
            act_pct  = round(min(100.0, (act_ipo / act_ocw) * 100), 2) if act_ocw > 0 else 0.0
            act_progress[act["pp_activity"]] = {
                "progress_pct": act_pct,
                "ipo_pct":      act_ipo,
                "ocw_pct":      act_ocw,
            }

        out[rc] = {
            "progress_pct":   progress_pct,
            "ipo_pct":        ipo_pct,
            "ocw_pct":        ocw_pct,
            "island":         road["island"],
            "road_name":      road["road_name"],
            "contract_km":    road.get("contract_length", 0.0),
            "activities":     act_progress,
        }

    found = sum(1 for v in out.values() if v["progress_pct"] > 0)
    print(f"    Progress % read: {len(out)} roads, {found} with progress > 0")
    return out


# ══════════════════════════════════════════════════════════════════════════════
# STEP 5 — EXPORT PROGRESS JSON  (for HTML map)
# ══════════════════════════════════════════════════════════════════════════════

def export_progress_json(all_updates: dict, all_progress_pct: dict,
                         output_path: Path):
    """
    Save iris_progress_data.json for the HTML map.

    Primary metric: progress_pct = IPO / OCW * 100  (from formula rows)
    Secondary: per-activity executed_km (capped at contract)

    JSON structure:
    {
      "generated": "2026-04-15T18:05:00Z",
      "projects": {
        "163.22": {
          "total_progress_pct": 12.3,
          "roads": {
            "SO-202": {
              "road_name": "Muyuni A...",
              "island": "UNGUJA",
              "contract_km": 3.668,
              "progress_pct": 78.4,         ← IPO / OCW × 100
              "ipo_pct": 0.045,             ← raw formula value
              "ocw_pct": 0.057,             ← raw formula value
              "activities": {
                "Asphalt":   {"executed_km": 2.1, "contract_km": 3.668,
                             "progress_pct": 85.3, "ipo_pct": ..., "ocw_pct": ...},
                ...
              }
            }
          }
        }
      }
    }
    """
    from datetime import timezone
    out = {
        "generated": datetime.now(timezone.utc).isoformat(),
        "projects":  {},
    }

    for project, updates in all_updates.items():
        proj_pct  = all_progress_pct.get(project, {})

        # Build per-road activity dict from km updates
        roads_km: dict[str, dict] = defaultdict(lambda: {"activities": {}})
        for u in updates:
            rc = u["road_code"]
            roads_km[rc]["road_name"]   = u["road_name"]
            roads_km[rc]["island"]      = u["island"]
            roads_km[rc]["activities"][u["pp_activity"]] = {
                "executed_km": u["executed_km"],
                "contract_km": u["contract_km"],
                "write_value": u["write_value"],
            }

        # Merge with progress % from formula rows
        roads_out: dict[str, dict] = {}
        all_codes = set(roads_km.keys()) | set(proj_pct.keys())
        for rc in all_codes:
            pct_info = proj_pct.get(rc, {})
            km_info  = roads_km.get(rc, {"activities": {}})
            # Try code variants for pct lookup
            if not pct_info:
                for v in _code_variants(rc):
                    if v in proj_pct:
                        pct_info = proj_pct[v]
                        break

            # For each activity in roads_km, augment with progress_pct from pct_info
            activities_out = {}
            for act_id, act_km in km_info.get("activities", {}).items():
                act_pct_info = pct_info.get("activities", {}).get(act_id, {})
                activities_out[act_id] = {
                    **act_km,
                    "progress_pct": act_pct_info.get("progress_pct", 0.0),
                    "ipo_pct":      act_pct_info.get("ipo_pct", 0.0),
                    "ocw_pct":      act_pct_info.get("ocw_pct", 0.0),
                }

            roads_out[rc] = {
                "road_name":    pct_info.get("road_name") or km_info.get("road_name", rc),
                "island":       pct_info.get("island")    or km_info.get("island", ""),
                "contract_km":  pct_info.get("contract_km", 0.0),
                "progress_pct": pct_info.get("progress_pct", 0.0),
                "ipo_pct":      pct_info.get("ipo_pct", 0.0),
                "ocw_pct":      pct_info.get("ocw_pct", 0.0),
                "activities":   activities_out,
            }

        # Overall project progress = sum(IPO) / sum(OCW) * 100
        total_ipo = sum(v["ipo_pct"] for v in roads_out.values() if v["ipo_pct"] > 0)
        total_ocw = sum(v["ocw_pct"] for v in roads_out.values() if v["ocw_pct"] > 0)
        overall_pct = round(min(100.0, (total_ipo / total_ocw) * 100), 2) if total_ocw > 0 else 0.0

        out["projects"][project] = {
            "total_progress_pct": overall_pct,
            "roads": roads_out,
        }

    output_path.write_text(json.dumps(out, indent=2), encoding="utf-8")
    print(f"\n  Progress JSON saved → {output_path}")


# ══════════════════════════════════════════════════════════════════════════════
# MAIN
# ══════════════════════════════════════════════════════════════════════════════

_today_override: Optional[date] = None


def main():
    global _today_override

    parser = argparse.ArgumentParser(description="Update Physical Progress Google Sheets from IRIS RFI data")
    parser.add_argument("--dry-run",   action="store_true", help="Print changes without writing")
    parser.add_argument("--date",      default=None,        help="Override today (YYYY-MM-DD)")
    parser.add_argument("--project",   default="both",      choices=["65.5", "163.22", "both"])
    parser.add_argument("--list-tabs", action="store_true", help="List all sheet tabs in RFI & Progress spreadsheets, then exit")
    parser.add_argument("--diagnose",  action="store_true", help="Dump PP sheet header rows to find IPO/OCW columns, then exit")
    args = parser.parse_args()

    # ── Diagnostic mode ───────────────────────────────────────────────────────
    if args.list_tabs:
        print("\n[DIAGNOSTIC] Listing all spreadsheet tabs …\n")
        svc = auth()
        all_ids = {
            "RFI 65.5":          RFI_ID["65.5"],
            "RFI 163.22":        RFI_ID["163.22"],
            "RFI Combined":      (BASE_DIR / "rfi_spreadsheet_id.txt").read_text().strip(),
            "PhysProgress 65.5": PP_ID["65.5"],
            "PhysProgress 163":  PP_ID["163.22"],
        }
        for label, ssid in all_ids.items():
            try:
                meta = svc.spreadsheets().get(spreadsheetId=ssid).execute()
                tabs = [s["properties"]["title"] for s in meta.get("sheets", [])]
                print(f"  {label} ({ssid[:20]}…):")
                for t in tabs:
                    print(f"    • {t}")
            except Exception as e:
                print(f"  {label}: ERROR — {e}")
        sys.exit(0)

    # ── Diagnose mode ─────────────────────────────────────────────────────────
    if args.diagnose:
        print("\n[DIAGNOSE] Dumping Physical Progress header rows …\n")
        svc = auth()
        for proj in (["65.5", "163.22"] if args.project == "both" else [args.project]):
            ssid = PP_ID[proj]
            print(f"\n  === {proj} KM  (sheet id: {ssid[:24]}…) ===")
            try:
                meta = svc.spreadsheets().get(spreadsheetId=ssid).execute()
                pp_tab = next((s["properties"]["title"] for s in meta["sheets"]
                               if "physical progress" in s["properties"]["title"].lower()),
                              meta["sheets"][0]["properties"]["title"])
                result = svc.spreadsheets().values().get(
                    spreadsheetId=ssid,
                    range=f"'{pp_tab}'!A:BZ",
                ).execute()
                rows = result.get("values", [])
                print(f"  Sheet: {pp_tab!r}   Rows returned: {len(rows)}")
                for ri, row in enumerate(rows[:8]):
                    print(f"\n  Row {ri+1} ({len(row)} cols):")
                    for ci, cell in enumerate(row):
                        if str(cell).strip():
                            print(f"    Col {col_letter(ci)} (idx {ci}): {str(cell)[:80]!r}")
            except Exception as e:
                print(f"  ERROR: {e}")
        sys.exit(0)

    if args.date:
        _today_override = datetime.strptime(args.date, "%Y-%m-%d").date()
        print(f"Date override: {_today_override}")

    print("=" * 70)
    print("  IRIS Physical Progress Updater")
    print(f"  Run at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"  Mode: {'DRY RUN' if args.dry_run else 'LIVE WRITE'}")
    print("=" * 70)

    # Authenticate
    print("\n[1] Authenticating with Google …")
    service = auth()
    print("    ✓ Authenticated")

    projects = ["65.5", "163.22"] if args.project == "both" else [args.project]
    all_updates      = {}
    all_progress_pct = {}
    total_written    = 0
    summary_lines    = []

    for project in projects:
        print(f"\n{'─'*60}")
        print(f"  PROJECT: {project} KM")
        print(f"{'─'*60}")

        # Read IRIS RFI executed lengths
        print("\n[A] Reading IRIS RFI executed lengths …")
        rfi_data = read_rfi_data(service, project)
        print(f"    {len(rfi_data)} (road, island, activity) combinations loaded")

        # Parse Physical Progress sheet structure
        print("\n[B] Parsing Physical Progress sheet …")
        pp_data = parse_pp_sheet(service, project)
        if not pp_data:
            print(f"    ERROR: Could not parse Physical Progress for {project} KM")
            continue

        # Compute updates
        print("\n[C] Computing progress values …")
        updates = compute_updates(pp_data, rfi_data)
        changed = [u for u in updates if u["changed"]]
        print(f"    {len(updates)} activity cells analysed, {len(changed)} will be updated")

        # Print full summary
        print(f"\n[D] {'Writing' if not args.dry_run else 'Simulating'} updates …")
        n = write_updates(service, pp_data["ssid"], pp_data["sheet_title"], updates, args.dry_run)
        total_written += n
        all_updates[project] = updates

        # Read back formula progress rows (IPO / OCW)
        print(f"\n[E] Reading formula progress rows …")
        pct_data = read_progress_pct(service, pp_data["ssid"], pp_data["sheet_title"], pp_data)
        all_progress_pct[project] = pct_data

        # Print top roads by progress
        if pct_data:
            top = sorted(pct_data.items(), key=lambda x: -x[1]["progress_pct"])[:5]
            print(f"    Top 5 roads by progress:")
            for rc, info in top:
                print(f"      {rc:12s} {info['road_name'][:35]:37s} {info['progress_pct']:6.1f}%  "
                      f"(IPO={info['ipo_pct']:.4f}% / OCW={info['ocw_pct']:.4f}%)")

            # Count roads where IPO/OCW were successfully read
            roads_with_pct = sum(1 for v in pct_data.values() if v["ipo_pct"] > 0 or v["ocw_pct"] > 0)
            print(f"    Roads with IPO/OCW data: {roads_with_pct}/{len(pct_data)}")
            if pp_data.get("ipo_col") is not None and pp_data.get("ocw_col") is not None:
                print(f"    IPO column: {col_letter(pp_data['ipo_col'])}, OCW column: {col_letter(pp_data['ocw_col'])}")

        # Overall progress for this project
        t_ipo = sum(v["ipo_pct"] for v in pct_data.values())
        t_ocw = sum(v["ocw_pct"] for v in pct_data.values())
        proj_pct = round((t_ipo / t_ocw) * 100, 2) if t_ocw > 0 else 0.0
        print(f"    Overall {project} KM progress: {proj_pct:.1f}%")

        # Build summary
        summary_lines.append(
            f"  {project} KM: {len(pp_data['roads'])} roads, "
            f"{len(updates)} activities, {n} cells {'written' if not args.dry_run else 'to write'}, "
            f"overall progress = {proj_pct:.1f}%"
        )

    # Export JSON for HTML map
    json_out = BASE_DIR / "iris_progress_data.json"
    print(f"\n[5] Exporting progress JSON …")
    export_progress_json(all_updates, all_progress_pct, json_out)

    print("\n" + "=" * 70)
    print("  SUMMARY")
    print("=" * 70)
    for line in summary_lines:
        print(line)
    print(f"\n  Total cells updated: {total_written}")
    print(f"  Progress JSON: {json_out}")
    print(f"\n  ✓ Done at {datetime.now().strftime('%H:%M:%S')}")


if __name__ == "__main__":
    main()
