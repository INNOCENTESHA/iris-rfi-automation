#!/usr/bin/env python3
"""
update_physical_progress.py
===========================

Reads the manually-maintained RFI_PIVOT tabs in each project's RFI workbook,
extracts the cumulative "Length Executed (km)" per (road code, activity), and
writes those figures — plus a per-road summation — into the "PHYSICAL PROGRESS"
tab of each project's Physical_Progress workbook, in the current reporting
column.

Projects
--------
163.22  snapshot day = 19  (columns dated the 19th of each month)
65.5    snapshot day = 15  (columns dated the 15th of each month)

Target column rule (per user)
-----------------------------
Write into the snapshot column whose date is the FIRST one >= today.
Equivalently: once today is past this month's snapshot day (19th / 15th),
roll forward to next month's column.

Safety
------
DRY-RUN by default. It logs every cell it WOULD write and changes nothing.
Set environment variable  PHYS_PROGRESS_LIVE=1  to actually write.
Always do one dry-run and eyeball the log before enabling live writes — the
row/activity matching is heuristic and should be verified against the real
sheet layout the first time.

Auth
----
Reuses the same Google OAuth token the RFI sheet script uses:
  GOOGLE_TOKEN_PATH  (default ~/.iris/token.json)
  GOOGLE_CREDS_PATH  (default ~/.iris/credentials.json)
"""

from __future__ import annotations

import os
import re
import sys
import datetime as _dt
from typing import Dict, List, Tuple, Optional

from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build

# --------------------------------------------------------------------------- #
# Configuration
# --------------------------------------------------------------------------- #

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

PROJECTS = {
    "163.22": {
        "rfi_id": "1IjN4W79VEZrAjevG5zZVKyJPuCdGbm_xAbcttTNjfOA",
        "pp_id":  "1YbDCzePl76CqtO9dLl2YZjIZtVmJKSgVkmU6LKqXKUc",
        "snapshot_day": 19,
        "islands": ["UNGUJA", "PEMBA"],
    },
    "65.5": {
        "rfi_id": "1WV58zo-jYgqOTBGOtMAjjlHC3WKdQ83s-XdRFlxZLLU",
        "pp_id":  "16RE08Ivv2nSltqxVk1X9G8cBYcZP82kIGhzGAwRtZ7s",
        "snapshot_day": 15,
        "islands": ["UNGUJA", "PEMBA"],
    },
}

PP_TAB = "PHYSICAL PROGRESS"
LIVE = os.environ.get("PHYS_PROGRESS_LIVE", "0") == "1"

# Road-code pattern, e.g. NA-200, CE-201, SO-202, MK-233, WB-216, UR-223 …
ROAD_CODE_RE = re.compile(r"^[A-Z]{2,3}-?\d{3}$")


# --------------------------------------------------------------------------- #
# Google auth
# --------------------------------------------------------------------------- #

def get_service():
    token_path = os.environ.get("GOOGLE_TOKEN_PATH",
                                os.path.expanduser("~/.iris/token.json"))
    creds_path = os.environ.get("GOOGLE_CREDS_PATH",
                                os.path.expanduser("~/.iris/credentials.json"))
    creds = None
    if os.path.exists(token_path):
        creds = Credentials.from_authorized_user_file(token_path, SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            raise SystemExit(
                "No valid Google token. In CI, refresh token.json locally and "
                "update the GOOGLE_TOKEN_B64 secret."
            )
    return build("sheets", "v4", credentials=creds, cache_discovery=False)


# --------------------------------------------------------------------------- #
# Small helpers
# --------------------------------------------------------------------------- #

def _col_letter(idx0: int) -> str:
    """0-based column index -> A1 letter."""
    s, n = "", idx0 + 1
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def _to_float(v) -> Optional[float]:
    if v is None:
        return None
    s = str(v).strip().replace(",", "")
    if s == "" or s == "-":
        return None
    try:
        return float(s)
    except ValueError:
        return None


def _norm(s: str) -> str:
    """Normalise an activity name for matching."""
    s = str(s).upper()
    s = s.replace("—", "-").replace("–", "-")
    s = re.sub(r"[^A-Z0-9]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def get_values(service, ssid: str, tab: str) -> List[List[str]]:
    rng = f"'{tab}'"
    resp = service.spreadsheets().values().get(
        spreadsheetId=ssid, range=rng,
        valueRenderOption="UNFORMATTED_VALUE",
    ).execute()
    return resp.get("values", [])


# --------------------------------------------------------------------------- #
# Pivot parsing
# --------------------------------------------------------------------------- #

def parse_pivot(rows: List[List[str]]) -> Dict[Tuple[str, str], float]:
    """
    Parse one RFI_PIVOT tab.

    Returns {(road_code, activity_norm): cumulative_length}.

    Layout: the pivot is a set of activity blocks laid out horizontally.
    Each block's header row contains a "... Road Code" cell at the block's
    first column and an "Activity Name <ACTIVITY>" cell at the value column.
    Per-road totals appear on rows whose first cell is "<CODE> Total".
    """
    out: Dict[Tuple[str, str], float] = {}
    if not rows:
        return out
    header = rows[0]

    # Identify block value columns: header cell starting with "Activity Name"
    # plus subgrade layer / grand-total columns.
    value_cols: Dict[int, str] = {}      # col_idx -> activity label
    code_col_for: Dict[int, int] = {}    # value_col -> its block's road-code col
    last_code_col = 0
    for c, cell in enumerate(header):
        t = str(cell or "").strip()
        if "road code" in t.lower():
            last_code_col = c
        if t.lower().startswith("activity name"):
            value_cols[c] = t[len("Activity Name"):].strip()
            code_col_for[c] = last_code_col
        elif t.lower() in ("grand total",) or "layer" in t.lower():
            value_cols[c] = t
            code_col_for[c] = last_code_col

    # Walk rows; capture "<CODE> Total" lines.
    for row in rows[1:]:
        for vcol, activity in value_cols.items():
            ccol = code_col_for.get(vcol, 0)
            if ccol >= len(row):
                continue
            label = str(row[ccol] or "").strip()
            m = re.match(r"^(.*?)\s+Total$", label, re.IGNORECASE)
            if not m:
                continue
            code = m.group(1).strip().upper()
            if not code:
                continue
            val = _to_float(row[vcol]) if vcol < len(row) else None
            if val is None:
                continue
            key = (code, _norm(activity))
            out[key] = out.get(key, 0.0) + val
    return out


# --------------------------------------------------------------------------- #
# PHYSICAL PROGRESS sheet handling
# --------------------------------------------------------------------------- #

def find_date_header(rows: List[List[str]]) -> Tuple[int, Dict[int, _dt.date]]:
    """Return (header_row_index, {col_idx: date}) for the dated snapshot row."""
    best_row, best_map = -1, {}
    for ri, row in enumerate(rows[:12]):
        dmap: Dict[int, _dt.date] = {}
        for ci, cell in enumerate(row):
            d = _parse_date(cell)
            if d:
                dmap[ci] = d
        if len(dmap) > len(best_map):
            best_map, best_row = dmap, ri
    return best_row, best_map


def _parse_date(cell) -> Optional[_dt.date]:
    s = str(cell or "").strip()
    for fmt in ("%d-%b-%y", "%d-%b-%Y", "%d/%m/%Y", "%Y-%m-%d"):
        try:
            return _dt.datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def pick_target_col(date_map: Dict[int, _dt.date], today: _dt.date) -> Optional[int]:
    """First snapshot column whose date is >= today; else the latest column."""
    future = sorted(((d, c) for c, d in date_map.items() if d >= today))
    if future:
        return future[0][1]
    if date_map:
        return max(date_map.items(), key=lambda kv: kv[1])[0]
    return None


def find_road_code_col(rows: List[List[str]], header_row: int) -> int:
    """Column index most populated with road-code-looking values."""
    counts: Dict[int, int] = {}
    for row in rows[header_row + 1:]:
        for ci, cell in enumerate(row):
            if ROAD_CODE_RE.match(str(cell or "").strip().upper()):
                counts[ci] = counts.get(ci, 0) + 1
    return max(counts, key=counts.get) if counts else 3  # default col D


def build_pp_plan(rows, code_col, target_col, pivot, today):
    """
    Build a list of (row_index_0based, value, note) writes.

    For every road (a row carrying a road code) we write, into target_col:
      * each matching activity sub-row's cumulative value, and
      * the road-level summation (sum of that road's activity values) on the
        road row itself.
    """
    writes = []
    n = len(rows)
    # Locate road rows.
    road_rows = [i for i, r in enumerate(rows)
                 if code_col < len(r)
                 and ROAD_CODE_RE.match(str(r[code_col] or "").strip().upper())]
    for idx, ri in enumerate(road_rows):
        code = str(rows[ri][code_col]).strip().upper()
        end = road_rows[idx + 1] if idx + 1 < len(road_rows) else n
        road_total = 0.0
        matched_any = False
        # Activity sub-rows are the rows between this road row and the next.
        for rj in range(ri, end):
            row = rows[rj]
            row_text = " ".join(_norm(c) for c in row if c)
            # match against this road's pivot activities
            for (pcode, pact), pval in pivot.items():
                if pcode != code:
                    continue
                if pact and pact in row_text:
                    writes.append((rj, target_col, pval,
                                   f"{code} / {pact} = {pval}"))
                    road_total += pval
                    matched_any = True
        if matched_any:
            # Road-level summation written on the road row.
            writes.append((ri, target_col, round(road_total, 3),
                           f"{code} ROAD TOTAL = {round(road_total,3)}"))
    return writes


# --------------------------------------------------------------------------- #
# Main
# --------------------------------------------------------------------------- #

def run():
    today = _dt.date.today()
    try:
        service = get_service()
    except Exception as e:
        print(f"  SKIP: Google auth unavailable ({e}).")
        print("  Fix: the GOOGLE_TOKEN_B64 secret must be a base64 of a VALID token.json.")
        return
    mode = "LIVE WRITE" if LIVE else "DRY-RUN (no changes)"
    print(f"\n=== update_physical_progress | {today} | {mode} ===")

    for km, cfg in PROJECTS.items():
        print(f"\n--- Project {km} KM ---")

        # 1) Gather cumulative per (road_code, activity) from all island pivots.
        pivot: Dict[Tuple[str, str], float] = {}
        for island in cfg["islands"]:
            tab = f"{km} {island} RFI_PIVOT"
            try:
                rows = get_values(service, cfg["rfi_id"], tab)
                part = parse_pivot(rows)
                pivot.update(part)
                print(f"  pivot '{tab}': {len(part)} (road,activity) totals")
            except Exception as e:
                print(f"  WARN: could not read pivot '{tab}': {e}")

        if not pivot:
            print("  No pivot data — skipping project.")
            continue

        # 2) Read PHYSICAL PROGRESS and locate the write target.
        try:
            pp = get_values(service, cfg["pp_id"], PP_TAB)
        except Exception as e:
            print(f"  ERROR: cannot read '{PP_TAB}': {e}")
            continue

        hrow, dmap = find_date_header(pp)
        target = pick_target_col(dmap, today)
        if target is None:
            print("  ERROR: no dated snapshot columns found — skipping.")
            continue
        code_col = find_road_code_col(pp, hrow)
        tdate = dmap.get(target)
        print(f"  header row={hrow+1}  road-code col={_col_letter(code_col)}  "
              f"target col={_col_letter(target)} ({tdate})")

        # 3) Build the write plan.
        writes = build_pp_plan(pp, code_col, target, pivot, today)
        print(f"  {len(writes)} cells to update:")
        for (rj, cj, val, note) in writes[:200]:
            print(f"    {_col_letter(cj)}{rj+1:<4} <- {val:<10} [{note}]")

        # 4) Apply (or not).
        if not writes:
            print("  Nothing matched — verify activity names line up.")
            continue
        if not LIVE:
            print("  DRY-RUN: no cells written. Set PHYS_PROGRESS_LIVE=1 to apply.")
            continue
        data = [{"range": f"'{PP_TAB}'!{_col_letter(cj)}{rj+1}",
                 "values": [[val]]} for (rj, cj, val, _n) in writes]
        service.spreadsheets().values().batchUpdate(
            spreadsheetId=cfg["pp_id"],
            body={"valueInputOption": "USER_ENTERED", "data": data},
        ).execute()
        print(f"  ✓ wrote {len(data)} cells to {PP_TAB}.")

    print("\n=== done ===")


if __name__ == "__main__":
    run()
