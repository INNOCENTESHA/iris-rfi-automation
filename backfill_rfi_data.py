"""
backfill_rfi_data.py
====================
Scans all ZIP and RAR archives in the RFI-ZIPs folder and fills in
missing Attachment Path values in the four RFI data TXT files.

What it does:
  1. Reads all four TXT files:
       65.5 PEMBA RFI_DATA.txt
       65.5 UNGUJA RFI_DATA.txt
       163.22 PEMBA RFI_DATA.txt
       163.22 UNGUJA RFI_DATA.txt
  2. Identifies rows where the Attachment Path column is empty.
  3. Builds an index of every file inside every ZIP and RAR in RFI-ZIPs
     WITHOUT extracting (read-only inventory).
  4. For each row with a missing path, searches the index for a PDF whose
     name contains the RFI number (e.g. IR-Q-RFI-P-25-0001).
  5. Writes updated TXT files back (tab-separated, same format).
  6. Prints a summary of how many rows were filled vs. still missing.

Run once:
    python C:\\AI-Projects\\backfill_rfi_data.py

NOTE: This script never extracts ZIP or RAR contents — it only lists them.
"""

import os
import sys
import re
import zipfile
from pathlib import Path

# ── Paths ─────────────────────────────────────────────────────────────────────
RFI_ZIPS_FOLDER = r"D:\HIGHWAY ENGINEERING CAREER\RFI-ZIPs"

# Uploaded TXT files (place copies in C:\AI-Projects\rfi_data\ or adjust paths)
# The script tries multiple locations.
_TXT_CANDIDATES = [
    r"D:\HIGHWAY ENGINEERING CAREER",
    r"C:\AI-Projects\rfi_data",
    r"C:\AI-Projects",
]

TXT_FILES = [
    "65.5 PEMBA RFI_DATA.txt",
    "65.5 UNGUJA RFI_DATA.txt",
    "163.22 PEMBA RFI_DATA.txt",
    "163.22 UNGUJA RFI_DATA.txt",
]

# Column indices (0-based) in the tab-separated TXT files
COL_SERIAL       = 0
COL_RFI_NO       = 1
COL_DATE_RECV    = 2
COL_INSP_DATE    = 3
COL_ROAD_CODE    = 4
COL_ROAD_NAME    = 5
COL_ISLAND       = 6
COL_CONTRACT     = 7
COL_START_STA    = 8
COL_END_STA      = 9
COL_LENGTH       = 10
COL_MERGED_LEN   = 11
COL_ACTIVITY     = 12
COL_WORK_DESC    = 13
COL_STA_RANGE    = 14
COL_STATUS       = 15
COL_RE_REMARKS   = 16
COL_REMARK_DATE  = 17
COL_ATTACH_PATH  = 18
COL_PROJECT      = 19

TOTAL_COLS = 20


# ─────────────────────────────────────────────────────────────────────────────
# Auto-install rarfile if available
# ─────────────────────────────────────────────────────────────────────────────

def _ensure_rarfile():
    try:
        import rarfile  # noqa
    except ImportError:
        try:
            import subprocess as _sp
            _sp.check_call(
                [sys.executable, "-m", "pip", "install", "rarfile", "-q"],
                stdout=_sp.DEVNULL, stderr=_sp.DEVNULL,
            )
        except Exception:
            pass


_ensure_rarfile()


# ─────────────────────────────────────────────────────────────────────────────
# Build archive index: {archive_basename: [member_filename, ...]}
# ─────────────────────────────────────────────────────────────────────────────

def build_archive_index(folder: str) -> dict:
    """
    Scan folder for .zip and .rar files.
    Returns {archive_filename: [member_path_inside_archive, ...]}.
    Never extracts anything.
    """
    index = {}
    if not os.path.isdir(folder):
        print(f"  WARNING: RFI-ZIPs folder not found: {folder}")
        return index

    entries = sorted(os.scandir(folder), key=lambda e: e.name)
    zips = [e for e in entries if e.name.lower().endswith(".zip") and e.is_file()]
    rars = [e for e in entries if e.name.lower().endswith(".rar") and e.is_file()]

    print(f"  Found {len(zips)} ZIP(s) and {len(rars)} RAR(s) in {folder}")

    # Index ZIPs
    for entry in zips:
        try:
            with zipfile.ZipFile(entry.path, "r") as zf:
                members = [m.filename for m in zf.infolist() if not m.is_dir()]
                index[entry.name] = members
        except zipfile.BadZipFile:
            print(f"    WARNING: Bad ZIP (skipped): {entry.name}")
        except Exception as e:
            print(f"    WARNING: Could not read ZIP '{entry.name}': {e}")

    # Index RARs
    for entry in rars:
        try:
            import rarfile
            with rarfile.RarFile(entry.path, "r") as rf:
                members = [m.filename for m in rf.infolist() if not m.is_dir()]
                index[entry.name] = members
        except ImportError:
            # rarfile not available — list via unrar subprocess if present
            members = _list_rar_subprocess(entry.path)
            if members:
                index[entry.name] = members
        except Exception as e:
            print(f"    WARNING: Could not read RAR '{entry.name}': {e}")

    total_files = sum(len(v) for v in index.values())
    print(f"  Archive index built: {len(index)} archives, {total_files} total file entries")
    return index


def _list_rar_subprocess(rar_path: str) -> list:
    """Fallback: list RAR contents using unrar/7z command."""
    import shutil as _sh
    import subprocess as _sp
    candidates = [
        r"C:\Program Files\WinRAR\UnRAR.exe",
        r"C:\Program Files (x86)\WinRAR\UnRAR.exe",
        r"C:\Program Files\7-Zip\7z.exe",
    ]
    tool = None
    for c in candidates:
        if os.path.exists(c):
            tool = c
            break
    if not tool:
        for t in ("unrar", "7z"):
            found = _sh.which(t)
            if found:
                tool = found
                break
    if not tool:
        return []

    try:
        if "7z" in tool.lower():
            cmd = [tool, "l", rar_path]
        else:
            cmd = [tool, "lb", rar_path]
        result = _sp.run(cmd, capture_output=True, text=True, timeout=30)
        lines = result.stdout.splitlines()
        # For unrar 'lb', each line is a filename
        return [ln.strip() for ln in lines if ln.strip() and not ln.startswith("---")]
    except Exception:
        return []


# ─────────────────────────────────────────────────────────────────────────────
# Build a fast lookup:  normalised_rfi_number → (archive_name, member_path)
# ─────────────────────────────────────────────────────────────────────────────

def _normalise_rfi(rfi_str: str) -> str:
    """
    Return a canonical lowercase key for matching.
    Examples:
      IR-Q-RFI-P-25-0001   → ir-q-rfi-p-25-0001
      IR-QAQC-RFI-P-25-0001 → ir-qaqc-rfi-p-25-0001
      IRIS-QAQC-RFI-P-25-0007 → iris-qaqc-rfi-p-25-0007
    """
    return rfi_str.strip().lower()


def build_rfi_lookup(archive_index: dict) -> dict:
    """
    Build {normalised_rfi_number: [(archive_name, member_filename), ...]}.
    Extracts the RFI number from each member filename (e.g. "IR-Q-RFI-P-25-0001.pdf").
    """
    # Pattern: any variant of RFI number
    _RFI_PATTERN = re.compile(
        r"(IR(?:IS)?-(?:QAQC-)?RFI-[PU]-\d{2}-\d{4}(?:-\d+)?)",
        re.IGNORECASE,
    )

    lookup: dict = {}
    for archive_name, members in archive_index.items():
        for member_path in members:
            member_basename = os.path.basename(member_path.replace("\\", "/"))
            match = _RFI_PATTERN.search(member_basename)
            if match:
                rfi_key = _normalise_rfi(match.group(1))
                if rfi_key not in lookup:
                    lookup[rfi_key] = []
                lookup[rfi_key].append((archive_name, member_path))

    print(f"  RFI lookup index: {len(lookup)} unique RFI numbers found in archives")
    return lookup


# ─────────────────────────────────────────────────────────────────────────────
# Find TXT files
# ─────────────────────────────────────────────────────────────────────────────

def find_txt_file(filename: str) -> str | None:
    for folder in _TXT_CANDIDATES:
        candidate = os.path.join(folder, filename)
        if os.path.exists(candidate):
            return candidate
    return None


# ─────────────────────────────────────────────────────────────────────────────
# Parse and update a single TXT file
# ─────────────────────────────────────────────────────────────────────────────

def parse_txt(filepath: str) -> list[list[str]]:
    """Read tab-separated TXT file, return list of rows (each row is list of cells)."""
    rows = []
    with open(filepath, "r", encoding="utf-8", errors="replace") as f:
        for line in f:
            cells = line.rstrip("\n").split("\t")
            rows.append(cells)
    return rows


def pad_row(cells: list[str], target_len: int = TOTAL_COLS) -> list[str]:
    """Ensure a row has exactly target_len cells."""
    while len(cells) < target_len:
        cells.append("")
    return cells[:target_len]


def process_txt_file(filepath: str, rfi_lookup: dict) -> tuple[int, int]:
    """
    Read the TXT file, fill in missing Attachment Paths using rfi_lookup,
    and write the updated file back.
    Returns (filled_count, still_missing_count).
    """
    rows = parse_txt(filepath)
    filled = 0
    still_missing = 0

    for row_idx, cells in enumerate(rows):
        # Skip header rows and blank rows
        if not cells or not cells[0].strip():
            continue
        if cells[COL_RFI_NO if len(cells) > COL_RFI_NO else 0].strip() in ("RFI No.", ""):
            continue

        cells = pad_row(cells)
        rfi_no  = cells[COL_RFI_NO].strip()
        attach  = cells[COL_ATTACH_PATH].strip()

        if not rfi_no or rfi_no == "RFI No.":
            continue

        # Only fill rows that have an RFI number but no attachment path
        if attach:
            continue

        rfi_key = _normalise_rfi(rfi_no)
        hits    = rfi_lookup.get(rfi_key, [])

        if hits:
            # Use the first match; build a Windows-style virtual path
            archive_name, member_path = hits[0]
            # member_path may use / or \ separators; normalise to \
            member_basename = member_path.replace("/", "\\")
            if not member_basename.startswith("\\"):
                pass  # keep as-is
            # Build full path: folder\archive.zip\member.pdf
            virtual_path = os.path.join(RFI_ZIPS_FOLDER, archive_name, member_basename)
            cells[COL_ATTACH_PATH] = virtual_path
            rows[row_idx] = cells
            filled += 1
        else:
            still_missing += 1

    # Write back
    with open(filepath, "w", encoding="utf-8", newline="") as f:
        for cells in rows:
            f.write("\t".join(cells) + "\n")

    return filled, still_missing


# ─────────────────────────────────────────────────────────────────────────────
# Main
# ─────────────────────────────────────────────────────────────────────────────

def main():
    print()
    print("=" * 60)
    print("  IRIS — RFI Data Backfill")
    print("=" * 60)
    print()

    # 1. Build archive index (no extraction)
    print("Step 1: Indexing archives in RFI-ZIPs folder ...")
    archive_index = build_archive_index(RFI_ZIPS_FOLDER)
    if not archive_index:
        print("  WARNING: No archives found. Nothing to match against.")

    print()
    print("Step 2: Building RFI number lookup ...")
    rfi_lookup = build_rfi_lookup(archive_index)

    print()
    print("Step 3: Processing TXT files ...")
    print()

    grand_filled  = 0
    grand_missing = 0

    for txt_name in TXT_FILES:
        filepath = find_txt_file(txt_name)
        if not filepath:
            print(f"  NOT FOUND: {txt_name}")
            print(f"    (Looked in: {', '.join(_TXT_CANDIDATES)})")
            print(f"    Copy the file to one of those folders and re-run.")
            continue

        print(f"  Processing: {txt_name}")
        print(f"    Path: {filepath}")

        try:
            filled, missing = process_txt_file(filepath, rfi_lookup)
            grand_filled  += filled
            grand_missing += missing
            print(f"    Filled: {filled}  |  Still missing: {missing}")
        except Exception as e:
            print(f"    ERROR: {e}")

        print()

    print("=" * 60)
    print(f"  SUMMARY: {grand_filled} attachment path(s) filled in.")
    print(f"           {grand_missing} row(s) still have no matching archive entry.")
    if grand_missing > 0:
        print()
        print("  Tip: Rows still missing are RFIs whose PDF files have not been")
        print("  downloaded yet. Run download_rfi_attachments.py to fetch them,")
        print("  then run this script again.")
    print("=" * 60)
    print()


if __name__ == "__main__":
    main()
