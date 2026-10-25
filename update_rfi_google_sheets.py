"""
IRIS RFI → Google Sheets Uploader & Daily Updater  (v5 — April 2026)
======================================================================
Reads ALL documents found in the folder named by RFI_ZIP_DIR —
both inside ZIP archives and standalone files — extracts RFI metadata,
derives canonical Activity Names from Work Descriptions, handles multi-
segment chainage ranges, then INSERTS new records ABOVE THE FOOTER ROW
of pre-created named Google Sheets Tables in the existing RFI Detailed
Log spreadsheets.

TARGET NAMED TABLES (must already exist — created manually by the user):
  In "65.5 IRIS RFI Detailed Log":
      KM 65.5 UNGUJA SITE RFI_DATA    KM 65.5 UNGUJA LAB RFI_DATA
      KM 65.5 PEMBA  SITE RFI_DATA    KM 65.5 PEMBA  LAB RFI_DATA

  In "163.22 IRIS RFI Detailed Log":
      KM 163.22 UNGUJA SITE RFI_DATA  KM 163.22 UNGUJA LAB RFI_DATA
      KM 163.22 PEMBA  SITE RFI_DATA  KM 163.22 PEMBA  LAB RFI_DATA

ISLAND DETECTION (from RFI No. — primary rule):
  "P" in the RFI number → Pemba Island
  "U" in the RFI number → Unguja Island
  e.g.  IRIS-RFI-P-26-001 → Pemba     IR-QAQC-RFI-U-26-001 → Unguja

SITE / LAB ROUTING  (RFI No. always wins over ZIP filename):
  QAQC in RFI No. OR in PDF/ZIP name → LAB table
  Q (isolated) in RFI No.            → SITE table
  A QAQC document zipped inside a SITE archive is still routed to LAB.

TABLE INSERTION RULES:
  • New rows are inserted ABOVE the table footer (totals row).
  • No entry is ever placed outside the named table range.
  • Orphan entries found below a table are rescued: deleted from their
    stray position and re-queued for proper insertion above the footer.
  • S/No. in each table counts 1, 2, 3 … from the first data row.

STATUS DROPDOWN (Google Sheets data-validation applied automatically):
  APPROVED | APPROVED WITH COMMENTS | NOT APPROVED |
  CANCELLED | OPENED | CLOSED

FORMULA FIELDS (written as live Sheets formulas):
  • Length Executed (km)  — calculated from Start/End Station columns
  • Contract Number       — IF formula driven by Project column

ROAD NAMES:
  Canonical road names looked up from the Physical Progress sheets.
  Names extracted from RFI PDFs have any leading chainage stripped.
  Road names are written as they appear in the Physical Progress sheets.

APPEND LOGIC (no data loss):
  Reads existing RFI Nos. in each table + scans for orphan rows below
  the table.  Only records whose RFI No. is not already present are
  inserted.  S/No. continues from the current maximum in each table.

CONFIGURATION (environment variables — nothing is hard-coded):
    RFI_ZIP_DIR         required  folder holding the RFI archives / documents
    GOOGLE_CREDS_PATH   required  path to the OAuth client-secret JSON
    GOOGLE_TOKEN_PATH   required  path to the authorised OAuth token JSON
    N8N_CALLBACK_URL    optional  webhook notified about run status

Install dependencies with:  pip install -r requirements.txt
Run with:                   python update_rfi_google_sheets.py
"""

import os, re, io, sys, json, zipfile, datetime, time, shutil, subprocess, tempfile
from pathlib import Path
from collections import defaultdict

# ── Output / Input paths ──────────────────────────────────────────────────────
ZIP_FOLDER = os.environ.get("RFI_ZIP_DIR", "").strip()
if not ZIP_FOLDER:
    raise SystemExit(
        "CONFIGURATION ERROR: set RFI_ZIP_DIR to the folder holding the RFI archives."
    )

# ── Valid RFI filename patterns ────────────────────────────────────────────────
# Only PDFs/docs matching these formats are logged to Google Sheets.
# IRIS-[QAQC|Q]-RFI-[U|P]-##-####[.pdf]  or  IR-… variant
# e.g. IRIS-QAQC-RFI-U-25-0001.pdf, IR-Q-RFI-P-25-0042-1.pdf
_VALID_RFI_RE = re.compile(
    # Prefix match only (no $ anchor) — any trailing suffix (_Stamped, -1, -Rev2, etc.) is allowed.
    # Separators between tokens are collapsed to a single "-" before matching.
    r"^(?:IRIS|IR)-?(?:QAQC|Q)-?RFI-?[UP]-?\d{2}-?\d{3,4}",
    re.IGNORECASE,
)

# Status-summary PDFs — must NOT be logged as individual RFIs.
# They ARE parsed to retrieve per-RFI status/remarks for existing sheet rows.
_STATUS_PDF_PATTERNS = [
    re.compile(r"65\.?5\s*KM.*RFI.*status", re.IGNORECASE),
    re.compile(r"163\.?22\s*KM.*RFI.*status", re.IGNORECASE),
    re.compile(r"RFI.*status.*65\.?5", re.IGNORECASE),
    re.compile(r"RFI.*status.*163\.?22", re.IGNORECASE),
]


def is_valid_rfi_file(filename: str) -> bool:
    """Return True if the filename starts with the official RFI prefix (any suffix allowed)."""
    stem = Path(filename).stem
    # Collapse all runs of dashes, underscores, and spaces into a single dash before matching.
    # e.g. "IRIS- QAQC-RFI-P-25-0009" → "IRIS-QAQC-RFI-P-25-0009"  ✓
    #      "IRIS-Q-RFI-U-25-0028_Stamped"  → "IRIS-Q-RFI-U-25-0028-Stamped"  ✓
    stem_norm = re.sub(r"[-_\s]+", "-", stem.strip())
    return bool(_VALID_RFI_RE.match(stem_norm))


def is_status_pdf(filename: str) -> bool:
    """Return True for the project-level status summary PDFs (skip as entries)."""
    name = os.path.basename(filename)
    return any(p.search(name) for p in _STATUS_PDF_PATTERNS)


# ── WinRAR / unrar tool discovery ─────────────────────────────────────────────
def _find_rar_tool() -> str | None:
    """Find WinRAR, unrar, or 7-Zip on this machine."""
    import shutil as _sh
    candidates = [
        r"C:\Program Files\WinRAR\UnRAR.exe",
        r"C:\Program Files\WinRAR\WinRAR.exe",
        r"C:\Program Files (x86)\WinRAR\UnRAR.exe",
        r"C:\Program Files (x86)\WinRAR\WinRAR.exe",
        r"C:\Program Files\7-Zip\7z.exe",
    ]
    for c in candidates:
        if os.path.exists(c):
            return c
    for name in ("unrar", "unar", "rar", "7z", "bsdtar"):
        found = _sh.which(name)
        if found:
            return found
    return None

_RAR_TOOL = None   # cached after first lookup


def _open_rar(rar_path: str):
    """
    Context manager that yields a list of (member_name, data_bytes) tuples
    from a RAR archive, using rarfile library with auto-configured tool path.
    Falls back to subprocess extraction into a temp dir.
    """
    global _RAR_TOOL
    if _RAR_TOOL is None:
        _RAR_TOOL = _find_rar_tool() or ""

    # Strategy 1: rarfile library
    try:
        import rarfile
        if _RAR_TOOL:
            rarfile.UNRAR_TOOL = _RAR_TOOL
        with rarfile.RarFile(rar_path, "r") as rf:
            members = []
            for info in rf.infolist():
                if info.is_dir():
                    continue
                try:
                    data = rf.read(info)
                    members.append((info.filename, data))
                except Exception:
                    pass
            return members
    except Exception as e1:
        pass

    # Strategy 2: subprocess into temp dir, then read files
    if _RAR_TOOL:
        tmp = tempfile.mkdtemp()
        try:
            tool = _RAR_TOOL.lower()
            if "7z" in tool:
                cmd = [_RAR_TOOL, "e", rar_path, f"-o{tmp}", "-y", "-aoa"]
            elif "winrar" in tool:
                cmd = [_RAR_TOOL, "e", "-y", "-ibck", rar_path, tmp + "\\"]
            else:
                cmd = [_RAR_TOOL, "e", "-y", rar_path, tmp]
            subprocess.run(cmd, capture_output=True, timeout=120)
            members = []
            for root, _, files in os.walk(tmp):
                for fname in files:
                    fp = os.path.join(root, fname)
                    with open(fp, "rb") as fh:
                        members.append((fname, fh.read()))
            return members
        except Exception as e2:
            print(f"  WARNING: RAR subprocess failed: {e2}")
        finally:
            shutil.rmtree(tmp, ignore_errors=True)

    print(f"  WARNING: Cannot open '{os.path.basename(rar_path)}' — "
          "install WinRAR (https://www.rarlab.com) and add to PATH.")
    return []

# ── Dependency check ─────────────────────────────────────────────────────────
# Dependencies are declared in requirements.txt. Installing packages at runtime
# hides real problems and is unsafe in CI, so we only report what is missing.
import importlib

_REQUIRED_MODULES = {
    "gspread": "gspread",
    "google.oauth2": "google-auth",
    "google_auth_oauthlib": "google-auth-oauthlib",
    "googleapiclient": "google-api-python-client",
    "pdfplumber": "pdfplumber",
    "docx": "python-docx",
}

_missing_modules = []
for _module, _package in _REQUIRED_MODULES.items():
    try:
        importlib.import_module(_module)
    except ImportError:
        _missing_modules.append(_package)

if _missing_modules:
    raise SystemExit(
        "CONFIGURATION ERROR: missing dependencies: "
        + ", ".join(sorted(_missing_modules))
        + ". Install them with: pip install -r requirements.txt"
    )

import gspread
from google.oauth2.credentials import Credentials
from google.oauth2 import service_account
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import pdfplumber


# ══════════════════════════════════════════════════════════════════════════════
# CONFIGURATION
# ══════════════════════════════════════════════════════════════════════════════
CREDENTIALS_FILE = os.environ.get("GOOGLE_CREDS_PATH", "").strip()
TOKEN_FILE = os.environ.get("GOOGLE_TOKEN_PATH", "").strip()
SERVICE_ACCOUNT_FILE = os.environ.get("GOOGLE_SERVICE_ACCOUNT_PATH", "").strip()

# Either a service account (preferred for unattended/cloud runs -- it never
# expires) or the OAuth client + token pair (interactive, local Windows runs).
if not SERVICE_ACCOUNT_FILE:
    for _label, _value in (("GOOGLE_CREDS_PATH", CREDENTIALS_FILE),
                           ("GOOGLE_TOKEN_PATH", TOKEN_FILE)):
        if not _value:
            raise SystemExit(
                f"CONFIGURATION ERROR: {_label} is not set (or set "
                "GOOGLE_SERVICE_ACCOUNT_PATH to use a service account)."
            )

# ── RFI Detailed Log spreadsheet names ────────────────────────────────────────
SPREADSHEET_655  = "65.5 IRIS RFI Detailed Log"
SPREADSHEET_163  = "163.22 IRIS RFI Detailed Log"

# ── Physical Progress spreadsheet IDs (READ-ONLY — for road name lookup) ──────
from project_sheets import PP_ID, PP_SHEET_NAME, RFI_ID_BY_NAME, physical_progress_tab

PHYS_PROGRESS_ID_655 = PP_ID["65.5"]
PHYS_PROGRESS_ID_163 = PP_ID["163.22"]

# ── Canonical Road Code → Road Name lookup ────────────────────────────────────
# Built from all IRIS RFI records; chainage stripped.
# Used as primary fallback when Physical Progress sheets are unreachable,
# and applied during every audit pass to clean existing sheet entries.
ROAD_CODE_TO_NAME: dict[str, str] = {
    "AR-227":    "Access Road to District sport Centers",
    "CE-100":    "Ubago - Jumbi (AL2)",
    "CE-100.":   "Ubago - Jumbi (AL1)",
    "CE-102":    "Kianga - Istiqama (AL1)",
    "CE-102.":   "Kianga - Istiqama (AL2)",
    "CE-201":    "Istiqaama-Zanzibar University",
    "CE-213":    "Access Road to JKU Dunga camp including car parking",
    "CE-219":    "Jumbi makonter (Meli tisa) and spur",
    "CH-106":    "Ole Mihogoni - Uhanga - Amani",
    "CH-107":    "Pujini - Mwalimu Nyerere",
    "CH-241":    "Dutani JCT - Dutani",
    "CH-248":    "Chanjaani-Mitamani-Pujini",
    "CH-252":    "Pujini(Kilimo)-Mkandandume _Kibaridi",
    "MI-240":    "Kiuyu Mbuyuni_Bandari(Micheweni Access Road)",
    "MI-240-A":  "Micheweni Access Roads Alignment 1A (Kiuyu - Bandarini)",
    "MI-240-A.": "Micheweni Access Roads Alignment 1B (Bandarini - KMKM)",
    "MI-240-D":  "Mjini kiuyu msituni - Kiuyu road (Micheweni Access Road) AL4",
    "MI-244":    "Konde Sokoni - Kiuyu kwa Manda",
    "MK-105":    "Mtambile-Kengeja(Polisi) AL 1",
    "MK-108":    "Kwapopo - Goli la Mpira",
    "MK-109":    "Minazini-Mazrui",
    "MK-110":    "Mwambe (Mawasiliano)-Chanjaani School",
    "MK-112":    "Mazrui - Usiambini",
    "MK-113":    "Kengeja-Mwambe(Polisi) AL 2",
    "MK-233":    "Kiwani Sokoni - Mchochani",
    "MK-234":    "Kichaka-Chwaka Kubwa + Chwaka Ndogo Spar",
    "MK-235":    "Kichungwani-Kwa Sanani",
    "MK-236":    "Mauwani-Kajirambeni + Kirimdomo and Kihogoni Spar",
    "MK-237":    "Mahuduthi-Sipwese + Mto Mkuu Spar",
    "MK-238":    "Minazini-Shikazako",
    "MK-239":    "Mtamba-Sitondoo + Tomondo-Tondooni Spar",
    "MK-242":    "Mizingani - Wamba + Funfulagoni Cocolagoon Spur",
    "MK-243":    "Chambani (kwa Mbunge) - JKU - Mtangani",
    "MK-245":    "Kuyuni - Kalani-Zawadini",
    "MK-249":    "Mtambile (Police Post)-Kangaani",
    "MK-250":    "Kenya-Kwa Mtora (Chambani)",
    "NA-200":    "Matemwe - Sun Shine Hotel & Spur",
    "NB-217":    "Donge Mbiji - Karange + Spur",
    "PE-IC01":   "Gando - Wete - Konde",
    "PE-IC02":   "Ole-Mihogoni - Kangagani - Amani",
    "PE-IC03":   "Pujini-Mwalimu Nyerere",
    "SO-202":    "Muyuni A (Kibaoni)-Muyuni C (Dindi)",
    "SO-203":    "Muyuni B (Kidodoni)-Ng'ombeni + Skuli-Ndorobo Spur",
    "SO-204":    "Dindi-Mchangamle + Wasusanya-Kilimamatangi Spur",
    "SO-205":    "Kidodofirishi-Kizimachakacha",
    "SO-206":    "Kisokoni - Kaskazi (Kikwecheni) (AL1)",
    "SO-207":    "Madete Skuli - Skuli ya Kusini",
    "SO-208":    "Mnarani-Kigaeni",
    "SO-209":    "Mtendemndo-Pwani & Spur",
    "SO-210":    "Mwembekungu - Kaskazi Pemba",
    "SO-211":    "Mwembesanda-Mavua",
    "SO-212":    "Nganani - Bungi - Mtegani (AL2)",
    "SO-212.":   "Nganani - Kumbini - Chuini (AL1)",
    "SO-212..":  "Nganani - Bungi - Mtegani (AL3)",
    "SO-212\u2026": "Nganani - Bungi - Mtegani (AL4)",
    "SO-214":    "Tasani - Mnarani-Pwani",
    "SO-215":    "Wadini-Cheju",
    "UR-223":    "Ziwani Police Access Roads",
    "UR-224":    "Kilimani Prison Headquarters Access Roads",
    "WA-103":    "Mbuzini - Kijinchi (AL1)",
    "WA-103.":   "Mbuzini - Kijinchi (AL2)",
    "WA-103..":  "Mbuzini - Kijinchi (AL3)",
    "WA-104":    "Konde - Wete",
    "WA-104.":   "Wete - Gando",
    "WA-105":    "Konde - Wete",
    "WA-11":     "Raha - Ukunjwi Road",
    "WA-111":    "Raha - Ukunjwi Road",
    "WA-221":    "Pili Space (Dole) - Shamba la Bishafika Feeder Road",
    "WA-222":    "Flawa - Suleiman Kondo Mskitini Access Road",
    "WA-226":    "Access Road to Mtoni JWTZ Camp",
    "WA-256":    "Kama Msikitini",
    "WB-101":    "Langoni - Kwabiasha (AL1)",
    "WB-101.":   "Langoni - Kwabiasha (AL2)",
    "WB-101..":  "Langoni - Kwabiasha (AL3)",
    "WB-220":    "Kisauni Mazurui - Tangi la Maji Feeder Road and Spur",
    "WB-229":    "Mitondooni Bakery Access Road",
    "WB-230":    "Kisauni Mtundani Feeder Road",
    "WB-232":    "Maungani Vijijini - Shakani Feeder Road and Spur",
    "WB-255":    "Shakani-Nyamanzi",
    "WB-257":    "Kwakibunju-Maungani-Bondeni-Kisakasaka",
    "WE-246":    "Mianzini-Mabatini (Gando)",
    "WE-251":    "Wete Access Road",
}

# ── Contract numbers ───────────────────────────────────────────────────────────
CONTRACT_655 = "SMZ/P01/W/DP/2024-2025/05"   # 65.5 KM
CONTRACT_163 = "SMZ/P01/W/DP/2025-2026/30"   # 163.22 KM

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

# ── Named table name builder ───────────────────────────────────────────────────
def build_table_name(km_key: str, island: str, rfi_type: str) -> str:
    """e.g.  build_table_name('163.22', 'Unguja', 'SITE') → 'KM 163.22 UNGUJA SITE RFI_DATA'"""
    return f"KM {km_key} {island.upper()} {rfi_type.upper()} RFI_DATA"

# ── Collaborator email lists ───────────────────────────────────────────────────
EMAILS_655 = [
    "innoptesha@gmail.com",
    "egbertgosbert47@gmail.com",
    "annahmanumbu@gmail.com",
    "metachazy@gmail.com",
    "amiryhalifahida@gmail.com",
    "romatypeter09@gmail.com",
    "keddylamar96@gmail.com",
    "nyatir@yahoo.com",
    "msyl1952@gmail.com",
    "jessicamoses220@gmail.com",
    "pmanyele@gmail.com",
    "jumajafaris79@gmail.com",
]

EMAILS_163 = [
    "innoptesha@gmail.com",
    "andrewisaya@yahoo.com",
    "egbertgosbert47@gmail.com",
    "annahmanumbu@gmail.com",
    "metachazy@gmail.com",
    "kelvinmmbaga@gmail.com",
    "amiryhalifahida@gmail.com",
    "jamalyabdul360@gmail.com",
    "azadjamal03@gmail.com",
    "keddylamar96@gmail.com",
    "notckelmbonde1@gmail.com",
    "hsamadu2000@gmail.com",
    "jessicamoses220@gmail.com",
    "pmanyele@gmail.com",
    "cartimber5@gmail.com",
    "kabelegedenice02@gmail.com",
    "hurumabosco408@gmail.com",
    "jumajafaris79@gmail.com",
    "saidasaidi11@icloud.com",
]


# ══════════════════════════════════════════════════════════════════════════════
# ACTIVITY NORMALISATION
# ══════════════════════════════════════════════════════════════════════════════
ACTIVITY_PATTERNS = [
    # SITE CLEARANCE
    (r"site\s*clear|clearing\s*(?:of\s+)?site|vegetation\s*clear|bush\s*clear"
     r"|tree\s*(?:cut|remov)|demolish|demolition|strip\s*topsoil",
     "SITE CLEARANCE"),

    # ROADBED PREPARATION
    (r"roadbed|road\s*bed|formation\s*(?:level|grading)|shaping\s*(?:of\s+)?road"
     r"|profile\s*grading",
     "ROADBED PREPARATION"),

    # SUBGRADE — stratified layers (ordinals must not be alignment numbers like AL1)
    (r"(?:sub[\s\-]?grade|subgrade\s*lay|compaction).*?"
     r"(?:\b1st\b|first\s+layer|layer\s*[:\-]?\s*1\b)"
     r"|1st\s+(?:sub[\s\-]?grade|subgrade)\s*(?:lay|compact|course)?",
     "SUBGRADE LAYING — 1st LAYER"),
    (r"(?:sub[\s\-]?grade|subgrade\s*lay|compaction).*?"
     r"(?:\b2nd\b|second\s+layer|layer\s*[:\-]?\s*2\b)"
     r"|2nd\s+(?:sub[\s\-]?grade|subgrade)\s*(?:lay|compact|course)?",
     "SUBGRADE LAYING — 2nd LAYER"),
    (r"(?:sub[\s\-]?grade|subgrade\s*lay|compaction).*?"
     r"(?:\b3rd\b|third\s+layer|layer\s*[:\-]?\s*3\b)"
     r"|3rd\s+(?:sub[\s\-]?grade|subgrade)\s*(?:lay|compact|course)?",
     "SUBGRADE LAYING — 3rd LAYER"),
    (r"(?:sub[\s\-]?grade|subgrade\s*lay|compaction).*?"
     r"(?:\b4th\b|fourth\s+layer|layer\s*[:\-]?\s*4\b)"
     r"|4th\s+(?:sub[\s\-]?grade|subgrade)\s*(?:lay|compact|course)?",
     "SUBGRADE LAYING — 4th LAYER"),
    (r"(?:sub[\s\-]?grade|subgrade\s*lay|compaction).*?"
     r"(?:\b5th\b|fifth\s+layer|layer\s*[:\-]?\s*5\b)"
     r"|5th\s+(?:sub[\s\-]?grade|subgrade)\s*(?:lay|compact|course)?",
     "SUBGRADE LAYING — 5th LAYER"),
    (r"(?:sub[\s\-]?grade|subgrade\s*lay|compaction).*?"
     r"(?:\b6th\b|sixth\s+layer|layer\s*[:\-]?\s*6\b)"
     r"|6th\s+(?:sub[\s\-]?grade|subgrade)\s*(?:lay|compact|course)?",
     "SUBGRADE LAYING — 6th LAYER"),
    (r"(?:sub[\s\-]?grade|subgrade\s*lay|compaction).*?\b(\d+)(?:st|nd|rd|th)\s+layer"
     r"|\b(\d+)(?:st|nd|rd|th)\s+(?:sub[\s\-]?grade|subgrade)\s*(?:lay|compact)?",
     "SUBGRADE LAYING — nth LAYER"),
    # Generic subgrade / excavation / general filling
    (r"sub[\s\-]?grade|subgrade\s*lay|compaction\s*(?:of\s+)?subgrade"
     r"|excavat|general\s*fill|earthwork|embankment|cut\s*and\s*fill"
     r"|filling\s*(?:and\s*)?compaction",
     "SUBGRADE LAYING"),

    # SUBBASE LAYING
    (r"sub[\s\-]?base\s*(?:lay|cours|material|compaction)?",
     "SUBBASE LAYING"),

    # BASECOURSE
    (r"base[\s\-]?course|base\s*lay|granular\s*(?:base|sub)",
     "BASECOURSE"),

    # PRIME COAT
    (r"prime\s*coat|bituminous\s*prime|priming",
     "PRIME COAT"),

    # TACK COAT
    (r"tack\s*coat|bond\s*coat|bituminous\s*tack",
     "TACK COAT"),

    # ASPHALT LAYING
    (r"asphalt|bituminous\s*(?:concrete|layer|wearing|binder)"
     r"|AC\s*(?:lay|wearing|binder)|wearing\s*course|binder\s*course"
     r"|surface\s*dressing|hot\s*mix",
     "ASPHALT LAYING"),

    # ROAD MARKING
    (r"road\s*mark|line\s*mark|thermoplastic\s*mark|paint(?:ing)?\s*(?:of\s+)?road"
     r"|lane\s*mark|centre\s*line",
     "ROAD MARKING"),

    # TRAFFIC SIGNS INSTALLATION
    (r"traffic\s*sign|road\s*sign|signage\s*install|erect(?:ion)?\s*(?:of\s+)?sign",
     "TRAFFIC SIGNS INSTALLATION"),

    # CULVERT
    (r"culvert|pipe\s*(?:drain|culvert)|headwall|wingwall|inlet|outlet\s*structure",
     "CULVERT"),

    # SERVICE DUCTS
    (r"service\s*duct|utility\s*duct|conduit|duct\s*install",
     "SERVICE DUCTS"),

    # GUARD RAIL REPARATION
    (r"guard[\s\-]?rail|guardrail|crash\s*barrier|safety\s*barrier|railing",
     "GUARD RAIL REPARATION"),

    # STRUCTURE REPARATION
    (r"structure\s*rep|bridge\s*rep|retaining\s*wall\s*rep|repair.*?structure"
     r"|structural\s*rep|abutment\s*rep",
     "STRUCTURE REPARATION"),

    # POTHOLE REPARATION
    (r"pothole|patching|patch\s*(?:work|repair)|filling\s*(?:of\s+)?pothole",
     "POTHOLE REPARATION"),

    # OVERLAY
    (r"overlay|re[\s\-]?surfac|mill(?:ing)?\s*and\s*(?:re)?pave|rehabilitation\s*overlay",
     "OVERLAY"),

    # SIDE DRAINS CLEARANCE (must precede SIDE DRAINS)
    (r"side\s*drain\s*clear|drain(?:age)?\s*clear|clear(?:ing)?\s*(?:of\s+)?drain"
     r"|de[\s\-]?silt|cleaning\s*(?:of\s+)?drain",
     "SIDE DRAINS CLEARANCE"),

    # SIDE DRAINS
    (r"side\s*drain|kerb\s*drain|edge\s*drain|drainage\s*channel"
     r"|longitudinal\s*drain|u[\s\-]?drain|open\s*drain\s*construct",
     "SIDE DRAINS"),

    # MIX DESIGNS AND MATERIAL SAMPLING
    (r"sampling|mix\s*design|material\s*(?:test|samp)|lab(?:oratory)?\s*test"
     r"|quality\s*(?:test|control)|trial\s*mix|marshal(?:l)?|proctor"
     r"|sieve|gradation|cbr|atterberg|moisture\s*content|density\s*test"
     r"|core\s*sample|field\s*test",
     "MIX DESIGNS AND MATERIAL SAMPLING"),

    # LOW WATER CROSSING  (LWC in work description → "LOW WATER CROSSING")
    (r"\blwc\b|low[\s\-]?water[\s\-]?cross|low\s*level\s*cross|drift\s*cross",
     "LOW WATER CROSSING"),

    # DRIFT
    (r"\bdrift\b|drift\s*(?:construct|install|repair|replac)",
     "DRIFT"),

    # JOINT / EXPANSION JOINTS
    (r"expansion\s*joint|bridge\s*joint|joint\s*seal|movement\s*joint",
     "JOINT"),

    # ENVIRONMENTAL / EROSION CONTROL
    (r"erosion\s*control|gabion|rip[\s\-]?rap|slope\s*protect|berm\s*protect"
     r"|geotextile|geo[\s\-]?fabric",
     "EROSION CONTROL"),

    # REINSTATEMENT
    (r"reinstat|restore\s*(?:road|surface)|temporary\s*(?:road|access)",
     "REINSTATEMENT"),
]

# Collect ALL canonical activity names for the dropdown (preserves insertion order,
# deduplicated).
_SEEN: set = set()
ACTIVITY_DROPDOWN: list = []
for _pat, _name in ACTIVITY_PATTERNS:
    if _name not in _SEEN:
        _SEEN.add(_name)
        ACTIVITY_DROPDOWN.append(_name)
ACTIVITY_DROPDOWN.append("OTHER")   # fallback always last
del _SEEN, _pat, _name

# Patterns to check FIRST when the chainage is a single point (start == end).
SINGLE_POINT_STRUCTURE_PATTERNS = [
    (r"culvert|pipe\s*(?:drain|culvert)|headwall|wingwall|inlet|outlet", "CULVERT"),
    (r"service\s*duct|utility\s*duct|conduit", "SERVICE DUCTS"),
    (r"\blwc\b|low[\s\-]?water[\s\-]?cross|low\s*level\s*cross|drift\s*cross",
     "LOW WATER CROSSING"),
    (r"\bdrift\b", "DRIFT"),
    (r"ford|causeway", "LOW WATER CROSSING"),
    (r"structure|bridge|abutment|retaining\s*wall|foundation|footing", "STRUCTURE REPARATION"),
    (r"guard[\s\-]?rail|guardrail|barrier", "GUARD RAIL REPARATION"),
]

_MERGE_KEYWORDS = ("SUBGRADE", "EXCAVAT", "GENERAL FILL", "EARTHWORK",
                   "EMBANKMENT", "FILLING COMPACTION", "CUT AND FILL")


def normalise_activity(work_description: str, is_single_point: bool = False) -> str:
    if not work_description:
        return "OTHER"
    wd = work_description.strip()

    if is_single_point:
        for pattern, canonical in SINGLE_POINT_STRUCTURE_PATTERNS:
            if re.search(pattern, wd, re.IGNORECASE):
                return canonical
        if re.search(r'excavat|general\s*fill|lean\s*concret|formation\s*level', wd, re.IGNORECASE):
            return "CULVERT"
        for pattern, canonical in ACTIVITY_PATTERNS:
            if canonical.startswith("SUBGRADE"):
                continue
            if re.search(pattern, wd, re.IGNORECASE):
                return canonical
        return "OTHER"

    for pattern, canonical in ACTIVITY_PATTERNS:
        if re.search(pattern, wd, re.IGNORECASE):
            return canonical
    return "OTHER"


def is_mergeable_activity(activity_name: str) -> bool:
    a = activity_name.upper()
    return any(kw in a for kw in _MERGE_KEYWORDS)


# ══════════════════════════════════════════════════════════════════════════════
# STATION / CHAINAGE HELPERS
# ══════════════════════════════════════════════════════════════════════════════

def parse_chainage(s: str) -> float:
    if not s:
        return 0.0
    s = str(s).strip().replace(" ", "").replace(",", "")
    if "+" in s:
        parts = s.split("+", 1)
        try:
            return float(parts[0]) * 1000 + float(parts[1])
        except ValueError:
            pass
    try:
        return float(s)
    except ValueError:
        return 0.0


def metres_to_chainage(m: float) -> str:
    km  = int(m // 1000)
    rem = round(m % 1000)
    return f"{km}+{rem:03d}"


def parse_chainage_ranges(text: str) -> list:
    if not text:
        return []
    tokens = re.findall(r'\d+\+\d+', str(text))
    if not tokens:
        return []
    if len(tokens) == 1:
        m = parse_chainage(tokens[0])
        return [(m, m)]
    pairs = []
    for i in range(0, len(tokens) - 1, 2):
        s = parse_chainage(tokens[i])
        e = parse_chainage(tokens[i + 1])
        pairs.append((s, e))
    if len(tokens) % 2 == 1:
        m = parse_chainage(tokens[-1])
        pairs.append((m, m))
    return pairs


def total_length_km(ranges: list) -> float:
    total = sum(abs(e - s) for s, e in ranges)
    return round(total / 1000, 4)


def merge_intervals(intervals: list) -> list:
    if not intervals:
        return []
    ivs = sorted((min(a, b), max(a, b)) for a, b in intervals if a != b)
    if not ivs:
        return []
    merged = [ivs[0]]
    for s, e in ivs[1:]:
        if s <= merged[-1][1]:
            merged[-1] = (merged[-1][0], max(merged[-1][1], e))
        else:
            merged.append((s, e))
    return merged


def merged_length_km(intervals: list):
    m = merge_intervals(intervals)
    if not m:
        return None
    return round(sum(e - s for s, e in m) / 1000, 4)


# ══════════════════════════════════════════════════════════════════════════════
# TEXT EXTRACTION
# ══════════════════════════════════════════════════════════════════════════════

def extract_text_pdf(data: bytes) -> str:
    try:
        with pdfplumber.open(io.BytesIO(data)) as pdf:
            return "\n".join(p.extract_text() or "" for p in pdf.pages)
    except Exception:
        return ""


def extract_text_docx(data: bytes) -> str:
    try:
        from docx import Document
        doc = Document(io.BytesIO(data))
        parts = [p.text for p in doc.paragraphs]
        for table in doc.tables:
            for row in table.rows:
                parts.append(" | ".join(c.text for c in row.cells))
        return "\n".join(parts)
    except Exception:
        return ""


# ══════════════════════════════════════════════════════════════════════════════
# FIELD EXTRACTION HELPERS
# ══════════════════════════════════════════════════════════════════════════════

def _ex(pattern, text, flags=re.IGNORECASE):
    m = re.search(pattern, text, flags)
    return m.group(1).strip() if m else ""


def _clean(v, maxlen=400):
    if not v:
        return ""
    return re.sub(r"\s+", " ", str(v)).strip()[:maxlen]


def get_contract_number(rfi_no: str, project_km: str) -> str:
    rfi_upper = rfi_no.upper()
    if "IRIS" in rfi_upper or "163" in project_km:
        return CONTRACT_163
    if "IR" in rfi_upper or "65" in project_km:
        return CONTRACT_655
    return ""


# ══════════════════════════════════════════════════════════════════════════════
# STATUS CANONICALISATION
# Normalises any extracted text to one of the six dropdown values.
# ══════════════════════════════════════════════════════════════════════════════

_STATUS_MAP = [
    (r"approved\s+for\s+(?:further\s+work|construction)", "APPROVED"),
    (r"approved\s+with\s+comment",                        "APPROVED WITH COMMENTS"),
    (r"not\s+approved",                                   "NOT APPROVED"),
    (r"\bcancel",                                         "CANCELLED"),
    (r"\bclos(?:ed|e)\b",                                 "CLOSED"),
    (r"\bapproved\b",                                     "APPROVED"),
]

_STATUS_DROPDOWN = [
    "APPROVED",
    "APPROVED WITH COMMENTS",
    "NOT APPROVED",
    "CANCELLED",
    "OPENED",
    "CLOSED",
]


def canonicalize_status(raw: str) -> str:
    """Map any extracted status text to one of the six canonical values."""
    if not raw or not raw.strip():
        return "OPENED"
    for pattern, canonical in _STATUS_MAP:
        if re.search(pattern, raw, re.IGNORECASE):
            return canonical
    return "OPENED"


# ══════════════════════════════════════════════════════════════════════════════
# SITE / LAB ROUTING  (RFI No. is authoritative; ZIP filename is fallback only)
# ══════════════════════════════════════════════════════════════════════════════

def get_rfi_type(rfi_no: str, fallback_type: str,
                 zip_name: str = "", pdf_name: str = "",
                 text_snippet: str = "") -> str:
    """
    Determine SITE or LAB from ALL available signals.

    QAQC in ANY of: RFI No., ZIP name, PDF name, or first 500 chars of
    document text → LAB (regardless of which ZIP folder the doc came from).
    Isolated Q in RFI No. → SITE.
    Falls back to fallback_type from the ZIP filename if no signal found.
    """
    combined = " ".join([rfi_no, zip_name, pdf_name, text_snippet]).upper()
    if "QAQC" in combined:
        return "LAB"
    rfi_upper = rfi_no.upper()
    if re.search(r'[-/_]Q[-/_\d]|[-/_]Q$', rfi_upper):
        return "SITE"
    if re.search(r'\bQ\b', rfi_upper):
        return "SITE"
    return fallback_type


# ══════════════════════════════════════════════════════════════════════════════
# ISLAND DETECTION
# ══════════════════════════════════════════════════════════════════════════════

def island_from_rfi_no(rfi_no: str) -> str:
    """
    Primary island detection from RFI number.
    Rule: P in RFI No. → Pemba,  U in RFI No. → Unguja.
    """
    rfi_up = rfi_no.upper()
    m = re.search(r'[-/]([PU])[-/\d]', rfi_up)
    if m:
        return "Pemba" if m.group(1) == "P" else "Unguja"
    m = re.search(r'RFI([PU])', rfi_up)
    if m:
        return "Pemba" if m.group(1) == "P" else "Unguja"
    m = re.search(r'[-/]([PU])(\d{2})', rfi_up)
    if m:
        return "Pemba" if m.group(1) == "P" else "Unguja"
    return ""


# ══════════════════════════════════════════════════════════════════════════════
# PROJECT ASSIGNMENT — road code series, then RFI-number prefix fallback
# Rule: road codes 100-199 (XX-1##) → 65.5 KM project.
#       road codes 200-299 (XX-2##) → 163.22 KM project.
#       Non-numeric codes (N/A, location names, PE-IC02 …) → '' (unknown).
#
# When road code is unknown, use the RFI-number prefix as a reliable hint:
#   IRIS-QAQC-* (full "IRIS QAQC" prefix)  → 163.22 KM
#   IR-QAQC-*   (short "IR QAQC" prefix)   → 65.5 KM
#   IRIS-Q-*    (ambiguous short form)      → '' (need road code to decide)
# ══════════════════════════════════════════════════════════════════════════════

_RC_65_RE  = re.compile(r'^[A-Za-z]{1,3}-1\d{2}', re.IGNORECASE)
_RC_163_RE = re.compile(r'^[A-Za-z]{1,3}-2\d{2}', re.IGNORECASE)


def project_from_road_code(road_code: str) -> str:
    """
    Return project from NUMERIC road code series only.
      XX-2## → '163.22 KM'   XX-1## → '65.5 KM'
      Anything else (N/A, location names, alphanumeric codes) → ''
    """
    rc = (road_code or "").strip()
    if not rc:
        return ""
    if _RC_163_RE.match(rc):
        return "163.22 KM"
    if _RC_65_RE.match(rc):
        return "65.5 KM"
    return ""   # non-standard: defer to RFI-prefix or ZIP context


# Collapse any run of separators before prefix-matching
_PREFIX_NORM_RE = re.compile(r'[-_\s]+')


def project_from_rfi_prefix(rfi_no: str) -> str:
    """
    Secondary project hint from RFI-number prefix when road code is unknown.
      IRIS-QAQC-* → '163.22 KM'   (LAB records, always 163.22 contract)
      IR-QAQC-*   → '65.5 KM'     (LAB records, always 65.5 contract)
      IRIS-Q-* and all others → '' (ambiguous — road code must decide)
    """
    rfi_norm = _PREFIX_NORM_RE.sub('-', rfi_no.strip().upper())
    if re.match(r'^IRIS-QAQC-', rfi_norm):
        return "163.22 KM"
    if re.match(r'^IR-QAQC-', rfi_norm):
        return "65.5 KM"
    return ""


# Pre-compiled patterns for tier-3 PDF text scan
_TEXT_163_RE        = re.compile(r'163[.,]22\s*KM',  re.IGNORECASE)
_TEXT_65_RE         = re.compile(r'65[.,]5\s*KM',    re.IGNORECASE)
_TEXT_YEAR_163_RE   = re.compile(r'2025[-/]2026',    re.IGNORECASE)
_TEXT_YEAR_65_RE    = re.compile(r'2024[-/]2025',    re.IGNORECASE)
# Project-name KM value: "THE 113.50 KM ROAD(S) …" — if value > 65.5 → 163.22 KM
_TEXT_KM_PROJECT_RE = re.compile(
    r'\bTHE\s+(\d+(?:[.,]\d+)?)\s*KM\b', re.IGNORECASE
)


def project_from_text(text: str) -> str:
    """
    Tier-3 fallback: scan raw PDF/DOCX text for explicit project-size,
    contract-year, or project-name KM-value mentions.
      '163.22 KM' / '163,22 KM'  → '163.22 KM'
      '65.5 KM'   / '65,5 KM'    → '65.5 KM'
      '2025-2026'                 → '163.22 KM'
      '2024-2025'                 → '65.5 KM'
      'THE X KM ROAD(S)…' where X > 65.5  → '163.22 KM'
      'THE X KM ROAD(S)…' where X ≤ 65.5  → '65.5 KM'
    Explicit size > year > project-name KM value.
    """
    if not text:
        return ""
    if _TEXT_163_RE.search(text):
        return "163.22 KM"
    if _TEXT_65_RE.search(text):
        return "65.5 KM"
    if _TEXT_YEAR_163_RE.search(text):
        return "163.22 KM"
    if _TEXT_YEAR_65_RE.search(text):
        return "65.5 KM"
    # Project-name KM value rule (user-defined):
    # if KM value > 65.5 → 163.22 KM project; else → 65.5 KM project
    m = _TEXT_KM_PROJECT_RE.search(text)
    if m:
        try:
            km = float(m.group(1).replace(",", "."))
            return "163.22 KM" if km > 65.5 else "65.5 KM"
        except (ValueError, IndexError):
            pass
    return ""


# ── Tier-0 override: manually verified RFI→project for scanned-image PDFs ───
# These 6 early RFIs are scanned images — pdfplumber extracts no text, so
# tiers 1-3 all return ''.  Read visually: project header states
# "113.50 KM ROADS IN UNGUJA AND PEMBA" → 113.50 > 65.5 → 163.22 KM.
_KNOWN_RFI_PROJECTS: dict = {
    "IRIS-Q-RFI-P-25-0001": "163.22 KM",
    "IRIS-Q-RFI-P-25-0004": "163.22 KM",
    "IRIS-Q-RFI-U-25-0001": "163.22 KM",
    "IRIS-Q-RFI-U-25-0002": "163.22 KM",
    "IRIS-Q-RFI-U-25-0003": "163.22 KM",
    "IRIS-Q-RFI-U-25-0005": "163.22 KM",
}


# ══════════════════════════════════════════════════════════════════════════════
# ROAD NAME CLEANING
# ══════════════════════════════════════════════════════════════════════════════

def clean_road_name(raw: str) -> str:
    """
    Remove ALL chainage notation from a road name — leading, trailing, or
    embedded in parentheses.  Handles every common format found on IRIS RFIs:

      '0+000-5+500 Mwanakwerekwe – Bubujiko road'
      'Ch: 0+000 to 5+500 Mwanakwerekwe road'
      'STA 1+100 – 3+200 Pemba ring road'
      'Mwanakwerekwe road (0+000-5+500)'
      'Mwanakwerekwe road 0+000-5+500'
      '65.5 KM — Mwanakwerekwe – Bubujiko road'

    All become e.g. → 'Mwanakwerekwe – Bubujiko road'
    """
    if not raw:
        return ""
    s = raw.strip()

    # ── 1. Strip optional "Ch:", "STA", "Chainage:" label at the very start ───
    s = re.sub(r'^\s*(?:Ch\.?|STA\.?|Chainage)[:\s]+', '', s, flags=re.IGNORECASE)

    # ── 2. Strip a leading km-project marker like "65.5 KM —" or "163.22 KM -"
    #       \u2013 = en-dash, \u2014 = em-dash
    s = re.sub(r'^\s*[\d\.]+\s*KM\s*[-\u2013\u2014]?\s*', '', s, flags=re.IGNORECASE)

    # ── 2b. Remove any stray leading dash/en-dash/em-dash left after the above ──
    s = re.sub(r'^[-\u2013\u2014\s]+', '', s)

    # ── 3. Strip a leading chainage range (with any separator: - – to ,) ────────
    #       Matches: 0+000-5+500  |  0+000 – 5+500  |  0+000 to 5+500  |  0+000
    s = re.sub(
        r'^\s*\d+\+\d+\s*(?:[-\u2013,]|to)?\s*(?:\d+\+\d+)?\s*',
        '', s, flags=re.IGNORECASE
    )

    # ── 4. Strip trailing chainage in parentheses: "Road name (0+000-5+500)" ───
    s = re.sub(
        r'\s*[\(\[]\s*\d+\+\d+\s*(?:[-\u2013,]|to)?\s*(?:\d+\+\d+)?\s*[\)\]]\s*$',
        '', s, flags=re.IGNORECASE
    )

    # ── 5. Strip trailing bare chainage: "Road name 0+000-5+500" ────────────────
    s = re.sub(
        r'\s+\d+\+\d+\s*(?:[-\u2013,]|to)?\s*(?:\d+\+\d+)?\s*$',
        '', s, flags=re.IGNORECASE
    )

    # ── 5b. Strip "CHAINAGE: …" keyword-style suffix (any content after it) ─────
    #        e.g. "Ubago - Jumbi (AL2) CHAINAGE: 4+260 - 4+320"
    #             "Raha - Ukunjwi Road CHAINAGE: N/A"
    s = re.sub(r'\s+CHAINAGE\s*:.*$', '', s, flags=re.IGNORECASE).strip()

    # ── 6. Normalise whitespace ──────────────────────────────────────────────────
    result = re.sub(r'\s+', ' ', s).strip()
    return result if result else raw.strip()


# ══════════════════════════════════════════════════════════════════════════════
# ROAD DATA LOOKUP  (from Physical Progress Google Sheets)
# ══════════════════════════════════════════════════════════════════════════════

def load_road_data(service) -> dict:
    """
    Read canonical Road Code → Road Name from the Physical Progress sheets.
    Returns: {(km_key, road_code): road_name}   road_code is upper-cased.
    """
    lookup = {}
    sources = [
        (PHYS_PROGRESS_ID_655, "65.5"),
        (PHYS_PROGRESS_ID_163, "163.22"),
    ]

    for ssid, km_key in sources:
        print(f"\n  Loading road data for {km_key} KM from Physical Progress …")
        try:
            meta = service.spreadsheets().get(spreadsheetId=ssid).execute()
            sheet_names = [s["properties"]["title"] for s in meta.get("sheets", [])]

            candidates = [physical_progress_tab(sheet_names)]

            loaded = 0
            for sheet_name in candidates:
                try:
                    result = service.spreadsheets().values().get(
                        spreadsheetId=ssid,
                        range=f"'{sheet_name}'!A:Z"
                    ).execute()
                    rows = result.get("values", [])
                    if len(rows) < 2:
                        continue

                    header = [str(h).strip().lower() for h in rows[0]]

                    code_col = next(
                        (i for i, h in enumerate(header)
                         if h in ("road code", "code", "al") or
                         ("code" in h and "road" in " ".join(header[:6]))),
                        None
                    )
                    name_col = next(
                        (i for i, h in enumerate(header)
                         if "road name" in h or h == "name"),
                        None
                    )

                    if code_col is None or name_col is None:
                        continue

                    for row in rows[1:]:
                        code = str(row[code_col]).strip().upper() if len(row) > code_col else ""
                        name = str(row[name_col]).strip()         if len(row) > name_col else ""
                        if not code or not name:
                            continue
                        if code.lower() in ("road code", "code", "s/no.", "#", ""):
                            continue
                        # Store with cleaned name (no chainage prefix)
                        lookup[(km_key, code)] = clean_road_name(name)
                        loaded += 1

                    if loaded > 0:
                        print(f"    ✓ {loaded} road entries from sheet '{sheet_name}'")
                        break

                except Exception as e:
                    print(f"    Warning reading sheet '{sheet_name}': {e}")
                    continue

            if loaded == 0:
                print(f"    ⚠  No road data loaded for {km_key} KM "
                      f"(will use names extracted from PDFs)")

        except Exception as e:
            print(f"    ⚠  Could not access Physical Progress for {km_key} KM: {e}")

    return lookup


# ══════════════════════════════════════════════════════════════════════════════
# FIELD EXTRACTION
# ══════════════════════════════════════════════════════════════════════════════

def extract_fields(text: str, zip_name: str, pdf_name: str,
                   project: str, island_hint: str, date_hint: str,
                   source_path: str, fallback_rfi_type: str) -> dict:
    """Extract all fields from one RFI document's raw text."""
    rec = {}

    # ── RFI No. ────────────────────────────────────────────────────────────────
    rec["rfi_no"] = (
        _ex(r"RFI\s+NO[:\s]+([A-Z0-9\-\/]+)", text) or
        _ex(r"\b(IRIS[-\s]?(?:QAQC[-\s]?)?RFI[-\s]?[PU]?[-\s]?\d{2}[-\s]?\d+)\b", text) or
        _ex(r"\b(IR[-\s]?(?:QAQC[-\s]?)?RFI[-\s]?[PU]?[-\s]?\d{2}[-\s]?\d+)\b", text) or
        os.path.splitext(pdf_name)[0]
    )

    # ── SITE or LAB  — use ALL signals so a QAQC doc in a SITE ZIP goes to LAB ──
    rec["rfi_type"] = get_rfi_type(
        rec["rfi_no"], fallback_rfi_type,
        zip_name=zip_name,
        pdf_name=pdf_name,
        text_snippet=text[:600],
    )

    # ── Contract Number ────────────────────────────────────────────────────────
    rec["contract_number"] = get_contract_number(rec["rfi_no"], project)

    # ── Dates ──────────────────────────────────────────────────────────────────
    rec["date_received"] = (
        _ex(r"DATE\s+OF\s+SUBMISSION[:\s]+(\d{1,2}[/\-\.]\d{1,2}[/\-\.]\d{2,4})", text) or
        _ex(r"DATE\s+RECEIVED[:\s]+(\d{1,2}[/\-\.]\d{1,2}[/\-\.]\d{2,4})", text) or
        _ex(r"Date\s*[:\-]\s*(\d{1,2}[/\-\.]\d{1,2}[/\-\.]\d{4})", text) or
        date_hint
    )
    rec["inspection_date"] = (
        _ex(r"DATE\s+OF\s+INSPECTION[:\s]+(\d{1,2}[/\-\.]\d{1,2}[/\-\.]\d{2,4})", text) or
        _ex(r"INSPECTION\s+DATE[:\s]+(\d{1,2}[/\-\.]\d{1,2}[/\-\.]\d{2,4})", text)
    )

    # ── Road Code  (real road codes only; AL\d+ alignment indicators are NOT codes) ─
    # Priority: explicit ROAD CODE field → LOCATION field (if it looks like a code)
    # AL1/AL2/… alignment designators are stripped; fuzzy-matching during the
    # audit pass will derive the real road code from the road name.
    _rc_explicit = _ex(r'ROAD\s+CODE[:\s]+([A-Z0-9\-\/\.]{2,20})', text)
    _rc_location = _ex(r'LOCATION[:\s]+([A-Z]{1,3}[-][\d\.]+[A-Z\-]*)', text)
    road_code_raw = _rc_explicit or _rc_location or ""
    rec["road_code"] = _clean(road_code_raw).upper().strip() if road_code_raw else ""
    # Discard anything that is purely an alignment number (AL1, AL 2, AL-3 …)
    if re.match(r'^AL\s*[-]?\s*\d+$', rec["road_code"], re.IGNORECASE):
        rec["road_code"] = ""

    # ── Road Name  ────────────────────────────────────────────────────────────
    # Taken VERBATIM from the "ROAD NAME:" field of the RFI document.
    # Do NOT strip or modify — the chainage is already in a separate "CHAINAGE:"
    # field and must never be confused with the road name.
    # clean_road_name() is intentionally NOT applied here; it is only used on
    # names sourced from the Physical Progress lookup sheets (which may carry
    # spreadsheet-formatting chainage prefixes).
    _raw_road = _clean(
        _ex(r"ROAD\s+NAME[:\s]+([^\n\r]{3,120})", text) or
        _ex(r"ROAD[:\s]+([^\n\r]{3,120})", text)
        # Note: LOCATION is deliberately excluded here — it holds a section/
        # milestone reference (e.g. MI-240-B), not a road name.
    )
    # Strip any trailing "CHAINAGE: …" that may have been captured alongside
    # the road name (e.g. "Ubago - Jumbi (AL2) CHAINAGE: 4+260 - 4+320")
    rec["road_name"] = re.sub(
        r'\s+CHAINAGE\s*:.*$', '', _raw_road, flags=re.IGNORECASE
    ).strip()

    # ── Island (PRIMARY: from RFI No.; FALLBACK: document text / filename) ─────
    island_text = island_from_rfi_no(rec["rfi_no"])

    if not island_text:
        island_text = _ex(r"ISLAND[:\s]+(PEMBA|UNGUJA|ZANZIBAR)", text, re.IGNORECASE)

    if not island_text:
        fn_up = (zip_name + " " + pdf_name).upper()
        if "UNGUJA" in fn_up or "-U-" in fn_up or "_U_" in fn_up:
            island_text = "Unguja"
        elif "PEMBA" in fn_up or "-P-" in fn_up or "_P_" in fn_up:
            island_text = "Pemba"
        else:
            island_text = island_hint

    rec["island"] = island_text.capitalize() if island_text else ""

    # ── Station Range & Length ─────────────────────────────────────────────────
    chainage_raw = (
        _ex(r"CHAINAGE[:\s]+([0-9\+\.\s\-\u2013&/,]+)", text) or
        _ex(r"(?:Ch\.?|STA\.?)\s+([0-9\+\.\s\-\u2013]{4,120})", text)
    )
    rec["station_range"] = _clean(chainage_raw)

    ranges = parse_chainage_ranges(rec["station_range"])

    if ranges:
        rec["start_station"] = metres_to_chainage(ranges[0][0])
        rec["end_station"]   = metres_to_chainage(ranges[-1][1])
        computed_km          = total_length_km(ranges)
        rec["length_km"]     = computed_km if computed_km > 0 else ""
        rec["all_ranges"]    = ranges
    else:
        st = _ex(r"START\s+(?:CH|STA|CHAINAGE)[:\s]+([\d\+\.]+)", text)
        en = _ex(r"END\s+(?:CH|STA|CHAINAGE)[:\s]+([\d\+\.]+)", text)
        rec["start_station"] = st
        rec["end_station"]   = en
        s_m = parse_chainage(st)
        e_m = parse_chainage(en)
        if s_m or e_m:
            rec["length_km"]  = round(abs(e_m - s_m) / 1000, 4) or ""
            rec["all_ranges"] = [(s_m, e_m)]
        else:
            doc_len = _ex(r"LENGTH[:\s]+([\d\.,]+)\s*(?:KM|M\b)", text, re.IGNORECASE)
            try:
                v = float(doc_len.replace(",", "."))
                rec["length_km"] = v if v < 100 else round(v / 1000, 4)
            except Exception:
                rec["length_km"] = ""
            rec["all_ranges"] = []

    # ── Work Description & Activity ────────────────────────────────────────────
    wd = _ex(
        r"WORK\s+DESCRIPTION\s*[:\s]+(.+?)(?=\n\s*(?:ENCLOSURES|RELATED|Requested|$))",
        text, re.IGNORECASE | re.DOTALL
    )
    if not wd:
        wd = _ex(r"(?:WORK|WORKS|ACTIVITY)\s*[:\-]\s*(.{5,250}?)(?:\n|$)", text)
    rec["work_description"] = _clean(wd)

    rng = rec["all_ranges"]
    is_single_point = (len(rng) == 1 and rng[0][0] == rng[0][1])
    rec["activity_name"] = normalise_activity(rec["work_description"], is_single_point)

    # ── Status — canonicalised to dropdown values ──────────────────────────────
    status_m = re.search(
        r"(Approved\s+for\s+(?:Further\s+Work|Construction)"
        r"|Approved\s+with\s+Comment[s]?"
        r"|Not\s+Approved[^,\n]*"
        r"|Cancelled"
        r"|Closed"
        r"|Approved)",
        text, re.IGNORECASE
    )
    raw_status = _clean(status_m.group(1)) if status_m else ""
    rec["status"] = canonicalize_status(raw_status)

    # ── RE's Remarks ───────────────────────────────────────────────────────────
    re_rem = _ex(
        r"Comments?\s+by\s+(?:the\s+)?RE[:\s]*(.+?)(?=\nRESIDENT ENGINEER|\nRequest:|\Z)",
        text, re.IGNORECASE | re.DOTALL
    )
    if not re_rem:
        re_rem = _ex(
            r"(?:RE|RESIDENT ENGINEER)['\u2019]?s?\s+(?:REMARKS?|COMMENTS?)[:\s]*"
            r"(.+?)(?=\n\n|Date|\Z)",
            text, re.IGNORECASE | re.DOTALL
        )
    rec["re_remarks"]  = _clean(re_rem)
    rec["remark_date"] = (
        _ex(r"RE\s+SIGNATURE.*?DATE[:\s]+(\d{1,2}[/\-\.]\d{1,2}[/\-\.]\d{2,4})",
            text, re.IGNORECASE | re.DOTALL) or
        _ex(r"REMARK\s+DATE[:\s]+(\d{1,2}[/\-\.]\d{1,2}[/\-\.]\d{2,4})", text)
    )

    # ── Provenance ─────────────────────────────────────────────────────────────
    rec["attachment_path"]  = source_path
    rec["zip_file"]         = zip_name
    rec["doc_file"]         = pdf_name
    rec["merged_length_km"] = ""   # filled later by attach_merged_lengths

    # ── Project assignment (five-tier fallback) ───────────────────────────────
    # 0. Manual override for scanned-image PDFs (no extractable text)
    # 1. Road code series XX-1## → 65.5 KM  |  XX-2## → 163.22 KM  (authoritative)
    # 2. RFI-number prefix  IRIS-QAQC-* → 163.22 KM  |  IR-QAQC-* → 65.5 KM
    # 3. PDF/DOCX full-text scan (exact size / contract year / project KM value)
    # 4. ZIP/filename context  (e.g. "65.5 KM & 163.22 KM … LAB.zip")
    # A conflict between tier-1 and tier-4 flags the record as CANCELLED.
    rfi_key = rec.get("rfi_no", "").strip().upper()
    rc_implied_project = _KNOWN_RFI_PROJECTS.get(rfi_key, "")
    if not rc_implied_project:
        rc_implied_project = project_from_road_code(rec.get("road_code", ""))
    if not rc_implied_project:
        rc_implied_project = project_from_rfi_prefix(rec.get("rfi_no", ""))
    if not rc_implied_project:
        rc_implied_project = project_from_text(text)

    if rc_implied_project:
        rec["project_km"] = rc_implied_project
        if project and rc_implied_project != project:
            rec["_project_mismatch"] = True
    else:
        # Still unknown — use ZIP context (may be "" for generic ZIP.zip)
        rec["project_km"] = project

    return rec


# ══════════════════════════════════════════════════════════════════════════════
# LOAD ALL RECORDS
# ══════════════════════════════════════════════════════════════════════════════

def _project_island_from_name(name: str):
    n = name.upper()
    project     = "163.22 KM" if "163" in n else ("65.5 KM" if "65" in n else "")
    island_hint = ("Unguja" if ("UNGUJA" in n or "-U-" in n or "_U_" in n) else
                   ("Pemba"  if ("PEMBA"  in n or "-P-" in n or "_P_" in n) else ""))
    rfi_type    = "LAB" if "LAB" in n else "SITE"
    dm          = re.search(r"(\d{2})\.(\d{2})\.(\d{4})", name)
    date_hint   = f"{dm.group(1)}/{dm.group(2)}/{dm.group(3)}" if dm else ""
    return project, island_hint, rfi_type, date_hint


DOC_EXTENSIONS = {".pdf", ".docx", ".doc"}


def _process_doc(data: bytes, inner: str, archive_name: str,
                 project: str, island_hint: str, date_hint: str,
                 rfi_type: str, seen_rfi_nos: set,
                 rfi_records: list, skipped_invalid: list,
                 status_pdf_data: dict) -> None:
    """
    Validate the filename, then extract and accumulate one RFI document.
    Handles both individual RFI PDFs and status-summary PDFs.

    Args:
        data            : raw bytes of the document
        inner           : bare filename  (e.g. 'IRIS-Q-RFI-U-25-0042.pdf')
        archive_name    : parent ZIP/RAR name ('' for standalone)
        status_pdf_data : dict populated with status-summary bytes for later use
        skipped_invalid : list of filenames skipped because pattern didn't match
    """
    ext = Path(inner).suffix.lower()
    if ext not in DOC_EXTENSIONS:
        return

    base = os.path.basename(inner)

    # ── Is this a status-summary PDF? Store for post-processing, don't log. ──
    if is_status_pdf(base):
        status_pdf_data[base] = data
        print(f"      [STATUS-PDF] {base[:60]} — stored for status cross-reference")
        return

    # ── Parse first, then validate ────────────────────────────────────────────
    # We do NOT gate on filename alone — suffixes like _Stamped, spaces/double-
    # dashes between tokens, and sub-numbers all create false negatives.
    # Instead: parse the document, then accept it if EITHER the filename OR the
    # extracted content identifies it as an RFI.  Non-RFI supporting documents
    # (cover letters, photos converted to PDF, etc.) are silently ignored.
    try:
        text = (extract_text_pdf(data)
                if ext == ".pdf"
                else extract_text_docx(data))
        if not text.strip():
            return

        src = (os.path.join(ZIP_FOLDER, archive_name, base)
               if archive_name
               else os.path.join(ZIP_FOLDER, base))

        rec = extract_fields(text, archive_name or base, base,
                             project, island_hint, date_hint, src, rfi_type)

        rno = rec.get("rfi_no", "").strip()

        # Accept if the extracted RFI No. looks like a real RFI number, OR if
        # the filename itself matches the (now lenient) RFI prefix pattern.
        _content_is_rfi = bool(re.search(
            r"(?:IRIS|IR)[-\s]?(?:QAQC|Q)?[-\s]?RFI", rno, re.IGNORECASE
        ))
        if not _content_is_rfi and not is_valid_rfi_file(base):
            # Truly not an RFI document (e.g. cover letter, supporting photo PDF)
            return

        # ── Rule: -P- → Pemba, -U- → Unguja.  Violation → CANCELLED ──────────
        rfi_island = island_from_rfi_no(rno)
        if rfi_island and island_hint and rfi_island.lower() != island_hint.lower():
            rec["status"] = "CANCELLED"
            print(f"      [ISLAND MISMATCH] {base}: "
                  f"RFI No. says {rfi_island} but submitted in {island_hint} context → CANCELLED")

        # ── Rule: road code XX-2## → 163.22 KM; else 65.5 KM.
        #    If the ZIP context disagrees with the road code → CANCELLED ──────
        if rec.get("_project_mismatch"):
            rec["status"] = "CANCELLED"
            print(f"      [PROJECT MISMATCH] {base}: "
                  f"road code '{rec.get('road_code','')}' implies "
                  f"{rec.get('project_km','')} but ZIP says {archive_name} → CANCELLED")

        if rno in seen_rfi_nos:
            return
        seen_rfi_nos.add(rno)
        rfi_records.append(rec)

    except Exception as e:
        print(f"      ERROR {base}: {e}")


def _apply_status_from_summary_pdfs(
        rfi_records: list, status_pdf_data: dict
) -> dict:
    """
    Parse every status-summary PDF collected in load_all_records() and:
      1. Update Status / RE's Remarks / Remark Date in rfi_records in-place.
      2. Return a combined {rfi_no_upper: {status, re_remarks, remark_date}} dict
         so process_spreadsheet() can also patch EXISTING sheet rows.

    status_pdf_data : {filename: bytes}  — the "65.5 KM … with status.pdf" files
    rfi_records     : list of record dicts (keys: rfi_no, status, re_remarks, …)
    """
    combined_status_lookup: dict[str, dict] = {}   # ← returned to caller

    if not status_pdf_data:
        return combined_status_lookup

    # Build lookup: upper-case RFI No. → record (so we can mutate it in-place)
    lookup: dict[str, dict] = {r.get("rfi_no", "").upper().strip(): r
                                for r in rfi_records}

    _VALID_STATUSES = {
        "APPROVED", "APPROVED WITH COMMENTS", "NOT APPROVED",
        "CANCELLED", "OPENED", "CLOSED",
    }

    def _norm_status(raw: str) -> str:
        u = raw.strip().upper()
        if u in _VALID_STATUSES:
            return u
        if "NOT APPROVED" in u:    return "NOT APPROVED"
        if "WITH COMMENT" in u:    return "APPROVED WITH COMMENTS"
        if "APPROVED" in u:        return "APPROVED"
        if "CANCEL" in u:          return "CANCELLED"
        if "CLOSE" in u:           return "CLOSED"
        if "OPEN" in u:            return "OPENED"
        return ""   # unrecognised — caller skips blank strings

    total_updated = 0

    for pdf_name, pdf_bytes in status_pdf_data.items():
        print(f"\n  [STATUS] Parsing '{pdf_name}' …")
        # Per-PDF status lookup: {rfi_no_upper: {status, re_remarks, remark_date}}
        pdf_status: dict[str, dict] = {}

        try:
            with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf_doc:
                for page in pdf_doc.pages:
                    # ── Strategy 1: structured table extraction ────────────────
                    tables = (
                        page.extract_tables({
                            "vertical_strategy":      "lines",
                            "horizontal_strategy":    "lines",
                            "snap_tolerance":          6,
                            "intersection_tolerance":  6,
                        })
                        or page.extract_tables()
                        or []
                    )

                    page_yielded = False
                    for table in tables:
                        if not table:
                            continue
                        col_map: dict[int, str] = {}
                        header_found = False

                        for row in table:
                            if not row:
                                continue
                            cells = [c.replace("\n", " ").strip() if c else ""
                                     for c in row]
                            joined_low = " ".join(cells).lower()

                            # Header row detection
                            if not header_found and (
                                "rfi no" in joined_low or "s/no" in joined_low
                            ):
                                for ci, cell in enumerate(cells):
                                    key = cell.lower().strip()
                                    if "rfi no" in key:
                                        col_map[ci] = "rfi_no"
                                    elif "remark" in key and "date" in key:
                                        col_map[ci] = "remark_date"
                                    elif "remark" in key:
                                        col_map[ci] = "re_remarks"
                                    elif key == "status":
                                        col_map[ci] = "status"
                                header_found = True
                                continue

                            if not header_found or not col_map:
                                continue

                            first = cells[0].upper() if cells[0] else ""
                            # Skip aggregate / header-repeat rows
                            if re.match(r"^(opened|approved|total|sum|count|s/no|"
                                        r"cancelled|closed|not\s+approved)", first, re.I):
                                continue

                            rfi_no = ""
                            status = ""
                            re_remarks = ""
                            remark_date = ""

                            for ci, key in col_map.items():
                                if ci >= len(cells):
                                    continue
                                v = cells[ci]
                                if key == "rfi_no":        rfi_no      = v
                                elif key == "status":      status      = v
                                elif key == "re_remarks":  re_remarks  = v
                                elif key == "remark_date": remark_date = v

                            # Fallback: scan cells for an RFI-like value
                            if not rfi_no:
                                for c in cells:
                                    if re.match(r"^IR(?:IS)?[-_\s]", c, re.I) or \
                                       re.search(r"RFI[-_]", c, re.I):
                                        rfi_no = c.strip()
                                        break

                            if not rfi_no:
                                continue

                            rfi_key = rfi_no.upper().strip()
                            ns = _norm_status(status) if status else ""
                            if rfi_key and rfi_key not in pdf_status:
                                pdf_status[rfi_key] = {
                                    "status":      ns,
                                    "re_remarks":  re_remarks.strip(),
                                    "remark_date": remark_date.strip(),
                                }
                                page_yielded = True

                    # ── Strategy 2: word-level x-column fallback ───────────────
                    # Only if table extraction found nothing on this page
                    if page_yielded:
                        continue

                    words = page.extract_words(x_tolerance=4, y_tolerance=4) or []
                    if not words:
                        continue

                    # Group words into rows by y-coordinate (bucket of 6 px)
                    rows_by_y: dict[int, list] = {}
                    for w in words:
                        y_key = round(w["top"] / 6) * 6
                        rows_by_y.setdefault(y_key, []).append(w)

                    # Find header row → calibrate column x-boundaries
                    col_x: dict[str, float] = {}
                    for y_key in sorted(rows_by_y):
                        row_words = sorted(rows_by_y[y_key], key=lambda w: w["x0"])
                        row_text  = " ".join(w["text"] for w in row_words).lower()
                        if "rfi" in row_text and ("no." in row_text or "no" in row_text):
                            for w in row_words:
                                t = w["text"].lower()
                                if t in ("rfi", "no.", "no"):
                                    col_x.setdefault("rfi_no", w["x0"])
                                elif t == "status":
                                    col_x["status"] = w["x0"]
                                elif t in ("re's", "remarks"):
                                    col_x.setdefault("re_remarks", w["x0"])
                                elif t == "date" and "re_remarks" in col_x:
                                    col_x.setdefault("remark_date", w["x0"])
                            break

                    if not col_x:
                        continue

                    col_order = sorted(col_x.items(), key=lambda kv: kv[1])

                    def _assign_col(wx: float) -> str:
                        for i, (key, x0) in enumerate(col_order):
                            x1 = col_order[i + 1][1] if i + 1 < len(col_order) else 9999
                            if x0 - 8 <= wx <= x1 + 8:
                                return key
                        return ""

                    current: dict[str, list] = {}
                    prev_y: int | None = None

                    def _flush_word_row(cur: dict):
                        rfi_no = " ".join(cur.get("rfi_no", []))
                        if re.match(r"^IR", rfi_no, re.I):
                            rfi_key = rfi_no.upper().strip()
                            ns = _norm_status(" ".join(cur.get("status", [])))
                            if rfi_key and rfi_key not in pdf_status:
                                pdf_status[rfi_key] = {
                                    "status":      ns,
                                    "re_remarks":  " ".join(cur.get("re_remarks", [])).strip(),
                                    "remark_date": " ".join(cur.get("remark_date", [])).strip(),
                                }

                    for y_key in sorted(rows_by_y):
                        row_words = sorted(rows_by_y[y_key], key=lambda w: w["x0"])
                        if prev_y is None or abs(y_key - prev_y) > 12:
                            if current:
                                _flush_word_row(current)
                            current = {}
                        prev_y = y_key
                        for w in row_words:
                            col = _assign_col(w["x0"])
                            if col:
                                current.setdefault(col, []).append(w["text"])

                    if current:
                        _flush_word_row(current)

        except Exception as e:
            print(f"    ERROR parsing status PDF '{pdf_name}': {e}")
            continue

        # ── Apply the status lookup to freshly-parsed rfi_records ─────────────
        updated = 0
        for rfi_key, info in pdf_status.items():
            if rfi_key in lookup:
                rec = lookup[rfi_key]
                if info["status"]:
                    rec["status"] = info["status"]
                if info["re_remarks"]:
                    rec["re_remarks"] = info["re_remarks"]
                if info["remark_date"]:
                    rec["remark_date"] = info["remark_date"]
                updated += 1
            # Also accumulate into the global status dict (for sheet row updates)
            combined_status_lookup[rfi_key] = info

        pct = f"{100*updated//len(pdf_status)}%" if pdf_status else "—"
        print(f"    ✓ Updated {updated}/{len(pdf_status)} RFI status entries "
              f"({pct} match rate) from '{pdf_name}'")
        total_updated += updated

    print(f"\n  [STATUS] Total rfi_records updated from summary PDFs: {total_updated}")
    return combined_status_lookup


def load_all_records():
    """
    Scan ZIP_FOLDER for:
      • ZIP archives  → extract & process every doc inside
      • RAR archives  → extract & process every doc inside
      • Standalone docs → process directly

    Filename filter: only files matching the official IRIS/IR RFI pattern are
    logged to Google Sheets.  Status-summary PDFs ("65.5 KM … with status.pdf")
    are collected separately for post-processing status updates.
    """
    all_files = sorted(os.listdir(ZIP_FOLDER))
    zip_files  = [f for f in all_files if f.lower().endswith(".zip")]
    rar_files  = [f for f in all_files if f.lower().endswith(".rar")]
    standalone = [f for f in all_files
                  if Path(f).suffix.lower() in DOC_EXTENSIONS
                  and not f.startswith("_")]

    print(f"Found {len(zip_files)} ZIP + {len(rar_files)} RAR "
          f"+ {len(standalone)} standalone files\n")

    rfi_records      = []
    seen_rfi_nos     = set()
    download_log     = []
    skipped_invalid  = []   # files that didn't match the RFI pattern
    status_pdf_data  = {}   # {filename: bytes}  — the summary PDFs

    # ── Helper: log archive to download_log ───────────────────────────────────
    def _log_archive(fname, atype):
        fpath = os.path.join(ZIP_FOLDER, fname)
        proj, isl, rt, dh = _project_island_from_name(fname)
        try:
            sz = round(os.stat(fpath).st_size / 1024, 1)
            mod = datetime.datetime.fromtimestamp(
                os.stat(fpath).st_mtime).strftime("%d/%m/%Y %H:%M")
        except Exception:
            sz, mod = 0, ""
        download_log.append({
            "filename": fname, "type": atype, "size_kb": sz,
            "modified": mod, "project": proj, "island": isl, "rfi_type": rt,
        })
        return proj, isl, rt, dh

    # ── ZIPs ──────────────────────────────────────────────────────────────────
    total_arc = len(zip_files) + len(rar_files)
    arc_idx   = 0

    for zf_name in zip_files:
        arc_idx += 1
        zf_path = os.path.join(ZIP_FOLDER, zf_name)
        project, island_hint, rfi_type, date_hint = _log_archive(zf_name, "ZIP")
        try:
            with zipfile.ZipFile(zf_path, "r") as zf:
                entries = [x for x in zf.namelist()
                           if Path(x).suffix.lower() in DOC_EXTENSIONS]
                print(f"  [{arc_idx}/{total_arc}] ZIP {zf_name[:55]} ({len(entries)} docs)")
                for entry in entries:
                    inner = os.path.basename(entry)
                    try:
                        data = zf.read(entry)
                    except Exception as e:
                        print(f"      READ ERROR {inner}: {e}")
                        continue
                    _process_doc(data, inner, zf_name,
                                 project, island_hint, date_hint, rfi_type,
                                 seen_rfi_nos, rfi_records,
                                 skipped_invalid, status_pdf_data)
        except zipfile.BadZipFile as e:
            print(f"  BAD ZIP {zf_name}: {e}")
        except Exception as e:
            print(f"  ERROR reading ZIP {zf_name}: {e}")

    # ── RARs ──────────────────────────────────────────────────────────────────
    for rf_name in rar_files:
        arc_idx += 1
        rf_path = os.path.join(ZIP_FOLDER, rf_name)
        project, island_hint, rfi_type, date_hint = _log_archive(rf_name, "RAR")
        print(f"  [{arc_idx}/{total_arc}] RAR {rf_name[:55]} … extracting")
        try:
            members = _open_rar(rf_path)   # → [(member_name, bytes), …]
            doc_members = [(n, d) for n, d in members
                           if Path(n).suffix.lower() in DOC_EXTENSIONS]
            print(f"    {len(doc_members)} document(s) found inside")
            for entry_name, data in doc_members:
                inner = os.path.basename(entry_name)
                _process_doc(data, inner, rf_name,
                             project, island_hint, date_hint, rfi_type,
                             seen_rfi_nos, rfi_records,
                             skipped_invalid, status_pdf_data)
        except Exception as e:
            print(f"  ERROR reading RAR {rf_name}: {e}")

    # ── Standalone docs ────────────────────────────────────────────────────────
    if standalone:
        print(f"\nProcessing {len(standalone)} standalone files …")
    for fname in standalone:
        fpath = os.path.join(ZIP_FOLDER, fname)
        project, island_hint, rfi_type, date_hint = _log_archive(fname, "FILE")
        ext = Path(fname).suffix.lower()
        try:
            with open(fpath, "rb") as f:
                data = f.read()
        except Exception as e:
            print(f"  ERROR reading {fname}: {e}")
            continue
        _process_doc(data, fname, "",
                     project, island_hint, date_hint, rfi_type,
                     seen_rfi_nos, rfi_records,
                     skipped_invalid, status_pdf_data)

    # ── Report skipped files ───────────────────────────────────────────────────
    if skipped_invalid:
        print(f"\n  Skipped {len(skipped_invalid)} file(s) (name doesn't match RFI pattern):")
        for s in skipped_invalid[:15]:
            print(f"    • {s}")
        if len(skipped_invalid) > 15:
            print(f"    … and {len(skipped_invalid) - 15} more")

    status_lookup: dict = {}
    if status_pdf_data:
        print(f"\n  Status-summary PDFs collected: {list(status_pdf_data.keys())}")
        # Parse PDFs → update freshly-parsed rfi_records + return lookup for sheets
        status_lookup = _apply_status_from_summary_pdfs(rfi_records, status_pdf_data)

    print(f"\nTotal unique RFI records: {len(rfi_records)}")
    return rfi_records, download_log, status_lookup


# ══════════════════════════════════════════════════════════════════════════════
# POST-PROCESSING: merged station lengths
# ══════════════════════════════════════════════════════════════════════════════

def attach_merged_lengths(records: list) -> None:
    groups = defaultdict(list)
    for rec in records:
        if not is_mergeable_activity(rec.get("activity_name", "")):
            continue
        key = (rec.get("project_km", ""), rec.get("island", ""),
               rec.get("road_code", ""),  rec.get("activity_name", ""))
        groups[key].append(rec)

    for key, grp in groups.items():
        intervals = []
        for r in grp:
            for s, e in r.get("all_ranges", []):
                if s != e:
                    intervals.append((min(s, e), max(s, e)))
        ml = merged_length_km(intervals)
        for r in grp:
            r["merged_length_km"] = ml if ml is not None else ""

    print("\n── Merged Length Arithmetic (Subgrade / Excavation / Filling) ──")
    for key, grp in sorted(groups.items()):
        intervals = []
        for r in grp:
            for s, e in r.get("all_ranges", []):
                if s != e:
                    intervals.append((min(s, e), max(s, e)))
        merged = merge_intervals(intervals)
        total  = round(sum(e - s for s, e in merged) / 1000, 4)
        rfi_nos = [r["rfi_no"] for r in grp]
        print(f"  {key[2] or '?'} | {key[3]} | {key[1]}")
        print(f"    RFIs  : {', '.join(rfi_nos[:6])}{'…' if len(rfi_nos) > 6 else ''}")
        print(f"    Total : {total} km")


# ══════════════════════════════════════════════════════════════════════════════
# COLUMN DEFINITIONS
# ══════════════════════════════════════════════════════════════════════════════

DATA_COLUMNS = [
    "S/No.",               # 0
    "RFI No.",             # 1
    "Date Received",       # 2
    "Inspection Date",     # 3
    "Road Code",           # 4
    "Road Name",           # 5
    "Island",              # 6
    "Contract Number",     # 7  ← formula field
    "Start Station",       # 8
    "End Station",         # 9
    "Length Executed (km)",# 10 ← formula field
    "Merged Length (km)",  # 11
    "Activity Name",       # 12
    "Work Description",    # 13
    "Station Range",       # 14
    "Status",              # 15 ← dropdown validated
    "RE's Remarks",        # 16
    "Remark Date",         # 17
    "Attachment Path",     # 18
    "Project",             # 19
]
_CI = {col: i for i, col in enumerate(DATA_COLUMNS)}

# Maximum number of columns written per row.  All writes stop at column T
# (the "Project" column, 0-based index 19).  Nothing is ever written to
# column U or beyond, regardless of what the table metadata reports.
MAX_COLS = len(DATA_COLUMNS)   # = 20  (A through T inclusive)


def _write_end_col(tinfo: dict) -> int:
    """
    Return the 0-based EXCLUSIVE end column for any write operation.
    Capped so that writes never pass column T (start_col + MAX_COLS).
    """
    return min(tinfo["end_col"], tinfo["start_col"] + MAX_COLS)


def _cap_row(row: list) -> list:
    """
    Truncate (or pad) a row list to exactly MAX_COLS entries.
    This ensures no cell beyond column T is ever touched.
    """
    if len(row) > MAX_COLS:
        return row[:MAX_COLS]
    if len(row) < MAX_COLS:
        return row + [""] * (MAX_COLS - len(row))
    return row


def record_to_row(sno: int, rec: dict) -> list:
    """Plain static row — used when header_map is unavailable."""
    return [
        sno,
        rec.get("rfi_no", ""),
        rec.get("date_received", ""),
        rec.get("inspection_date", ""),
        rec.get("road_code", ""),
        rec.get("road_name", ""),
        rec.get("island", ""),
        rec.get("contract_number", ""),
        rec.get("start_station", ""),
        rec.get("end_station", ""),
        rec.get("length_km", ""),
        rec.get("merged_length_km", ""),
        rec.get("activity_name", ""),
        rec.get("work_description", ""),
        rec.get("station_range", ""),
        canonicalize_status(rec.get("status", "")),
        rec.get("re_remarks", ""),
        rec.get("remark_date", ""),
        rec.get("attachment_path", ""),
        rec.get("project_km", ""),
    ][:MAX_COLS]   # hard cap — nothing beyond column T


def build_row_for_table(sno: int, rec: dict, header_map: dict,
                         tinfo: dict, row_1based: int) -> list:
    """
    Build a row aligned to the actual table header.
    Uses Google Sheets formula strings for:
      • Length Executed (km)  — derived from Start/End Station cells
      • Contract Number       — IF formula driven by Project cell

    row_1based : the 1-based sheet row number where this row will land
                 (needed to write correct A1-notation cell references in formulas).
    tinfo      : table info dict (provides start_col for column letter computation).
    """
    if not header_map:
        return record_to_row(sno, rec)

    def _col(col_name: str) -> str:
        """Return the sheet column letter for a named column, or empty string."""
        if col_name in header_map:
            return col_letter(tinfo["start_col"] + header_map[col_name])
        return ""

    r = row_1based

    # ── Formula: Length Executed (km) ─────────────────────────────────────────
    sc = _col("Start Station")
    ec = _col("End Station")
    if sc and ec:
        length_val = (
            f'=IF(AND({sc}{r}<>"",{ec}{r}<>"",'
            f'ISNUMBER(FIND("+",{sc}{r})),ISNUMBER(FIND("+",{ec}{r}))),'
            f'ABS('
            f'(VALUE(LEFT({ec}{r},FIND("+",{ec}{r})-1))*1000'
            f'+VALUE(MID({ec}{r},FIND("+",{ec}{r})+1,100)))'
            f'-(VALUE(LEFT({sc}{r},FIND("+",{sc}{r})-1))*1000'
            f'+VALUE(MID({sc}{r},FIND("+",{sc}{r})+1,100)))'
            f')/1000,'
            f'IF(AND({sc}{r}<>"",{ec}{r}<>"",'
            f'ISNUMBER(VALUE({sc}{r})),ISNUMBER(VALUE({ec}{r}))),'
            f'ABS(VALUE({ec}{r})-VALUE({sc}{r}))/1000,""))'
        )
    else:
        length_val = rec.get("length_km", "")

    # ── Formula: Contract Number ───────────────────────────────────────────────
    pc = _col("Project")
    if pc:
        contract_val = (
            f'=IF(ISNUMBER(SEARCH("163",{pc}{r})),'
            f'"{CONTRACT_163}","{CONTRACT_655}")'
        )
    else:
        contract_val = rec.get("contract_number", "")

    # ── Road Name: prefer extracted value; fall back to VLOOKUP formula ─────────
    road_name_val = rec.get("road_name", "").strip()
    if not road_name_val:
        # Build a VLOOKUP from the Road Code cell into "Road Code & Road Names"!A:B
        rc = _col("Road Code")
        if rc:
            road_name_val = (
                f'=IFERROR(VLOOKUP({rc}{r},'
                f"'{ROAD_CODE_SHEET_NAME}'!$A:$B,2,FALSE),"
                f'IF({rc}{r}<>"",{rc}{r},""))'
            )
        else:
            road_name_val = rec.get("road_code", "")

    # ── Island: default from RFI No. if blank ────────────────────────────────
    island_val = rec.get("island", "").strip()
    if not island_val:
        island_val = island_from_rfi_no(rec.get("rfi_no", "")) or ""

    # ── Activity Name: ensure never blank ─────────────────────────────────────
    activity_val = rec.get("activity_name", "").strip()
    if not activity_val:
        activity_val = normalise_activity(rec.get("work_description", "")) or "OTHER"

    # ── Static fields ──────────────────────────────────────────────────────────
    field_values = {
        "S/No.":              sno,
        "RFI No.":            rec.get("rfi_no", ""),
        "Date Received":      rec.get("date_received", ""),
        "Inspection Date":    rec.get("inspection_date", ""),
        "Road Code":          rec.get("road_code", ""),
        "Road Name":          road_name_val,
        "Island":             island_val,
        "Contract Number":    contract_val,
        "Start Station":      rec.get("start_station", ""),
        "End Station":        rec.get("end_station", ""),
        "Length Executed (km)": length_val,
        "Merged Length (km)": rec.get("merged_length_km", ""),
        "Activity Name":      activity_val,
        "Work Description":   rec.get("work_description", ""),
        "Station Range":      rec.get("station_range", ""),
        "Status":             canonicalize_status(rec.get("status", "")),
        "RE's Remarks":       rec.get("re_remarks", ""),
        "Remark Date":        rec.get("remark_date", ""),
        "Attachment Path":    rec.get("attachment_path", ""),
        "Project":            rec.get("project_km", ""),
    }

    n_cols = min(max(header_map.values()) + 1, MAX_COLS)  # never allocate past col T
    row    = [""] * n_cols
    for col_name, value in field_values.items():
        idx = header_map.get(col_name)
        if idx is not None and idx < n_cols:    # skip any col beyond T
            row[idx] = value if value is not None else ""

    return _cap_row(row)   # guarantee exactly MAX_COLS elements


# ══════════════════════════════════════════════════════════════════════════════
# GOOGLE AUTH
# ══════════════════════════════════════════════════════════════════════════════

def get_google_creds():
    """
    Load or create OAuth credentials.

    Auto-detects stale tokens that were generated with narrower scopes
    (e.g. spreadsheets.readonly from a previous run) and forces a clean
    re-authorisation so that write operations succeed.
    """
    # Service account short-circuit -- preferred for unattended/cloud runs
    # (no browser, no token to ever expire). Falls through to the existing
    # OAuth flow below when GOOGLE_SERVICE_ACCOUNT_PATH is not set.
    if SERVICE_ACCOUNT_FILE:
        if not os.path.exists(SERVICE_ACCOUNT_FILE):
            print(f"\n❌  Missing service account file: {SERVICE_ACCOUNT_FILE}")
            sys.exit(1)
        print(f"  Using service account credentials ({SERVICE_ACCOUNT_FILE})")
        return service_account.Credentials.from_service_account_file(
            SERVICE_ACCOUNT_FILE, scopes=SCOPES)

    creds = None
    if os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)

        # ── Scope validation ────────────────────────────────────────────────
        # If the saved token was created with fewer scopes than we now require
        # (e.g. the script previously ran with spreadsheets.readonly),
        # silently discard it and force a fresh browser authorisation.
        token_scopes = getattr(creds, "scopes", None) or []
        missing_scopes = set(SCOPES) - set(token_scopes)
        if missing_scopes:
            print(
                f"\n  ⚠  Token at {TOKEN_FILE} is missing required scopes:\n"
                + "".join(f"       • {s}\n" for s in sorted(missing_scopes))
                + "     Deleting stale token and re-authorising …\n"
            )
            try:
                os.remove(TOKEN_FILE)
            except OSError:
                pass
            creds = None

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            import socket as _socket
            _old_timeout = _socket.getdefaulttimeout()
            _socket.setdefaulttimeout(30)   # don't hang >30 s waiting for OAuth endpoint
            try:
                creds.refresh(Request())
            except Exception as refresh_err:
                err_str = str(refresh_err).lower()
                if any(kw in err_str for kw in ("invalid_grant", "token has been expired",
                                                 "token has been revoked", "revoked")):
                    # Token revoked — delete it so next run starts clean.
                    print(f"  ❌  Google token revoked/expired: {refresh_err}")
                    print(f"  ↳  Deleting {TOKEN_FILE}")
                    print(f"  ↳  Run this script MANUALLY once to re-authenticate in the browser.")
                    try:
                        os.remove(TOKEN_FILE)
                    except OSError:
                        pass
                    sys.exit(1)   # fast-fail
                else:
                    print(f"  ⚠  Token refresh failed ({refresh_err}); re-authorising …")
                creds = None
            finally:
                _socket.setdefaulttimeout(_old_timeout)

        if not creds:
            if not os.path.exists(CREDENTIALS_FILE):
                print(f"\n❌  Missing: {CREDENTIALS_FILE}")
                sys.exit(1)
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
            creds = flow.run_local_server(port=0)

        with open(TOKEN_FILE, "w") as f:
            f.write(creds.to_json())

    return creds


# ══════════════════════════════════════════════════════════════════════════════
# UTILITY
# ══════════════════════════════════════════════════════════════════════════════

# Counts consecutive 403-permission errors across _api_retry calls.
# When it reaches _PERM_ERR_THRESHOLD the script halts with a fix guide.
_perm_err_count: int = 0
_PERM_ERR_THRESHOLD: int = 3


def _api_retry(fn, *args, retries=5, delay=2, **kwargs):
    """
    Call a Google API method with its kwargs and execute(), retrying on
    quota errors AND transient TCP/connection errors (WinError 10053, etc.).

    403 "permission" errors are NOT retried — they indicate a stale/
    read-only token.  After _PERM_ERR_THRESHOLD such failures the script
    halts and prints the exact steps to fix the token.
    """
    global _perm_err_count
    for attempt in range(retries):
        try:
            result = fn(*args, **kwargs).execute()
            # Successful call — reset the 403 counter.
            _perm_err_count = 0
            return result
        except Exception as e:
            err_str = str(e).lower()

            # ── 403 permission error ─────────────────────────────────────
            if "403" in str(e) and "permission" in err_str:
                _perm_err_count += 1
                if _perm_err_count >= _PERM_ERR_THRESHOLD:
                    _banner = "═" * 68
                    print(
                        f"\n{_banner}\n"
                        "  ❌  FATAL — Google token lacks write permission\n"
                        f"{_banner}\n"
                        "  Every write is returning HTTP 403.  This means the saved\n"
                        "  token was authorised with read-only scopes.\n\n"
                        "  FIX (run on your Windows machine):\n\n"
                        f"    1. Delete the stale token:\n"
                        f"         del \"{TOKEN_FILE}\"\n\n"
                        "    2. Re-run the script:\n"
                        "         python C:\\AI-Projects\\update_rfi_google_sheets.py\n\n"
                        "    3. A browser window will open — sign in and click\n"
                        "       'Allow' for ALL requested permissions.\n\n"
                        "    4. The new token will be saved automatically and\n"
                        "       all future runs will have write access.\n"
                        f"{_banner}\n"
                    )
                    raise PermissionError(
                        "Token lacks write permission. "
                        f"Delete {TOKEN_FILE} and re-run."
                    ) from e
                # Let the caller print its own ⚠ warning; just raise.
                raise

            # ── Transient connection error (WinError 10053, reset, etc.) ─
            is_conn_err = isinstance(
                e, (OSError, ConnectionAbortedError, ConnectionResetError,
                    ConnectionError, BrokenPipeError)
            )
            if not is_conn_err:
                # Also catch httplib2 / socket errors carried inside other
                # exception types by checking the string representation.
                is_conn_err = any(tag in err_str for tag in (
                    "winerror 10053", "winerror 10054", "winerror 10061",
                    "connection aborted", "connection reset",
                    "connection refused", "remotedisconnected",
                    "broken pipe", "errno 104", "errno 111",
                ))

            # ── Quota error OR connection error — wait and retry ─────────
            if attempt < retries - 1 and ("quota" in err_str or is_conn_err):
                if "quota" in err_str:
                    wait = 65 if attempt == 0 else delay * (attempt + 1)
                    label = "Quota limit hit"
                else:
                    wait = min(10 * (attempt + 1), 30)   # 10 s, 20 s, 30 s …
                    label = "Connection error"
                print(
                    f"    ⏳ {label} — waiting {wait}s "
                    f"before retry {attempt + 2}/{retries} …"
                )
                time.sleep(wait)
            else:
                raise


def col_letter(n: int) -> str:
    """Convert 0-based column index to letter(s): 0→A, 25→Z, 26→AA …"""
    result = ""
    n += 1
    while n > 0:
        n, rem = divmod(n - 1, 26)
        result = chr(65 + rem) + result
    return result


# ══════════════════════════════════════════════════════════════════════════════
# NAMED TABLE FUNCTIONS
# ══════════════════════════════════════════════════════════════════════════════

def get_all_named_tables(service, ssid: str) -> dict:
    """
    Return all named Google Sheets Tables in the spreadsheet.
    Returns: {table_name: {sheet_title, sheet_id, start_row, end_row,
                           start_col, end_col, range_a1}}
    All row/col indices are 0-based; end_row is exclusive.
    """
    meta   = service.spreadsheets().get(spreadsheetId=ssid).execute()
    tables = {}

    for sheet in meta.get("sheets", []):
        props       = sheet.get("properties", {})
        sheet_id    = props.get("sheetId")
        sheet_title = props.get("title", "")
        # Progress tables share this workbook but are outside RFI maintenance.
        if sheet_title == PP_SHEET_NAME:
            continue

        for tbl in sheet.get("tables", []):
            name = tbl.get("name", "").strip()
            if not name:
                continue
            rng        = tbl.get("range", {})
            start_row  = rng.get("startRowIndex", 0)
            end_row    = rng.get("endRowIndex", start_row + 1)
            start_col  = rng.get("startColumnIndex", 0)
            end_col    = rng.get("endColumnIndex", len(DATA_COLUMNS))
            # A1 notation spanning the full table (header through footer inclusive)
            range_a1   = (f"'{sheet_title}'!"
                          f"{col_letter(start_col)}{start_row + 1}:"
                          f"{col_letter(end_col - 1)}{end_row}")
            tables[name] = {
                "sheet_title": sheet_title,
                "sheet_id":    sheet_id,
                "start_row":   start_row,   # 0-based, header row
                "end_row":     end_row,     # 0-based exclusive (= 1-based last row)
                "start_col":   start_col,
                "end_col":     end_col,
                "range_a1":    range_a1,
            }

    return tables


def get_table_existing_data(service, ssid: str, tinfo: dict) -> tuple:
    """
    Read existing data from a named table.
    Returns: (header_map, existing_rfi_nos, max_sno, data_rows_count)

      header_map        : {column_name: 0-based column index within the row}
      existing_rfi_nos  : set of RFI No. strings already in the table
      max_sno           : highest S/No. value found in the table (0 if none)
      data_rows_count   : number of data rows (excluding header and footer)
    """
    try:
        result = _api_retry(
            service.spreadsheets().values().get,
            spreadsheetId=ssid,
            range=tinfo["range_a1"]
        )
        rows = result.get("values", [])
    except Exception as e:
        print(f"    Warning: could not read table '{tinfo.get('range_a1')}': {e}")
        return {}, set(), 0, 0

    if not rows:
        return {}, set(), 0, 0

    # Row 0 = header
    header     = [str(h).strip() for h in rows[0]]
    header_map = {h: i for i, h in enumerate(header)}

    # Case-insensitive fallback: if a DATA_COLUMNS canonical name (e.g. "Road Name")
    # is absent from header_map but its lower-case matches a header that happens to be
    # all-caps (e.g. "ROAD NAME"), add the canonical name so all downstream lookups work.
    _hmap_lower = {k.lower(): v for k, v in header_map.items()}
    for _col in DATA_COLUMNS:
        if _col not in header_map:
            _alt = _hmap_lower.get(_col.lower())
            if _alt is not None:
                header_map[_col] = _alt

    rfi_col = header_map.get("RFI No.")
    sno_col = header_map.get("S/No.")

    existing_rfis  = set()
    max_sno        = 0
    data_row_count = 0

    # Skip header (rows[0]) and footer (rows[-1] if table has > 1 row).
    # Footer detection rules (any one is sufficient):
    #   a) Row contains "TOTAL" or "SUM"
    #   b) S/No. cell is non-numeric (e.g. "TOTAL", formula text)
    #   c) S/No. cell is blank — Google Sheets omits trailing empty cells,
    #      so sno_col >= len(last) means the cell was empty → footer
    data_rows = rows[1:]  # strip header
    if len(data_rows) > 0:
        last = data_rows[-1]
        last_text = " ".join(str(c) for c in last).upper()
        sno_is_blank_or_nonnumeric = (
            sno_col is not None and (
                sno_col >= len(last) or                              # blank cell (trailing empty omitted)
                not str(last[sno_col]).strip().lstrip("-").isdigit() # non-numeric value
            )
        )
        is_footer = (
            "TOTAL" in last_text or
            "SUM"   in last_text or
            sno_is_blank_or_nonnumeric
        )
        if is_footer:
            data_rows = data_rows[:-1]  # exclude footer

    for row in data_rows:
        if rfi_col is not None and len(row) > rfi_col:
            v = str(row[rfi_col]).strip()
            if v:
                existing_rfis.add(v)
                data_row_count += 1
        if sno_col is not None and len(row) > sno_col:
            try:
                val = int(float(str(row[sno_col]).strip()))
                max_sno = max(max_sno, val)
            except (ValueError, TypeError):
                pass

    return header_map, existing_rfis, max_sno, data_row_count


def _date_sort_key(date_str: str) -> tuple:
    """
    Parse a date string into a (year, month, day) tuple for chronological sorting.
    Tries multiple formats.  Unknown / blank dates sort to the very end.
    """
    if not date_str or not date_str.strip():
        return (9999, 12, 31)
    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y",
                "%d/%m/%y", "%m/%d/%Y", "%Y-%m-%d"):
        try:
            d = datetime.datetime.strptime(date_str.strip(), fmt)
            return (d.year, d.month, d.day)
        except ValueError:
            continue
    return (9999, 12, 31)


def _read_existing_date_rows(service, ssid: str,
                              tinfo: dict, header_map: dict) -> list:
    """
    Return a list of (date_sort_key, row_0based_abs) for every existing data row,
    in top-to-bottom order, so callers can binary-search for insertion positions.
    """
    date_col = header_map.get("Inspection Date")
    if date_col is None:
        return []
    raw_rows = _read_data_rows_raw(service, ssid, tinfo)
    result = []
    for row_0b, row in raw_rows:
        cell = row[date_col].strip() if date_col < len(row) else ""
        result.append((_date_sort_key(cell), row_0b))
    return result   # already in row order


def _find_insert_pos(existing_date_rows: list, new_date_key: tuple,
                     new_rfi_no: str, footer_0b: int) -> int:
    """
    Return the 0-based absolute row index to insert BEFORE.
    Inserts before the first existing row whose date is LATER than new_date_key.
    Same-date ties: sorted by RFI No. string (not implemented here — just use footer).
    Default: above the footer row.
    """
    for dk, row_0b in existing_date_rows:
        if dk > new_date_key:
            return row_0b
    return footer_0b


def _insert_single_blank_row(service, ssid: str, sheet_id: int,
                              before_row_0b: int) -> None:
    """Insert one blank row at `before_row_0b` (0-based), pushing everything down."""
    _api_retry(
        service.spreadsheets().batchUpdate,
        spreadsheetId=ssid,
        body={"requests": [{
            "insertDimension": {
                "range": {
                    "sheetId":    sheet_id,
                    "dimension":  "ROWS",
                    "startIndex": before_row_0b,
                    "endIndex":   before_row_0b + 1,
                },
                "inheritFromBefore": True,
            }
        }]}
    )
    time.sleep(0.3)


def _write_single_row(service, ssid: str, tinfo: dict,
                      row_0based: int, row_data: list) -> None:
    """Write row_data into the specific 0-based sheet row."""
    sheet_title = tinfo["sheet_title"]
    start_col   = tinfo["start_col"]
    capped      = _cap_row(row_data)
    row_1b      = row_0based + 1
    rng = (f"'{sheet_title}'!"
           f"{col_letter(start_col)}{row_1b}:"
           f"{col_letter(start_col + len(capped) - 1)}{row_1b}")
    _api_retry(
        service.spreadsheets().values().update,
        spreadsheetId=ssid,
        range=rng,
        valueInputOption="USER_ENTERED",
        body={"values": [capped]}
    )


def insert_rows_above_footer(service, ssid: str, tinfo: dict, rows: list) -> None:
    """
    Insert new data rows ABOVE the table footer row (last row of the table).

    Strategy:
      1. insertDimension  — add N blank rows at the footer's 0-based index,
                            pushing the footer down; inherit formatting from above.
      2. values.update    — fill the newly created rows with data.

    This guarantees every entry stays inside the named table range and above
    any totals/footer row.
    """
    if not rows:
        return

    n_rows      = len(rows)
    sheet_id    = tinfo["sheet_id"]
    sheet_title = tinfo["sheet_title"]
    start_col   = tinfo["start_col"]
    end_col     = _write_end_col(tinfo)   # capped at col T (start_col + MAX_COLS)

    # Cap every row so nothing is written past column T
    capped_rows = [_cap_row(r) for r in rows]

    # Footer is at 0-based index  end_row - 1
    # Insert new rows starting at that 0-based index (footer shifts down by n_rows)
    footer_0based = tinfo["end_row"] - 1

    # ── Step 1: insert blank rows above footer ─────────────────────────────────
    _api_retry(
        service.spreadsheets().batchUpdate,
        spreadsheetId=ssid,
        body={
            "requests": [{
                "insertDimension": {
                    "range": {
                        "sheetId":    sheet_id,
                        "dimension":  "ROWS",
                        "startIndex": footer_0based,
                        "endIndex":   footer_0based + n_rows,
                    },
                    "inheritFromBefore": True,
                }
            }]
        }
    )
    time.sleep(0.5)  # let Sheets settle before writing

    # ── Step 2: write values into the newly inserted rows ──────────────────────
    # After insertion the new rows are at 1-based rows:
    #   footer_0based + 1  through  footer_0based + n_rows
    first_1based = footer_0based + 1          # 1-based sheet row of first new data row
    last_1based  = footer_0based + n_rows     # 1-based sheet row of last new data row

    write_range = (
        f"'{sheet_title}'!"
        f"{col_letter(start_col)}{first_1based}:"
        f"{col_letter(end_col - 1)}{last_1based}"
    )

    _api_retry(
        service.spreadsheets().values().update,
        spreadsheetId=ssid,
        range=write_range,
        valueInputOption="USER_ENTERED",
        body={"values": capped_rows}
    )


def set_status_validation(service, ssid: str, tinfo: dict,
                          header_map: dict, data_start_1based: int) -> None:
    """
    Apply a dropdown data-validation rule to the Status column of the table
    for all data rows (i.e., between header and footer).

    data_start_1based : 1-based row number of the first data row (header row + 1).
    """
    status_col_in_hmap = header_map.get("Status")
    if status_col_in_hmap is None:
        return

    sheet_col_0based = tinfo["start_col"] + status_col_in_hmap
    sheet_id         = tinfo["sheet_id"]

    # Data rows span from data_start_1based to end_row (1-based, inclusive before footer)
    # In 0-based terms: data_start_1based-1  through  end_row - 2
    row_start_0based = data_start_1based - 1   # 0-based inclusive
    row_end_0based   = tinfo["end_row"] - 1    # 0-based exclusive (stops before footer)

    if row_end_0based <= row_start_0based:
        return

    _api_retry(
        service.spreadsheets().batchUpdate,
        spreadsheetId=ssid,
        body={
            "requests": [{
                "setDataValidation": {
                    "range": {
                        "sheetId":          sheet_id,
                        "startRowIndex":    row_start_0based,
                        "endRowIndex":      row_end_0based,
                        "startColumnIndex": sheet_col_0based,
                        "endColumnIndex":   sheet_col_0based + 1,
                    },
                    "rule": {
                        "condition": {
                            "type": "ONE_OF_LIST",
                            "values": [
                                {"userEnteredValue": v}
                                for v in _STATUS_DROPDOWN
                            ],
                        },
                        "showCustomUi": True,
                        "strict":       False,
                    }
                }
            }]
        }
    )


def set_activity_validation(service, ssid: str, tinfo: dict,
                            header_map: dict, data_start_1based: int) -> None:
    """
    Apply a dropdown data-validation rule to the Activity Name column so every
    cell shows the full canonical activity list as a picker.
    """
    act_col_in_hmap = header_map.get("Activity Name")
    if act_col_in_hmap is None:
        return

    sheet_col_0based = tinfo["start_col"] + act_col_in_hmap
    sheet_id         = tinfo["sheet_id"]
    row_start_0based = data_start_1based - 1
    row_end_0based   = tinfo["end_row"] - 1   # exclusive, stops before footer

    if row_end_0based <= row_start_0based:
        return

    _api_retry(
        service.spreadsheets().batchUpdate,
        spreadsheetId=ssid,
        body={
            "requests": [{
                "setDataValidation": {
                    "range": {
                        "sheetId":          sheet_id,
                        "startRowIndex":    row_start_0based,
                        "endRowIndex":      row_end_0based,
                        "startColumnIndex": sheet_col_0based,
                        "endColumnIndex":   sheet_col_0based + 1,
                    },
                    "rule": {
                        "condition": {
                            "type": "ONE_OF_LIST",
                            "values": [
                                {"userEnteredValue": v}
                                for v in ACTIVITY_DROPDOWN
                            ],
                        },
                        "showCustomUi": True,
                        "strict":       False,
                    }
                }
            }]
        }
    )


def _col_letter(col_0based: int) -> str:
    """Convert 0-based column index to A1 column letter (0→A, 25→Z, 26→AA …)."""
    result = ""
    n = col_0based + 1
    while n:
        n, r = divmod(n - 1, 26)
        result = chr(65 + r) + result
    return result


def _delete_conditional_formats_in_range(
        service, ssid: str, sheet_id: int,
        row_start_0b: int, row_end_0b: int) -> None:
    """
    Remove every existing conditional format rule on the given sheet that
    overlaps [row_start_0b, row_end_0b).  Deletes highest-index first so
    the remaining indices stay valid throughout.
    """
    try:
        meta = _api_retry(
            service.spreadsheets().get,
            spreadsheetId=ssid,
            fields="sheets(properties(sheetId),conditionalFormats)"
        )
        for sh in meta.get("sheets", []):
            if sh.get("properties", {}).get("sheetId") != sheet_id:
                continue
            rules = sh.get("conditionalFormats", [])
            to_delete = []
            for idx, rule in enumerate(rules):
                for rng in rule.get("ranges", []):
                    rs = rng.get("startRowIndex", 0)
                    re_ = rng.get("endRowIndex", 0)
                    if rs < row_end_0b and re_ > row_start_0b:
                        to_delete.append(idx)
                        break
            # Delete from highest index to lowest to keep indices valid
            for idx in sorted(to_delete, reverse=True):
                _api_retry(
                    service.spreadsheets().batchUpdate,
                    spreadsheetId=ssid,
                    body={"requests": [{"deleteConditionalFormatRule": {
                        "sheetId": sheet_id, "index": idx
                    }}]}
                )
            return
    except Exception as e:
        print(f"    ⚠  Could not clear existing conditional formats: {e}")


def apply_rfi_conditional_formatting(
        service, ssid: str, tinfo: dict,
        header_map: dict, data_start_1based: int) -> None:
    """
    Apply per-status conditional formatting to the table's data rows:
      OPENED                → Yellow fill on RFI No. cell
      APPROVED              → Green fill on RFI No. cell (white bold text)
      APPROVED WITH COMMENTS→ Orange fill on RFI No. cell
      NOT APPROVED          → Red fill on RFI No. cell (white bold text)
      CLOSED                → Near-black fill on RFI No. cell (white bold text)
      CANCELLED             → Strikethrough on the ENTIRE row (no fill change)

    The formula anchors on the Status column with an absolute column reference
    so the condition scales row-by-row across the full data range.
    """
    rfi_col_in_hm    = header_map.get("RFI No.")
    status_col_in_hm = header_map.get("Status")
    if rfi_col_in_hm is None or status_col_in_hm is None:
        return

    sheet_id       = tinfo["sheet_id"]
    tbl_start_col  = tinfo["start_col"]       # 0-based

    rfi_no_col_0b  = tbl_start_col + rfi_col_in_hm
    status_col_0b  = tbl_start_col + status_col_in_hm
    row_start_0b   = data_start_1based - 1    # 0-based inclusive
    row_end_0b     = tinfo["end_row"] - 1     # 0-based exclusive (stop before footer)
    tbl_end_col    = tinfo["end_col"]         # 0-based exclusive

    if row_end_0b <= row_start_0b:
        return

    # Clear any pre-existing rules that overlap our data range
    _delete_conditional_formats_in_range(service, ssid, sheet_id,
                                          row_start_0b, row_end_0b)

    status_letter = _col_letter(status_col_0b)
    formula_row   = data_start_1based   # 1-based anchor for formula

    def _rfi_range():
        return {"sheetId": sheet_id,
                "startRowIndex":    row_start_0b,
                "endRowIndex":      row_end_0b,
                "startColumnIndex": rfi_no_col_0b,
                "endColumnIndex":   rfi_no_col_0b + 1}

    def _row_range():
        return {"sheetId": sheet_id,
                "startRowIndex":    row_start_0b,
                "endRowIndex":      row_end_0b,
                "startColumnIndex": tbl_start_col,
                "endColumnIndex":   tbl_end_col}

    def _rgb(r, g, b):
        return {"red": r, "green": g, "blue": b}

    def _formula(status_val: str) -> str:
        return f'=${status_letter}{formula_row}="{status_val}"'

    requests = []

    # ── Rules for RFI No. cell (added at index 0 → later = higher priority) ──
    rfi_cell_rules = [
        # Lowest priority first (added first → overridden by later rules)
        ("APPROVED WITH COMMENTS",
            {"backgroundColor": _rgb(1.0, 0.65, 0.0),              # Orange
             "textFormat":       {"foregroundColor": _rgb(0,0,0)}}),
        ("NOT APPROVED",
            {"backgroundColor": _rgb(0.90, 0.18, 0.18),            # Red
             "textFormat":       {"foregroundColor": _rgb(1,1,1), "bold": True}}),
        ("APPROVED",
            {"backgroundColor": _rgb(0.20, 0.73, 0.35),            # Green
             "textFormat":       {"foregroundColor": _rgb(1,1,1), "bold": True}}),
        ("CLOSED",
            {"backgroundColor": _rgb(0.13, 0.13, 0.13),            # Near-black
             "textFormat":       {"foregroundColor": _rgb(1,1,1), "bold": True}}),
        ("OPENED",
            {"backgroundColor": _rgb(1.0, 0.95, 0.0),              # Yellow
             "textFormat":       {"foregroundColor": _rgb(0,0,0), "bold": False}}),
    ]

    for status_val, fmt in rfi_cell_rules:
        requests.append({
            "addConditionalFormatRule": {
                "rule": {
                    "ranges": [_rfi_range()],
                    "booleanRule": {
                        "condition": {
                            "type": "CUSTOM_FORMULA",
                            "values": [{"userEnteredValue": _formula(status_val)}],
                        },
                        "format": fmt,
                    }
                },
                "index": 0,
            }
        })

    # ── CANCELLED → strikethrough entire row ──────────────────────────────────
    requests.append({
        "addConditionalFormatRule": {
            "rule": {
                "ranges": [_row_range()],
                "booleanRule": {
                    "condition": {
                        "type": "CUSTOM_FORMULA",
                        "values": [{"userEnteredValue": _formula("CANCELLED")}],
                    },
                    "format": {
                        "textFormat": {"strikethrough": True},
                    }
                }
            },
            "index": 0,
        }
    })

    _api_retry(
        service.spreadsheets().batchUpdate,
        spreadsheetId=ssid,
        body={"requests": requests},
    )


def scan_and_rescue_orphans(service, ssid: str, tinfo: dict,
                             header_map: dict,
                             named_tables: dict = None) -> list:
    """
    Scan rows immediately BELOW the current table range for stray RFI entries
    (orphans placed outside the table by earlier script versions).

    KEY SAFEGUARDS:
      • Rows that fall inside ANOTHER named table on the same sheet are
        NEVER treated as orphans — they are skipped entirely.  This prevents
        the function from misidentifying a neighbouring table's data as orphans
        when two tables share the same sheet tab (e.g. SITE + LAB on one sheet).
      • All orphan deletions are submitted in a SINGLE batchUpdate request to
        avoid per-row write-quota exhaustion.
      • Scanning stops at the first row protected by a sibling table, so we
        don't read past table boundaries unnecessarily.

    Parameters:
        named_tables : the full {name: tinfo} dict returned by
                       get_all_named_tables().  Pass it in so we can build the
                       protected-row map.  If None, protection is disabled.

    Returns:
        List of raw row-lists rescued from below the table.
        The orphan rows are deleted from the sheet before returning.
    """
    sheet_title   = tinfo["sheet_title"]
    sheet_id      = tinfo["sheet_id"]
    table_end_row = tinfo["end_row"]        # 0-based exclusive end of THIS table
    scan_start_1b = table_end_row + 1       # 1-based first row to check

    # ── Build protected-row set from sibling tables on the same sheet ─────────
    # Any 0-based sheet row that belongs to another named table must not be
    # deleted or treated as an orphan.
    protected: set = set()
    if named_tables:
        for other_tinfo in named_tables.values():
            if other_tinfo["sheet_id"] == sheet_id and other_tinfo is not tinfo:
                for r in range(other_tinfo["start_row"], other_tinfo["end_row"]):
                    protected.add(r)

    # Scan all rows up to the nearest sibling table start (safety cap: 1000).
    # The protected-row set is the real boundary; the 1000-row cap is just a
    # safety net for sheets with no sibling tables at all.
    _MAX_ORPHAN_SCAN = 1000
    protected_ranges = []
    if named_tables:
        for other_tinfo in named_tables.values():
            if other_tinfo["sheet_id"] == sheet_id and other_tinfo is not tinfo:
                protected_ranges.append((other_tinfo["start_row"], other_tinfo["end_row"]))
    nearest_sibling = min(
        (s for s, _e in protected_ranges if s >= table_end_row),
        default=table_end_row + _MAX_ORPHAN_SCAN,
    )
    scan_rows = min(_MAX_ORPHAN_SCAN, nearest_sibling - table_end_row)

    if scan_rows <= 0:
        return []

    # Read TWO ranges in one call:
    #   (a) table-column-aligned range — used to extract clean data for rescue
    #   (b) wide range (A → BZ) — used to detect any misaligned orphan whose
    #       data was written into wrong columns by an older script version
    table_range = (
        f"'{sheet_title}'!"
        f"{col_letter(tinfo['start_col'])}{scan_start_1b}:"
        f"{col_letter(tinfo['end_col'] - 1)}{scan_start_1b + scan_rows - 1}"
    )
    # Wide scan: columns A–BZ (index 0–77) so we catch any horizontal offset
    wide_range = (
        f"'{sheet_title}'!"
        f"A{scan_start_1b}:BZ{scan_start_1b + scan_rows - 1}"
    )
    try:
        table_result = _api_retry(service.spreadsheets().values().get,
                                  spreadsheetId=ssid, range=table_range)
        wide_result  = _api_retry(service.spreadsheets().values().get,
                                  spreadsheetId=ssid, range=wide_range)
        all_rows      = table_result.get("values", [])
        wide_rows     = wide_result.get("values", [])
    except Exception as e:
        print(f"    ⚠  Could not scan below table for orphans: {e}")
        return []

    if not all_rows and not wide_rows:
        return []

    # Pad both lists to the same length
    max_rows = max(len(all_rows), len(wide_rows))
    while len(all_rows)  < max_rows: all_rows.append([])
    while len(wide_rows) < max_rows: wide_rows.append([])

    # Matches all known RFI number prefixes:
    #   IRIS-RFI-…   IR-RFI-…   IR-Q-RFI-…   IR-QAQC-RFI-…   IRIS-Q-RFI-…
    rfi_pattern = re.compile(
        r'\b(?:IRIS|IR)[-\s]?(?:(?:QAQC|Q)[-\s]?)?RFI', re.IGNORECASE
    )
    rfi_col_idx = header_map.get("RFI No.", 1)

    orphan_abs_0b = []   # 0-based sheet row indices of orphan rows
    orphan_rows   = []

    for i in range(max_rows):
        abs_0b    = table_end_row + i
        if abs_0b in protected:
            continue                       # belongs to a sibling table — skip

        t_row = all_rows[i]
        w_row = wide_rows[i]

        # Primary check: RFI No. column in the table-aligned row
        cell_val = t_row[rfi_col_idx] if rfi_col_idx < len(t_row) else ""
        found    = rfi_pattern.search(str(cell_val))

        # Fallback: scan EVERY cell in the wide row (catches misaligned orphans)
        if not found:
            for cell in w_row:
                if rfi_pattern.search(str(cell)):
                    found = True
                    break

        if found:
            orphan_abs_0b.append(abs_0b)
            # Use table-aligned row for rescue so column indices match
            orphan_rows.append(t_row)

    if not orphan_rows:
        return []

    print(f"    ⚠  Found {len(orphan_rows)} orphan row(s) below table "
          f"'{sheet_title}' — rescuing …")

    # Delete all orphans in ONE batchUpdate (= 1 write quota unit)
    try:
        _delete_rows_batch(service, ssid, sheet_id, orphan_abs_0b)
    except Exception as e:
        print(f"    ⚠  Could not delete orphan rows: {e}")

    return orphan_rows


# ══════════════════════════════════════════════════════════════════════════════
# AUDIT HELPERS  (scan existing entries, fix violations, reroute misplaced rows)
# ══════════════════════════════════════════════════════════════════════════════

# Maps Sheets column headers → rec dict keys
_COL_TO_KEY = {
    "RFI No.":              "rfi_no",
    "Date Received":        "date_received",
    "Inspection Date":      "inspection_date",
    "Road Code":            "road_code",
    "Road Name":            "road_name",
    "Island":               "island",
    "Contract Number":      "contract_number",
    "Start Station":        "start_station",
    "End Station":          "end_station",
    "Length Executed (km)": "length_km",
    "Merged Length (km)":   "merged_length_km",
    "Activity Name":        "activity_name",
    "Work Description":     "work_description",
    "Station Range":        "station_range",
    "Status":               "status",
    "RE's Remarks":         "re_remarks",
    "Remark Date":          "remark_date",
    "Attachment Path":      "attachment_path",
    "Project":              "project_km",
}


def _row_to_rec(row: list, header_map: dict, km_key: str) -> dict:
    """Convert a raw Sheets row list back to a rec dict."""
    rec = {}
    for col_name, rec_key in _COL_TO_KEY.items():
        idx = header_map.get(col_name)
        if idx is not None and idx < len(row):
            rec[rec_key] = str(row[idx]).strip()
        else:
            rec[rec_key] = ""
    if not rec.get("project_km"):
        rec["project_km"] = f"{km_key} KM"
    return rec


def _derive_rfi_type_from_rec(rec: dict) -> str:
    """
    Derive SITE or LAB from an existing rec dict (no ZIP/PDF context).
    Uses rfi_no + activity_name + work_description as detection signals.
    """
    rfi_no       = rec.get("rfi_no", "")
    activity     = rec.get("activity_name", "")
    work_desc    = rec.get("work_description", "")
    text_snippet = f"{activity} {work_desc}"
    return get_rfi_type(rfi_no, "SITE", text_snippet=text_snippet)


def _correct_table_name_for_rec(rec: dict, km_key: str) -> str:
    """Return the canonical table name this rec should live in."""
    island   = rec.get("island", "").strip().title()
    rfi_type = _derive_rfi_type_from_rec(rec)
    if island not in ("Pemba", "Unguja"):
        island = "Unguja"   # safe default
    return build_table_name(km_key, island, rfi_type)


def _read_data_rows_raw(service, ssid: str, tinfo: dict) -> list:
    """
    Read data rows (between header and footer, both exclusive).
    Returns list of (0based_sheet_row_index, raw_row_list) tuples.

    Layout (0-based):
      tinfo["start_row"]     = header row
      tinfo["start_row"] + 1 … tinfo["end_row"] - 2 = data rows
      tinfo["end_row"] - 1   = footer row (NEVER touched)
    """
    header_0b     = tinfo["start_row"]
    footer_0b     = tinfo["end_row"] - 1
    first_data_0b = header_0b + 1
    last_data_0b  = footer_0b - 1      # inclusive

    if last_data_0b < first_data_0b:
        return []

    sheet_title = tinfo["sheet_title"]
    start_col   = tinfo["start_col"]
    end_col     = _write_end_col(tinfo)   # read no further than col T
    first_1b    = first_data_0b + 1
    last_1b     = last_data_0b  + 1

    read_range = (
        f"'{sheet_title}'!"
        f"{col_letter(start_col)}{first_1b}:"
        f"{col_letter(end_col - 1)}{last_1b}"
    )
    try:
        result   = _api_retry(service.spreadsheets().values().get,
                               spreadsheetId=ssid, range=read_range)
        raw_rows = result.get("values", [])
    except Exception as e:
        print(f"    ⚠  Could not read data rows: {e}")
        return []

    return [(first_data_0b + offset, row) for offset, row in enumerate(raw_rows)]


def _update_row_in_place(service, ssid: str, tinfo: dict,
                         row_0based: int, row_data: list) -> None:
    """Overwrite a single row at the given 0-based sheet row index."""
    sheet_title = tinfo["sheet_title"]
    start_col   = tinfo["start_col"]
    end_col     = _write_end_col(tinfo)   # capped at col T
    row_1b      = row_0based + 1
    write_range = (
        f"'{sheet_title}'!"
        f"{col_letter(start_col)}{row_1b}:"
        f"{col_letter(end_col - 1)}{row_1b}"
    )
    _api_retry(
        service.spreadsheets().values().update,
        spreadsheetId=ssid,
        range=write_range,
        valueInputOption="USER_ENTERED",
        body={"values": [_cap_row(row_data)]}
    )


def _batch_update_rows_in_place(service, ssid: str, tinfo: dict,
                                 changes: list) -> None:
    """
    Write multiple corrected rows back in a SINGLE values().batchUpdate() call.
    This avoids per-row API quota exhaustion (60 writes/min limit).

    changes : list of (row_0based_index, row_data_list) pairs.

    One batchUpdate request counts as a single write quota unit regardless
    of how many ranges are included, so this is the safe pattern for bulk fixes.
    """
    if not changes:
        return

    sheet_title = tinfo["sheet_title"]
    start_col   = tinfo["start_col"]
    end_col     = _write_end_col(tinfo)   # capped at col T

    data = []
    for row_0b, row_data in changes:
        row_1b = row_0b + 1
        rng = (
            f"'{sheet_title}'!"
            f"{col_letter(start_col)}{row_1b}:"
            f"{col_letter(end_col - 1)}{row_1b}"
        )
        data.append({"range": rng, "values": [_cap_row(row_data)]})

    _api_retry(
        service.spreadsheets().values().batchUpdate,
        spreadsheetId=ssid,
        body={
            "valueInputOption": "USER_ENTERED",
            "data": data,
        }
    )


def _delete_rows_batch(service, ssid: str, sheet_id: int,
                       row_0b_indices: list) -> None:
    """
    Delete rows by 0-based sheet row index in a single batchUpdate.
    Sorted descending so higher indices are deleted first (no drift).
    """
    if not row_0b_indices:
        return
    requests = []
    for idx in sorted(set(row_0b_indices), reverse=True):
        requests.append({
            "deleteDimension": {
                "range": {
                    "sheetId":    sheet_id,
                    "dimension":  "ROWS",
                    "startIndex": idx,
                    "endIndex":   idx + 1,
                }
            }
        })
    _api_retry(
        service.spreadsheets().batchUpdate,
        spreadsheetId=ssid,
        body={"requests": requests}
    )
    time.sleep(0.5)


def renumber_sno_in_table(service, ssid: str, tinfo: dict,
                          header_map: dict) -> None:
    """
    Write sequential 1, 2, 3 … to the S/No. column for all data rows
    (between header and footer, exclusive).  Footer is never touched.
    """
    sno_col_in_hmap = header_map.get("S/No.")
    if sno_col_in_hmap is None:
        return

    data_rows = _read_data_rows_raw(service, ssid, tinfo)
    if not data_rows:
        return

    sheet_col_letter = col_letter(tinfo["start_col"] + sno_col_in_hmap)
    sheet_title      = tinfo["sheet_title"]
    first_row_0b     = data_rows[0][0]
    last_row_0b      = data_rows[-1][0]
    first_1b         = first_row_0b + 1
    last_1b          = last_row_0b  + 1

    values      = [[str(i + 1)] for i in range(len(data_rows))]
    write_range = (
        f"'{sheet_title}'!{sheet_col_letter}{first_1b}:"
        f"{sheet_col_letter}{last_1b}"
    )
    _api_retry(
        service.spreadsheets().values().update,
        spreadsheetId=ssid,
        range=write_range,
        valueInputOption="USER_ENTERED",
        body={"values": values}
    )
    print(f"    ✓ S/No. renumbered 1–{len(data_rows)}")


def audit_and_fix_table(service, ssid: str, tinfo: dict, tbl_name: str,
                        header_map: dict, km_key: str) -> dict:
    """
    Scan every data row in the table and:
      1. Detect misrouted entries (wrong island / LAB-vs-SITE routing).
         → Remove them from this table; queue for re-insertion elsewhere.
      2. Fix in-place violations:
         – Non-canonical Status value → canonicalize and rewrite the row.
         – Blank Road Name → fall back to Road Code so the field is never empty.
      3. S/No. renumbering is handled separately after all deletes.

    Returns:
        {dest_tbl_name: [rec_dicts]}  — records to be re-inserted in their
        correct table during the subsequent insert pass.

    Footer row (tinfo["end_row"] - 1) is NEVER read, written, or deleted.
    """
    rfi_col        = header_map.get("RFI No.")
    status_col     = header_map.get("Status")
    road_col       = header_map.get("Road Name")
    code_col       = header_map.get("Road Code")
    island_col     = header_map.get("Island")
    attachment_col = header_map.get("Attachment Path")
    activity_col   = header_map.get("Activity Name")
    work_col       = header_map.get("Work Description")

    data_rows = _read_data_rows_raw(service, ssid, tinfo)
    if not data_rows:
        return {}

    reroute_buffer  = defaultdict(list)
    to_delete_0b    = []
    pending_fixes   = []   # (row_0based, fixed_row) — batched, NOT written per-row

    for sheet_row_0b, raw_row in data_rows:
        # Skip completely blank rows
        rfi_val = ""
        if rfi_col is not None and rfi_col < len(raw_row):
            rfi_val = str(raw_row[rfi_col]).strip()
        if not rfi_val:
            continue

        # Convert row to rec dict for routing checks
        rec = _row_to_rec(raw_row, header_map, km_key)

        # ── Check routing ──────────────────────────────────────────────────
        correct_tbl = _correct_table_name_for_rec(rec, km_key)
        if correct_tbl != tbl_name:
            reroute_buffer[correct_tbl].append(rec)
            to_delete_0b.append(sheet_row_0b)
            print(f"    ⚠  Misrouted: {rfi_val!r} → moving to '{correct_tbl}'")
            continue

        # ── In-place fixes — collect only, write later in one batch ───────
        n_cols    = tinfo["end_col"] - tinfo["start_col"]
        fixed_row = list(raw_row)
        while len(fixed_row) < n_cols:
            fixed_row.append("")
        row_changed = False

        # Fix 0: CANCELLED when RFI island indicator (P/U) mismatches the
        #        road's actual island stored in the Island column.
        #   • P in RFI No. but Island = Unguja  → CANCELLED
        #   • U in RFI No. but Island = Pemba   → CANCELLED
        if status_col is not None and status_col < len(fixed_row):
            rfi_island = island_from_rfi_no(rfi_val)   # "Pemba" | "Unguja" | ""
            if rfi_island:
                row_island = ""
                if island_col is not None and island_col < len(fixed_row):
                    row_island = str(fixed_row[island_col]).strip().title()
                if row_island and row_island != rfi_island:
                    if str(fixed_row[status_col]).strip() != "CANCELLED":
                        fixed_row[status_col] = "CANCELLED"
                        row_changed = True

        # Fix 1: Status canonicalization (only if not already forced to CANCELLED above)
        if status_col is not None and status_col < len(fixed_row):
            raw_status   = str(fixed_row[status_col]).strip()
            if raw_status != "CANCELLED":     # don't overwrite island-mismatch cancel
                canon_status = canonicalize_status(raw_status)
                if canon_status != raw_status:
                    fixed_row[status_col] = canon_status
                    row_changed = True

        # Fix 2: Road Code + Road Name — alignment codes resolved, no chainage,
        # never blank.
        #
        # Alignment designators (AL1 … AL9) stored in Road Code are NOT valid
        # road codes.  When detected, we fuzzy-match the existing road name
        # against ROAD_CODE_TO_NAME (≥ 80 % similarity) to find the real code
        # and the canonical road name, then rewrite BOTH cells.
        #
        # Name resolution priority (after any alignment-code handling):
        #   1. ROAD_CODE_TO_NAME exact lookup by road code
        #   2. Re-read road name from the original ZIP attachment (blank names only)
        #   3. clean_road_name() to strip chainage from what is in the cell
        #   4. Google Sheets VLOOKUP formula as live fallback
        if road_col is not None and road_col < len(fixed_row):
            raw_road = str(fixed_row[road_col]).strip()
            road_code_val = (str(fixed_row[code_col]).strip()
                             if code_col is not None and code_col < len(fixed_row) else "")

            # ── Step A: resolve alignment codes via fuzzy road-name matching ──
            if is_alignment_code(road_code_val):
                matched_code, matched_name = fuzzy_road_match(raw_road, threshold=0.80)
                if matched_code:
                    # Overwrite Road Code cell with the real code
                    if code_col is not None and code_col < len(fixed_row):
                        fixed_row[code_col] = matched_code
                        row_changed = True
                    road_code_val = matched_code
                    # If road name was already correct keep it; otherwise use canonical
                    if matched_name and matched_name != raw_road:
                        fixed_row[road_col] = matched_name
                        row_changed = True
                    raw_road = matched_name or raw_road
                else:
                    # No fuzzy match — clear the spurious AL code so the VLOOKUP
                    # formula can try later, and keep road name as-is
                    if code_col is not None and code_col < len(fixed_row):
                        fixed_row[code_col] = ""
                        row_changed = True
                    road_code_val = ""

            # ── Step B: resolve road name from real code (or remaining blank) ─
            canonical = lookup_road_name(road_code_val) if road_code_val else ""

            if canonical:
                target_road = canonical
            elif not raw_road:
                # Blank — try to recover from original ZIP attachment first
                attach_val = (str(fixed_row[attachment_col]).strip()
                              if attachment_col is not None and attachment_col < len(fixed_row)
                              else "")
                recovered = road_name_from_attachment(attach_val) if attach_val else ""
                if recovered:
                    # Try fuzzy-matching recovered name too, in case attachment
                    # gives a name we can match to a real code
                    if not road_code_val:
                        m_code, m_name = fuzzy_road_match(recovered, threshold=0.80)
                        if m_code:
                            if code_col is not None and code_col < len(fixed_row):
                                fixed_row[code_col] = m_code
                                row_changed = True
                            recovered = m_name or recovered
                    target_road = recovered
                elif road_code_val:
                    # Live Sheets VLOOKUP formula — auto-fills from lookup tab
                    code_col_letter = col_letter(tinfo["start_col"] + code_col)
                    row_1b = sheet_row_0b + 1
                    target_road = (
                        f"=IFERROR(VLOOKUP({code_col_letter}{row_1b},"
                        f"'{ROAD_CODE_SHEET_NAME}'!$A:$B,2,FALSE),"
                        f'IF({code_col_letter}{row_1b}<>"",{code_col_letter}{row_1b},""))'
                    )
                else:
                    target_road = ""
            else:
                target_road = clean_road_name(raw_road)

            if target_road and target_road != str(fixed_row[road_col]).strip():
                fixed_row[road_col] = target_road
                row_changed = True

        # Fix 3: Island — fill from RFI No. if blank
        if island_col is not None and island_col < len(fixed_row):
            if not str(fixed_row[island_col]).strip():
                derived_island = island_from_rfi_no(rfi_val)
                if derived_island:
                    fixed_row[island_col] = derived_island
                    row_changed = True

        # Fix 4: Activity Name — fill from Work Description if blank
        if activity_col is not None and activity_col < len(fixed_row):
            if not str(fixed_row[activity_col]).strip():
                wd = (str(fixed_row[work_col]).strip()
                      if work_col is not None and work_col < len(fixed_row) else "")
                derived_act = normalise_activity(wd) if wd else "OTHER"
                fixed_row[activity_col] = derived_act
                row_changed = True

        if row_changed:
            pending_fixes.append((sheet_row_0b, fixed_row))

    # ── Flush all in-place fixes in ONE batchUpdate (= 1 write quota unit) ─
    if pending_fixes:
        try:
            _batch_update_rows_in_place(service, ssid, tinfo, pending_fixes)
            print(f"    ✓ In-place fixes applied: {len(pending_fixes)} row(s)")
        except Exception as e:
            print(f"    ⚠  Batch in-place fix failed: {e}")

    # ── Delete misrouted rows in one batch request ──────────────────────────
    if to_delete_0b:
        print(f"    ✓ Removing {len(to_delete_0b)} misrouted row(s) from '{tbl_name}' …")
        try:
            _delete_rows_batch(service, ssid, tinfo["sheet_id"], to_delete_0b)
        except Exception as e:
            print(f"    ⚠  Batch delete failed: {e}")

    return dict(reroute_buffer)


# ══════════════════════════════════════════════════════════════════════════════
# SPREADSHEET LOOKUP
# ══════════════════════════════════════════════════════════════════════════════

def get_spreadsheet_id(drive, name: str, id_file: str) -> str:
    # Explicit project IDs take precedence over stale caches and duplicate names.
    if name in RFI_ID_BY_NAME:
        return RFI_ID_BY_NAME[name]

    if os.path.exists(id_file):
        with open(id_file) as f:
            ssid = f.read().strip()
        if ssid:
            return ssid

    res   = drive.files().list(
        q=(f"name='{name}' and "
           f"mimeType='application/vnd.google-apps.spreadsheet' and trashed=false"),
        fields="files(id,name)"
    ).execute()
    files = res.get("files", [])
    if not files:
        raise ValueError(
            f"Spreadsheet '{name}' not found in Google Drive.\n"
            f"Please create it and ensure the named tables exist."
        )
    ssid = files[0]["id"]
    try:
        with open(id_file, "w") as f:
            f.write(ssid)
    except Exception:
        pass
    print(f"  Found '{name}' → {ssid}")
    return ssid


def _is_connection_err(e: Exception) -> bool:
    """Return True when *e* is a transient TCP / socket error."""
    if isinstance(e, (OSError, ConnectionAbortedError, ConnectionResetError,
                      ConnectionError, BrokenPipeError)):
        return True
    s = str(e).lower()
    return any(tag in s for tag in (
        "winerror 10053", "winerror 10054", "winerror 10061",
        "connection aborted", "connection reset", "connection refused",
        "remotedisconnected", "broken pipe", "errno 104", "errno 111",
    ))


def _share_once(drive, ssid: str, email: str,
                notify: bool, max_attempts: int = 4) -> None:
    """
    Call drive.permissions().create() for *email* with up to *max_attempts*
    retries on transient connection errors.  Raises on non-connection errors
    (caller decides how to handle them).
    """
    for attempt in range(max_attempts):
        try:
            drive.permissions().create(
                fileId=ssid,
                body={"type": "user", "role": "writer", "emailAddress": email},
                sendNotificationEmail=notify,
                fields="id"
            ).execute()
            return   # success
        except Exception as e:
            if _is_connection_err(e) and attempt < max_attempts - 1:
                wait = min(10 * (attempt + 1), 30)
                print(f"    ⏳ Connection error sharing with {email} — "
                      f"retrying in {wait}s (attempt {attempt + 2}/{max_attempts}) …")
                time.sleep(wait)
            else:
                raise   # non-connection error OR last attempt


# Domains that use Google as their mail backend and can receive shares silently.
_GOOGLE_DOMAINS = {"gmail.com", "googlemail.com", "google.com"}

def _needs_notify(email: str) -> bool:
    """Return True when the address is NOT a Google-managed inbox.

    Non-Google addresses (yahoo.com, icloud.com, hotmail.com, etc.) require
    sendNotificationEmail=True or the Drive API raises InvalidSharingRequest.
    Sending to them directly with notify=True avoids the wasted first attempt.
    """
    domain = email.lower().split("@")[-1] if "@" in email else ""
    return domain not in _GOOGLE_DOMAINS


def share_spreadsheet(drive, ssid: str, emails: list) -> None:
    print(f"  Sharing with {len(emails)} collaborators …")
    for email in emails:
        # Choose the right notify flag upfront — avoids an always-failing first attempt
        # for non-Google addresses (Yahoo, iCloud, Outlook, etc.)
        use_notify = _needs_notify(email)
        try:
            _share_once(drive, ssid, email, notify=use_notify)
            status = "with notification" if use_notify else "silently"
            print(f"    ✓ Shared {status}: {email}")
            time.sleep(0.3)
        except Exception as e:
            err_str = str(e).lower()
            # If silent share failed (shouldn't happen for Gmail but just in case),
            # fall back to notification share — no retry message, no warning.
            if not use_notify and ("invalidsharingrequest" in err_str
                                   or "no google account" in err_str):
                try:
                    _share_once(drive, ssid, email, notify=True)
                    print(f"    ✓ Shared with notification (fallback): {email}")
                    time.sleep(0.3)
                except Exception as e2:
                    print(f"    ⚠  Could not share with {email}: {e2}")
            else:
                print(f"    ⚠  Could not share with {email}: {e}")
    print(f"  ✓ Sharing complete")


# ══════════════════════════════════════════════════════════════════════════════
# POST-PROCESS ROAD NAMES  (apply canonical names from Physical Progress lookup)
# ══════════════════════════════════════════════════════════════════════════════

def lookup_road_name(road_code: str) -> str:
    """
    Return the canonical road name for a road code, checking ROAD_CODE_TO_NAME.
    Returns empty string if not found.
    road_code comparison is case-insensitive (lookup keys are upper-cased at call site).
    """
    # Try exact match first, then upper-cased
    return (ROAD_CODE_TO_NAME.get(road_code)
            or ROAD_CODE_TO_NAME.get(road_code.upper(), ""))


def is_alignment_code(code: str) -> bool:
    """
    Return True when the value in Road Code is actually an alignment designator
    (AL1, AL2, AL-3, AL 4 …) rather than a real road code like MK-235 or SO-212.
    Alignment designators are never valid road codes — they appear only as part
    of a road name and were incorrectly captured from the PDF.
    """
    return bool(re.match(r'^AL\s*[-]?\s*\d+$', code.strip(), re.IGNORECASE))


def fuzzy_road_match(road_name: str,
                     threshold: float = 0.80) -> tuple:
    """
    Find the best matching (road_code, canonical_name) from ROAD_CODE_TO_NAME
    using difflib SequenceMatcher similarity on the road name.

    Returns (road_code, canonical_name) if best similarity >= threshold,
    otherwise ("", "").

    The 80 % default threshold avoids false positives while still catching
    the common case where the road name in the sheet differs only in spacing,
    capitalisation, or the trailing alignment tag (e.g. "(AL2)").
    """
    from difflib import SequenceMatcher
    rn_lower = road_name.lower().strip()
    if not rn_lower:
        return "", ""
    best_ratio = 0.0
    best_code  = ""
    best_name  = ""
    for code, name in ROAD_CODE_TO_NAME.items():
        ratio = SequenceMatcher(None, rn_lower, name.lower()).ratio()
        if ratio > best_ratio:
            best_ratio = ratio
            best_code  = code
            best_name  = name
    if best_ratio >= threshold:
        return best_code, best_name
    return "", ""


def apply_road_lookup(records: list, road_lookup: dict) -> None:
    """
    Overwrite road_name with the canonical (chainage-stripped) name.
    Priority order:
      1. Physical Progress sheets (road_lookup dict keyed by (km_key, road_code))
      2. ROAD_CODE_TO_NAME hardcoded lookup keyed by road_code
      3. clean_road_name() applied to whatever was extracted from the PDF
    """
    phys_updated = 0
    static_updated = 0
    fuzzy_updated = 0
    cleaned = 0
    for rec in records:
        km_key    = rec.get("project_km", "").replace(" KM", "").strip()
        road_code = rec.get("road_code", "").strip()

        # ── Alignment code detected — clear it and try fuzzy road-name match ──
        if is_alignment_code(road_code):
            m_code, m_name = fuzzy_road_match(rec.get("road_name", ""), threshold=0.80)
            if m_code:
                rec["road_code"] = m_code
                rec["road_name"] = m_name
                road_code = m_code
                fuzzy_updated += 1
                continue
            else:
                rec["road_code"] = ""
                road_code = ""

        key = (km_key, road_code.upper())

        if road_code and key in road_lookup:
            rec["road_name"] = clean_road_name(road_lookup[key])
            phys_updated += 1
        elif road_code:
            static = lookup_road_name(road_code)
            if static:
                rec["road_name"] = static
                static_updated += 1
            elif rec.get("road_name"):
                rec["road_name"] = clean_road_name(rec["road_name"])
                cleaned += 1
            else:
                # Try fuzzy on road name even if no code
                m_code, m_name = fuzzy_road_match(rec.get("road_name", ""), threshold=0.80)
                if m_code:
                    rec["road_code"] = m_code
                    rec["road_name"] = m_name
                    fuzzy_updated += 1
                else:
                    rec["road_name"] = road_code
        else:
            rn = rec.get("road_name", "")
            if rn:
                cleaned_rn = clean_road_name(rn)
                # Try fuzzy even on a name-only record
                m_code, m_name = fuzzy_road_match(cleaned_rn, threshold=0.80)
                if m_code:
                    rec["road_code"] = m_code
                    rec["road_name"] = m_name
                    fuzzy_updated += 1
                else:
                    rec["road_name"] = cleaned_rn
                    cleaned += 1

    total = phys_updated + static_updated + fuzzy_updated + cleaned
    if phys_updated:
        print(f"  ✓ Road names from Physical Progress sheets: {phys_updated}")
    if static_updated:
        print(f"  ✓ Road names from built-in lookup: {static_updated}")
    if fuzzy_updated:
        print(f"  ✓ Road names resolved via fuzzy match (≥80%): {fuzzy_updated}")
    if cleaned:
        print(f"  ✓ Road names cleaned (chainage stripped): {cleaned}")
    if not total:
        print(f"  ✓ Road names: no changes needed")


# ══════════════════════════════════════════════════════════════════════════════
# ROAD NAME FROM ATTACHMENT  (re-read PDF inside ZIP to recover a blank road name)
# ══════════════════════════════════════════════════════════════════════════════

def road_name_from_attachment(attachment_path: str) -> str:
    """
    Given an attachment_path like
        <RFI_ZIP_DIR>/some.zip/IRIS-Q-RFI-P-25-0019.pdf
    open the ZIP (using ZIP_FOLDER as the local root) and extract the road name
    from the embedded PDF/DOCX.  Returns "" on any failure.
    """
    if not attachment_path:
        return ""
    # Normalise path separators
    ap = attachment_path.replace("\\", "/")
    # Find the .zip component
    zip_split = re.split(r'\.zip/', ap, maxsplit=1, flags=re.IGNORECASE)
    if len(zip_split) == 2:
        # Part before .zip gives the folder+zipname; part after is the member
        zip_tail  = zip_split[0].split("/")[-1] + ".zip"
        member    = zip_split[1]
        zip_fpath = os.path.join(ZIP_FOLDER, zip_tail)
        try:
            with zipfile.ZipFile(zip_fpath, "r") as zf:
                # member name may have sub-path — try basename match too
                names = zf.namelist()
                match = next((n for n in names
                              if n == member or os.path.basename(n) == os.path.basename(member)),
                             None)
                if not match:
                    return ""
                data = zf.read(match)
                text = (extract_text_pdf(data) if match.lower().endswith(".pdf")
                        else extract_text_docx(data))
                # Extract just the road name from the text
                raw = _clean(
                    _ex(r"ROAD\s+NAME[:\s]+([^\n\r]{3,120})", text) or
                    _ex(r"ROAD[:\s]+([^\n\r]{3,120})", text)
                )
                return clean_road_name(raw)
        except Exception:
            return ""
    else:
        # Standalone file
        fname = ap.split("/")[-1]
        fpath = os.path.join(ZIP_FOLDER, fname)
        if not os.path.exists(fpath):
            return ""
        try:
            with open(fpath, "rb") as f:
                data = f.read()
            ext  = os.path.splitext(fname)[1].lower()
            text = extract_text_pdf(data) if ext == ".pdf" else extract_text_docx(data)
            raw  = _clean(
                _ex(r"ROAD\s+NAME[:\s]+([^\n\r]{3,120})", text) or
                _ex(r"ROAD[:\s]+([^\n\r]{3,120})", text)
            )
            return clean_road_name(raw)
        except Exception:
            return ""


# ══════════════════════════════════════════════════════════════════════════════
# ROAD CODE & ROAD NAMES SHEET  (canonical lookup tab in each spreadsheet)
# ══════════════════════════════════════════════════════════════════════════════

ROAD_CODE_SHEET_NAME = "Road Code & Road Names"

def create_road_code_sheet(service, ssid: str) -> None:
    """
    Create (or refresh) a sheet tab called 'Road Code & Road Names' in the
    spreadsheet at ssid.  The tab is populated from ROAD_CODE_TO_NAME.
    If the tab already exists its content is overwritten; otherwise it is
    added as a new sheet.
    """
    meta = _api_retry(service.spreadsheets().get, spreadsheetId=ssid)
    existing_sheets = {s["properties"]["title"]: s["properties"]["sheetId"]
                       for s in meta.get("sheets", [])}

    if ROAD_CODE_SHEET_NAME not in existing_sheets:
        # Add the sheet
        resp = _api_retry(
            service.spreadsheets().batchUpdate,
            spreadsheetId=ssid,
            body={"requests": [{"addSheet": {"properties": {"title": ROAD_CODE_SHEET_NAME}}}]},
        )
        sheet_id = resp["replies"][0]["addSheet"]["properties"]["sheetId"]
        print(f"    ✓ Created sheet tab '{ROAD_CODE_SHEET_NAME}'")
    else:
        print(f"    ↩  Sheet tab '{ROAD_CODE_SHEET_NAME}' already exists — refreshing …")

    # Write header + all road codes
    rows = [["Road Code", "Road Name"]]
    for code in sorted(ROAD_CODE_TO_NAME):
        rows.append([code, ROAD_CODE_TO_NAME[code]])

    range_a1 = f"'{ROAD_CODE_SHEET_NAME}'!A1:B{len(rows)}"
    _api_retry(
        service.spreadsheets().values().update,
        spreadsheetId=ssid,
        range=range_a1,
        body={"values": rows},
        valueInputOption="USER_ENTERED",
    )
    print(f"    ✓ '{ROAD_CODE_SHEET_NAME}' written: {len(rows) - 1} road codes")


# ══════════════════════════════════════════════════════════════════════════════
# REMOVE EXTRA COLUMNS  (e.g. the legacy 'Table' column)
# ══════════════════════════════════════════════════════════════════════════════

def remove_columns_from_all_tables(service, ssid: str,
                                   named_tables: dict,
                                   col_names: list) -> bool:
    """
    Delete each column named in col_names from every table in the spreadsheet.
    Deletions are sorted descending so earlier deletes do not shift later indices.
    Returns True if any column was deleted (caller should re-read named_tables).
    """
    # Collect (sheet_id, abs_col_0b) pairs to delete — deduplicate per sheet
    to_delete: dict[int, set] = {}   # sheet_id -> set of absolute 0-based col indices

    for tbl_name, tinfo in named_tables.items():
        # Fetch the header row for this table to find the col
        header_range = (f"'{tinfo['sheet_title']}'!"
                        f"{col_letter(tinfo['start_col'])}{tinfo['start_row'] + 1}:"
                        f"{col_letter(tinfo['end_col'] - 1)}{tinfo['start_row'] + 1}")
        try:
            res = _api_retry(
                service.spreadsheets().values().get,
                spreadsheetId=ssid, range=header_range
            )
            header = [str(v).strip() for v in (res.get("values") or [[]])[0]]
        except Exception:
            continue
        header_lower = [h.lower() for h in header]
        for col_name in col_names:
            # Case-insensitive match so "TABLE" / "Table" / "table" all work
            target_lower = col_name.lower()
            if target_lower in header_lower:
                rel_idx = header_lower.index(target_lower)
                abs_idx = tinfo["start_col"] + rel_idx
                sid     = tinfo["sheet_id"]
                to_delete.setdefault(sid, set()).add(abs_idx)

    if not to_delete:
        return False

    requests = []
    for sheet_id, cols in to_delete.items():
        for col_abs in sorted(cols, reverse=True):   # descending to avoid drift
            requests.append({
                "deleteDimension": {
                    "range": {
                        "sheetId":    sheet_id,
                        "dimension":  "COLUMNS",
                        "startIndex": col_abs,
                        "endIndex":   col_abs + 1,
                    }
                }
            })
    if requests:
        _api_retry(service.spreadsheets().batchUpdate,
                   spreadsheetId=ssid, body={"requests": requests})
        deleted_names = ", ".join(f"'{c}'" for c in col_names)
        print(f"    ✓ Deleted legacy column(s) {deleted_names} from spreadsheet")
    return bool(requests)


def push_status_updates_to_sheet(
        service, ssid: str, tinfo: dict, header_map: dict,
        data_start_1based: int, status_lookup: dict) -> int:
    """
    Read every data row in the named table and, for any row whose RFI No.
    appears in status_lookup, overwrite the Status / RE's Remarks / Remark Date
    cells with the authoritative values from the status-summary PDFs.

    Returns the count of rows patched.
    """
    if not status_lookup:
        return 0

    rfi_col   = header_map.get("RFI No.")
    stat_col  = header_map.get("Status")
    rem_col   = header_map.get("RE's Remarks")
    rdate_col = header_map.get("Remark Date")

    if rfi_col is None or stat_col is None:
        return 0

    sheet_title  = tinfo["sheet_title"]
    start_col_0b = tinfo["start_col"]
    end_col_0b   = tinfo["end_col"]

    row_start_1b = data_start_1based
    row_end_1b   = tinfo["end_row"]    # last row of table (footer) — exclusive

    if row_end_1b <= row_start_1b:
        return 0

    # Read all data rows
    rng = (f"'{sheet_title}'!"
           f"{col_letter(start_col_0b)}{row_start_1b}:"
           f"{col_letter(end_col_0b - 1)}{row_end_1b - 1}")
    try:
        result = _api_retry(
            service.spreadsheets().values().get,
            spreadsheetId=ssid, range=rng
        )
    except Exception as e:
        print(f"    ⚠  push_status_updates: could not read rows — {e}")
        return 0

    rows_data = result.get("values", [])
    if not rows_data:
        return 0

    changes = []   # (row_0based_in_sheet, full_row_list)
    n_cols  = end_col_0b - start_col_0b

    for i, row_vals in enumerate(rows_data):
        # Pad row to full table width
        row_vals = list(row_vals) + [""] * (n_cols - len(row_vals))

        try:
            rfi_no = row_vals[rfi_col].strip().upper()
        except (IndexError, AttributeError):
            continue

        if not rfi_no or rfi_no not in status_lookup:
            continue

        info   = status_lookup[rfi_no]
        changed = False

        if info.get("status"):
            current_status = row_vals[stat_col].strip().upper() if stat_col < len(row_vals) else ""
            # Only overwrite if status is empty or OPENED (authoritative PDF wins)
            if not current_status or current_status == "OPENED":
                row_vals[stat_col] = info["status"]
                changed = True

        if rem_col is not None and info.get("re_remarks"):
            if rem_col < len(row_vals) and not row_vals[rem_col].strip():
                row_vals[rem_col] = info["re_remarks"]
                changed = True

        if rdate_col is not None and info.get("remark_date"):
            if rdate_col < len(row_vals) and not row_vals[rdate_col].strip():
                row_vals[rdate_col] = info["remark_date"]
                changed = True

        if changed:
            sheet_row_0b = row_start_1b - 1 + i    # 0-based sheet row
            changes.append((sheet_row_0b, row_vals))

    if changes:
        _batch_update_rows_in_place(service, ssid, tinfo, changes)

    return len(changes)


# ══════════════════════════════════════════════════════════════════════════════
# PROCESS ONE SPREADSHEET  (insert-above-footer, table-aware)
# ══════════════════════════════════════════════════════════════════════════════

def process_spreadsheet(service, drive, name: str,
                        project_filter: str, km_key: str,
                        emails: list, all_records: list,
                        status_lookup: dict | None = None) -> int:
    """
    3-pass pipeline for each RFI Detailed Log spreadsheet:

    PASS 1 — AUDIT
      • Rescue orphan rows below each table (move them inside).
      • Scan all existing data rows; fix in-place violations
        (non-canonical Status, blank Road Name).
      • Detect misrouted entries (wrong island, QAQC in SITE table, etc.)
        and queue them for re-insertion in the correct table.

    PASS 2 — RENUMBER
      • After deletions from the audit pass, rewrite S/No. as 1, 2, 3 …
        in every table so no gaps or duplicates remain.

    PASS 3 — INSERT
      • Combine freshly-parsed records + rerouted records from Pass 1.
      • Deduplicate by RFI No. against entries already in the table.
      • Insert all new rows ABOVE the footer row.
      • Apply/refresh Status dropdown validation on every table.
      • Renumber S/No. one final time to incorporate the new rows.

    Footer row (tinfo["end_row"] - 1) is NEVER read, written, or deleted
    at any point in any of the three passes.

    Returns total number of rows inserted (new + rerouted).
    """
    print(f"\n{'─'*70}")
    print(f"  Spreadsheet: {name}")
    print(f"{'─'*70}")

    id_file = rf"C:\AI-Projects\rfi_spreadsheet_{km_key.replace('.', '_')}.txt"
    ssid    = get_spreadsheet_id(drive, name, id_file)
    print(f"  URL: https://docs.google.com/spreadsheets/d/{ssid}")

    # ── Ensure 'Road Code & Road Names' lookup sheet exists ─────────────────
    print(f"\n  Setting up '{ROAD_CODE_SHEET_NAME}' lookup tab …")
    try:
        create_road_code_sheet(service, ssid)
    except PermissionError:
        return 0   # banner already printed
    except Exception as e:
        print(f"    ⚠  Could not create Road Code sheet: {e}")

    named_tables = get_all_named_tables(service, ssid)
    if named_tables:
        print(f"\n  Named tables found: {list(named_tables.keys())}")
    else:
        print(f"\n  ⚠  No named tables found. Ensure you have converted the "
              f"data ranges to Tables via Format → Convert to table in Google Sheets.")

    # ── Remove legacy columns (e.g. 'Table') from all tables ────────────────
    print(f"\n  Checking for legacy columns to remove …")
    try:
        if remove_columns_from_all_tables(service, ssid, named_tables, ["Table"]):
            # Column indices may have shifted — re-read table metadata
            named_tables = get_all_named_tables(service, ssid)
            print(f"  Named tables refreshed after column deletion.")
    except Exception as e:
        print(f"    ⚠  Column removal check failed: {e}")

    # Filter records for this project
    project_recs = [r for r in all_records
                    if project_filter in r.get("project_km", "")]

    routing = [
        ("Unguja", "SITE"),
        ("Unguja", "LAB"),
        ("Pemba",  "SITE"),
        ("Pemba",  "LAB"),
    ]

    # Accumulates rerouted recs across all tables: {dest_tbl_name: [rec_dicts]}
    global_reroute: dict = defaultdict(list)

    # ── Early permission probe ───────────────────────────────────────────────
    # Write one harmless value to cell A1 of the spreadsheet metadata range
    # so we catch a stale/read-only token BEFORE processing any tables.
    # (We immediately overwrite with the same value, so nothing changes.)
    try:
        _probe_range = "A1"
        _probe_val   = _api_retry(
            service.spreadsheets().values().get,
            spreadsheetId=ssid,
            range=_probe_range,
        )
        _probe_first = (_probe_val.get("values") or [[""]])[0]
        _api_retry(
            service.spreadsheets().values().update,
            spreadsheetId=ssid,
            range=_probe_range,
            valueInputOption="USER_ENTERED",
            body={"values": [_probe_first]},
        )
    except PermissionError:
        # The banner was already printed by _api_retry — nothing more to do.
        return 0
    except Exception:
        # Any other probe failure (e.g. empty sheet) is non-fatal; continue.
        pass

    # ══════════════════════════════════════════════════════════════════════════
    # PASS 1 — AUDIT
    # ══════════════════════════════════════════════════════════════════════════
    print(f"\n  ── Pass 1: Audit existing entries ──────────────────────────────")
    for island, rfi_type in routing:
        tbl_name = build_table_name(km_key, island, rfi_type)
        if tbl_name not in named_tables:
            continue

        tinfo = named_tables[tbl_name]
        header_map, _, _, _ = get_table_existing_data(service, ssid, tinfo)
        if not header_map:
            continue

        print(f"\n  Auditing '{tbl_name}' …")

        # Rescue below-table orphans first (pass named_tables so sibling
        # tables on the same sheet tab are not mistaken for orphans)
        orphan_rows = scan_and_rescue_orphans(
            service, ssid, tinfo, header_map, named_tables=named_tables
        )
        if orphan_rows:
            named_tables = get_all_named_tables(service, ssid)
            tinfo        = named_tables.get(tbl_name, tinfo)

        reroute = audit_and_fix_table(service, ssid, tinfo, tbl_name,
                                      header_map, km_key)
        for dest_tbl, recs in reroute.items():
            global_reroute[dest_tbl].extend(recs)

        # Re-fetch table metadata if rows were deleted
        if reroute or orphan_rows:
            named_tables = get_all_named_tables(service, ssid)
            time.sleep(0.3)

    # ══════════════════════════════════════════════════════════════════════════
    # PASS 2 — RENUMBER S/No. after audit deletions
    # ══════════════════════════════════════════════════════════════════════════
    print(f"\n  ── Pass 2: Renumber S/No. ──────────────────────────────────────")
    named_tables = get_all_named_tables(service, ssid)
    for island, rfi_type in routing:
        tbl_name = build_table_name(km_key, island, rfi_type)
        if tbl_name not in named_tables:
            continue
        tinfo = named_tables[tbl_name]
        header_map, _, _, data_count = get_table_existing_data(service, ssid, tinfo)
        if not header_map or data_count == 0:
            continue
        print(f"\n  Renumbering '{tbl_name}' ({data_count} rows) …")
        try:
            renumber_sno_in_table(service, ssid, tinfo, header_map)
        except PermissionError:
            return 0          # banner already printed
        except Exception as e:
            print(f"    ⚠  Renumber failed: {e}")
        time.sleep(0.3)

    # ══════════════════════════════════════════════════════════════════════════
    # PASS 3 — INSERT new & rerouted records
    # ══════════════════════════════════════════════════════════════════════════
    print(f"\n  ── Pass 3: Insert new & rerouted records ───────────────────────")
    named_tables   = get_all_named_tables(service, ssid)
    total_inserted = 0

    for island, rfi_type in routing:
        tbl_name = build_table_name(km_key, island, rfi_type)

        if tbl_name not in named_tables:
            print(f"\n  ⚠  Table not found: '{tbl_name}' — skipping.")
            continue

        tinfo = named_tables[tbl_name]
        header_map, existing_rfis, max_sno, _ = get_table_existing_data(
            service, ssid, tinfo
        )

        # Freshly-parsed records from ZIPs for this island+type
        freshly_parsed = [
            r for r in project_recs
            if r.get("island", "").title() == island
            and r.get("rfi_type", "SITE").upper() == rfi_type
        ]

        # Rerouted records from audit pass destined for this table
        rerouted_recs = global_reroute.get(tbl_name, [])

        # Merge: rerouted first (they have priority), then freshly-parsed
        # Deduplicate by RFI No. against each other and against existing
        seen_in_merge: set = set()
        combined_recs = []
        for rec in rerouted_recs + freshly_parsed:
            rno = rec.get("rfi_no", "").strip()
            if rno and rno not in existing_rfis and rno not in seen_in_merge:
                seen_in_merge.add(rno)
                combined_recs.append(rec)

        if not combined_recs:
            print(f"\n  '{tbl_name}': no new records  "
                  f"(existing: {len(existing_rfis)})")
            # Still apply/refresh dropdowns + conditional formatting
            if header_map:
                data_start_1b = tinfo["start_row"] + 2
                try:
                    set_status_validation(service, ssid, tinfo,
                                         header_map, data_start_1b)
                    set_activity_validation(service, ssid, tinfo,
                                            header_map, data_start_1b)
                except Exception:
                    pass
                try:
                    apply_rfi_conditional_formatting(service, ssid, tinfo,
                                                     header_map, data_start_1b)
                    print(f"    ✓ Conditional formatting refreshed")
                except Exception as e:
                    print(f"    ⚠  Conditional formatting: {e}")
                # Push status PDF updates to existing rows even when no new rows added
                if status_lookup:
                    try:
                        patched = push_status_updates_to_sheet(
                            service, ssid, tinfo, header_map,
                            data_start_1b, status_lookup)
                        if patched:
                            print(f"    ✓ Status updated for {patched} existing row(s)")
                    except Exception as e:
                        print(f"    ⚠  Status push: {e}")
            continue

        # ── Sort by Inspection Date (primary), then RFI No. (tiebreaker) ─────
        combined_recs.sort(key=lambda r: (
            _date_sort_key(r.get("inspection_date", "")),
            r.get("rfi_no", "")
        ))

        # ── Determine insertion strategy ───────────────────────────────────────
        # Read existing rows' dates to detect gap-fill (a new record's date
        # is EARLIER than the last existing row → it belongs in the middle).
        existing_date_rows = _read_existing_date_rows(service, ssid, tinfo, header_map)
        footer_0b = tinfo["end_row"] - 1

        last_existing_dk = existing_date_rows[-1][0] if existing_date_rows else (0, 0, 0)
        needs_gap_fill   = existing_date_rows and any(
            _date_sort_key(r.get("inspection_date", "")) < last_existing_dk
            for r in combined_recs
        )

        try:
            if not needs_gap_fill:
                # ── Common case: all new records go after all existing ──────────
                # Batch-insert sorted records above the footer.
                first_row_1b = footer_0b + 1
                rows = [
                    build_row_for_table(max_sno + i + 1, rec,
                                        header_map, tinfo, first_row_1b + i)
                    for i, rec in enumerate(combined_recs)
                ]
                insert_rows_above_footer(service, ssid, tinfo, rows)
                print(f"\n  ✓ '{tbl_name}': inserted {len(rows)} record(s) above footer")

            else:
                # ── Gap-fill: insert each record at its correct date position ──
                # Process from bottom to top so earlier insertions don't
                # invalidate the row indices of later ones.
                print(f"\n  '{tbl_name}': gap-fill mode — inserting {len(combined_recs)} "
                      f"record(s) at correct date positions")
                sheet_id = tinfo["sheet_id"]

                # Build list of (insert_before_0b, rec), then sort bottom→top
                insertions = []
                for rec in combined_recs:
                    dk  = _date_sort_key(rec.get("inspection_date", ""))
                    pos = _find_insert_pos(existing_date_rows, dk,
                                           rec.get("rfi_no", ""), footer_0b)
                    insertions.append((pos, rec))

                insertions.sort(key=lambda x: x[0], reverse=True)  # bottom → top

                for before_0b, rec in insertions:
                    _insert_single_blank_row(service, ssid, sheet_id, before_0b)
                    row_data = build_row_for_table(
                        0, rec, header_map, tinfo, before_0b + 1
                    )
                    _write_single_row(service, ssid, tinfo, before_0b, row_data)
                    # Shift all tracked row indices that are at or below insertion point
                    existing_date_rows = [
                        (dk, r0b + 1) if r0b >= before_0b else (dk, r0b)
                        for dk, r0b in existing_date_rows
                    ]
                    # Register the newly inserted row so subsequent insertions
                    # (processed bottom-up) see the updated state
                    new_dk = _date_sort_key(rec.get("inspection_date", ""))
                    existing_date_rows.append((new_dk, before_0b))
                    existing_date_rows.sort(key=lambda x: (x[0], x[1]))
                    # Footer also shifted down
                    footer_0b += 1

                print(f"\n  ✓ '{tbl_name}': gap-filled {len(combined_recs)} record(s)")

            total_inserted += len(combined_recs)

        except Exception as e:
            print(f"\n  ✗ '{tbl_name}': insert failed — {e}")
            continue

        # Apply / refresh Status + Activity dropdown validation + conditional formatting
        try:
            named_tables_fresh = get_all_named_tables(service, ssid)
            tinfo_fresh        = named_tables_fresh.get(tbl_name, tinfo)
            data_start_1b      = tinfo_fresh["start_row"] + 2
            set_status_validation(service, ssid, tinfo_fresh,
                                  header_map, data_start_1b)
            set_activity_validation(service, ssid, tinfo_fresh,
                                    header_map, data_start_1b)
            print(f"    ✓ Status + Activity dropdown validation applied")
        except Exception as e:
            print(f"    ⚠  Dropdown validation: {e}")

        # Apply conditional formatting (colour-code by Status on RFI No. cell)
        try:
            _nt_cf   = get_all_named_tables(service, ssid)
            tinfo_cf = _nt_cf.get(tbl_name, tinfo)
            ds_cf    = tinfo_cf["start_row"] + 2
            apply_rfi_conditional_formatting(service, ssid, tinfo_cf,
                                             header_map, ds_cf)
            print(f"    ✓ Conditional formatting applied "
                  f"(OPENED=Yellow / APPROVED=Green / NOT APPROVED=Red / "
                  f"CLOSED=Black / CANCELLED=Strikethrough)")
        except Exception as e:
            print(f"    ⚠  Conditional formatting: {e}")

        # Push status PDF updates to all existing rows (not just newly inserted)
        if status_lookup:
            try:
                _nt_st   = get_all_named_tables(service, ssid)
                tinfo_st = _nt_st.get(tbl_name, tinfo)
                ds_st    = tinfo_st["start_row"] + 2
                patched  = push_status_updates_to_sheet(
                    service, ssid, tinfo_st, header_map, ds_st, status_lookup)
                if patched:
                    print(f"    ✓ Status from summary PDF applied to "
                          f"{patched} existing row(s)")
            except Exception as e:
                print(f"    ⚠  Status push to existing rows: {e}")

        # Re-fetch after insert changed end_row
        named_tables = get_all_named_tables(service, ssid)
        time.sleep(0.5)

    # ── Final renumber after all inserts to give every row a correct S/No. ──
    if total_inserted > 0:
        print(f"\n  ── Final renumber after inserts ───────────────────────────────")
        named_tables = get_all_named_tables(service, ssid)
        for island, rfi_type in routing:
            tbl_name = build_table_name(km_key, island, rfi_type)
            if tbl_name not in named_tables:
                continue
            tinfo = named_tables[tbl_name]
            header_map, _, _, data_count = get_table_existing_data(
                service, ssid, tinfo)
            if not header_map or data_count == 0:
                continue
            try:
                renumber_sno_in_table(service, ssid, tinfo, header_map)
            except Exception as e:
                print(f"    ⚠  Final renumber for '{tbl_name}': {e}")
            time.sleep(0.3)

    # ── Share spreadsheet ─────────────────────────────────────────────────────
    print()
    share_spreadsheet(drive, ssid, emails)

    print(f"\n  ✅  {name}  ({total_inserted} new rows inserted)")
    print(f"      URL: https://docs.google.com/spreadsheets/d/{ssid}")
    return total_inserted


# ══════════════════════════════════════════════════════════════════════════════
# MAIN
# ══════════════════════════════════════════════════════════════════════════════

def main():
    print(f"\n{'='*70}")
    print("IRIS RFI → Google Sheets Uploader  (v5 — insert-above-footer, QAQC-aware)")
    print(f"{'='*70}\n")

    # ── Load & parse all RFI documents ────────────────────────────────────────
    all_records, _download_log, status_lookup = load_all_records()
    if not all_records:
        print("No records found. Check ZIP_FOLDER path.")
        return

    # ── Attach merged station lengths (subgrade / excavation / filling) ───────
    attach_merged_lengths(all_records)

    # ── Authenticate with Google ──────────────────────────────────────────────
    print("\nAuthenticating with Google …")
    creds = get_google_creds()
    from googleapiclient.discovery import build
    service = build("sheets", "v4", credentials=creds)
    drive   = build("drive",  "v3", credentials=creds)

    # ── Load canonical road names from Physical Progress sheets ───────────────
    print("\nLoading road names from Physical Progress sheets …")
    road_lookup = load_road_data(service)
    apply_road_lookup(all_records, road_lookup)

    # ── Tally ─────────────────────────────────────────────────────────────────
    r655     = [r for r in all_records if "65"  in r.get("project_km", "")]
    r163     = [r for r in all_records if "163" in r.get("project_km", "")]
    r_unknwn = [r for r in all_records if not r.get("project_km", "").strip()]
    print(f"\nRecords → 65.5 KM: {len(r655)}   163.22 KM: {len(r163)}")
    if r_unknwn:
        print(f"  ⚠  {len(r_unknwn)} record(s) have unknown project (road code absent/unreadable — NOT logged):")
        for r in r_unknwn[:20]:
            print(f"    • {r.get('rfi_no','?'):40s}  {r.get('doc_file','?')}")
        if len(r_unknwn) > 20:
            print(f"    … and {len(r_unknwn) - 20} more")

    # ── Spreadsheet 1: 65.5 KM ────────────────────────────────────────────────
    inserted_655 = process_spreadsheet(
        service=service, drive=drive,
        name=SPREADSHEET_655,
        project_filter="65.5",
        km_key="65.5",
        emails=EMAILS_655,
        all_records=all_records,
        status_lookup=status_lookup,
    )

    # ── Spreadsheet 2: 163.22 KM ──────────────────────────────────────────────
    inserted_163 = process_spreadsheet(
        service=service, drive=drive,
        name=SPREADSHEET_163,
        project_filter="163.22",
        km_key="163.22",
        emails=EMAILS_163,
        all_records=all_records,
        status_lookup=status_lookup,
    )

    # ── Summary ───────────────────────────────────────────────────────────────
    print(f"\n{'='*70}")
    print("✅  ALL DONE")
    print(f"   Total RFIs parsed    :  {len(all_records)}")
    print(f"     65.5 KM            :  {len(r655)}")
    print(f"     163.22 KM          :  {len(r163)}")
    print(f"\n   New rows inserted    :  {inserted_655 + inserted_163}")
    print(f"     65.5 KM            :  {inserted_655}")
    print(f"     163.22 KM          :  {inserted_163}")
    print(f"\n   Spreadsheet 1 (65.5 KM)   : {SPREADSHEET_655}")
    print(f"   Spreadsheet 2 (163.22 KM) : {SPREADSHEET_163}")
    print(f"\n   NOTE: Pivot tables are managed manually — they auto-update")
    print(f"         as new rows are inserted into the SITE/LAB tables.")
    print(f"{'='*70}\n")


if __name__ == "__main__":
    main()
