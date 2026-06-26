"""
IRIS RFI Detailed Log Builder  (Excel version)
================================================
Reads ALL files from D:\\HIGHWAY ENGINEERING CAREER\\RFI-ZIPs\\
(both inside ZIP archives AND standalone PDFs/Word docs),
derives canonical Activity Names from Work Descriptions (never verbatim),
stratifies Subgrade layers (1st, 2nd, 3rd …), merges overlapping station
ranges to compute true lengths, autofills contract numbers, then saves:

  D:\\HIGHWAY ENGINEERING CAREER\\RFI-ZIPs\\IRIS_RFI_Detailed_Log.xlsx

Sheets:
  ALL RFIs
  65.5 KM — Pemba SITE     65.5 KM — Pemba LAB
  65.5 KM — Unguja SITE    65.5 KM — Unguja LAB
  163.22 KM — Pemba SITE   163.22 KM — Pemba LAB
  163.22 KM — Unguja SITE  163.22 KM — Unguja LAB
  Download Log

Run:
  python C:\\AI-Projects\\build_rfi_detailed_log.py
"""

import os, re, io, sys, zipfile, json, shutil, datetime
from pathlib import Path
from collections import defaultdict

# ── Output / Input paths ──────────────────────────────────────────────────────
ZIP_FOLDER   = r"D:\HIGHWAY ENGINEERING CAREER\RFI-ZIPs"
OUTPUT_XLSX  = r"D:\HIGHWAY ENGINEERING CAREER\RFI-ZIPs\IRIS_RFI_Detailed_Log.xlsx"
TEMP_EXTRACT = r"C:\Temp\rfi_extract"

# ── Contract numbers ──────────────────────────────────────────────────────────
CONTRACT_NUMBERS = {
    "65.5 KM":   "SMZ/P01/W/DP/2024-2025/05",
    "163.22 KM": "SMZ/P01/W/DP/2025-2026/30",
}

# ── Libraries ─────────────────────────────────────────────────────────────────
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

try:
    import pdfplumber
    HAS_PDFPLUMBER = True
except ImportError:
    HAS_PDFPLUMBER = False
    print("WARNING: pdfplumber not installed. Run: pip install pdfplumber")

try:
    from docx import Document as DocxDocument
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False

try:
    from PIL import Image
    import pytesseract
    HAS_OCR = True
except ImportError:
    HAS_OCR = False


# ══════════════════════════════════════════════════════════════════════════════
# ACTIVITY NORMALISATION
# Activity Name is DERIVED from Work Description; never taken verbatim.
# Patterns checked in order; first match wins.
# Subgrade layers stratified by ordinal in description (1st layer, 2nd layer …).
# NOTE: AL1/AL2/… = Alignment Number (road section), NOT a layer number.
# ══════════════════════════════════════════════════════════════════════════════
ACTIVITY_PATTERNS = [
    # ── SITE CLEARANCE ────────────────────────────────────────────────────────
    (r"site\s*clear|clearing\s*(?:of\s+)?site|vegetation\s*clear|bush\s*clear"
     r"|tree\s*(?:cut|remov)|demolish|demolition|strip\s*topsoil", "SITE CLEARANCE"),

    # ── ROADBED PREPARATION ───────────────────────────────────────────────────
    (r"roadbed|road\s*bed|formation\s*(?:level|grading)|shaping\s*(?:of\s+)?road"
     r"|profile\s*grading", "ROADBED PREPARATION"),

    # ── SUBGRADE — stratified by actual layer ordinal in the work description ──
    # NOTE: "AL1", "AL2" … = Alignment Number (road section ID), NOT layer numbers.
    # Layer numbers appear as "1st layer", "2nd layer", "3rd layer" etc. in the text.
    (r"(?:sub[\s\-]?grade|subgrade\s*lay|compaction).*?(?:\b1st\b|first\s+layer|layer\s*[:\-]?\s*1\b|no\.?\s*1\b)"
     r"|1st\s+(?:sub[\s\-]?grade|subgrade)\s*(?:lay|compact|course)?",
     "SUBGRADE LAYING — 1st LAYER"),
    (r"(?:sub[\s\-]?grade|subgrade\s*lay|compaction).*?(?:\b2nd\b|second\s+layer|layer\s*[:\-]?\s*2\b|no\.?\s*2\b)"
     r"|2nd\s+(?:sub[\s\-]?grade|subgrade)\s*(?:lay|compact|course)?",
     "SUBGRADE LAYING — 2nd LAYER"),
    (r"(?:sub[\s\-]?grade|subgrade\s*lay|compaction).*?(?:\b3rd\b|third\s+layer|layer\s*[:\-]?\s*3\b|no\.?\s*3\b)"
     r"|3rd\s+(?:sub[\s\-]?grade|subgrade)\s*(?:lay|compact|course)?",
     "SUBGRADE LAYING — 3rd LAYER"),
    (r"(?:sub[\s\-]?grade|subgrade\s*lay|compaction).*?(?:\b4th\b|fourth\s+layer|layer\s*[:\-]?\s*4\b|no\.?\s*4\b)"
     r"|4th\s+(?:sub[\s\-]?grade|subgrade)\s*(?:lay|compact|course)?",
     "SUBGRADE LAYING — 4th LAYER"),
    (r"(?:sub[\s\-]?grade|subgrade\s*lay|compaction).*?(?:\b5th\b|fifth\s+layer|layer\s*[:\-]?\s*5\b|no\.?\s*5\b)"
     r"|5th\s+(?:sub[\s\-]?grade|subgrade)\s*(?:lay|compact|course)?",
     "SUBGRADE LAYING — 5th LAYER"),
    (r"(?:sub[\s\-]?grade|subgrade\s*lay|compaction).*?(?:\b6th\b|sixth\s+layer|layer\s*[:\-]?\s*6\b|no\.?\s*6\b)"
     r"|6th\s+(?:sub[\s\-]?grade|subgrade)\s*(?:lay|compact|course)?",
     "SUBGRADE LAYING — 6th LAYER"),
    # catch any ordinal layer beyond 6th
    (r"(?:sub[\s\-]?grade|subgrade\s*lay|compaction).*?\b(\d+)(?:st|nd|rd|th)\s+layer"
     r"|\b(\d+)(?:st|nd|rd|th)\s+(?:sub[\s\-]?grade|subgrade)\s*(?:lay|compact)?",
     "SUBGRADE LAYING — nth LAYER"),

    # ── SUBGRADE — generic (no layer ordinal in description) ─────────────────────
    (r"sub[\s\-]?grade|subgrade\s*lay|compaction\s*(?:of\s+)?subgrade"
     r"|excavat|general\s*fill|earthwork|embankment|cut\s*and\s*fill"
     r"|filling\s*(?:and\s*)?compaction", "SUBGRADE LAYING"),

    # ── SUBBASE LAYING ────────────────────────────────────────────────────────
    (r"sub[\s\-]?base\s*(?:lay|cours|material|compaction)?", "SUBBASE LAYING"),

    # ── BASECOURSE ────────────────────────────────────────────────────────────
    (r"base[\s\-]?course|base\s*lay|granular\s*(?:base|sub)", "BASECOURSE"),

    # ── PRIME COAT ────────────────────────────────────────────────────────────
    (r"prime\s*coat|bituminous\s*prime|priming", "PRIME COAT"),

    # ── TACK COAT ─────────────────────────────────────────────────────────────
    (r"tack\s*coat|bond\s*coat|bituminous\s*tack", "TACK COAT"),

    # ── ASPHALT LAYING ────────────────────────────────────────────────────────
    (r"asphalt|bituminous\s*(?:concrete|layer|wearing|binder)"
     r"|AC\s*(?:lay|wearing|binder)|wearing\s*course|binder\s*course"
     r"|surface\s*dressing|hot\s*mix", "ASPHALT LAYING"),

    # ── ROAD MARKING ─────────────────────────────────────────────────────────
    (r"road\s*mark|line\s*mark|thermoplastic\s*mark|paint(?:ing)?\s*(?:of\s+)?road"
     r"|lane\s*mark|centre\s*line", "ROAD MARKING"),

    # ── TRAFFIC SIGNS INSTALLATION ────────────────────────────────────────────
    (r"traffic\s*sign|road\s*sign|signage\s*install|erect(?:ion)?\s*(?:of\s+)?sign",
     "TRAFFIC SIGNS INSTALLATION"),

    # ── CULVERT ──────────────────────────────────────────────────────────────
    (r"culvert|pipe\s*(?:drain|culvert)|headwall|wingwall|inlet|outlet\s*structure",
     "CULVERT"),

    # ── SERVICE DUCTS ────────────────────────────────────────────────────────
    (r"service\s*duct|utility\s*duct|conduit|duct\s*install", "SERVICE DUCTS"),

    # ── GUARD RAIL REPARATION ────────────────────────────────────────────────
    (r"guard[\s\-]?rail|guardrail|crash\s*barrier|safety\s*barrier|railing",
     "GUARD RAIL REPARATION"),

    # ── STRUCTURE REPARATION ─────────────────────────────────────────────────
    (r"structure\s*rep|bridge\s*rep|retaining\s*wall\s*rep|repair.*?structure"
     r"|structural\s*rep|abutment\s*rep", "STRUCTURE REPARATION"),

    # ── DISTRESS ─────────────────────────────────────────────────────────────
    (r"distress|cracking|crack\s*(?:seal|repair)|rutting|rut\s*repair"
     r"|deformation|alligator|raveling|stripping", "DISTRESS"),

    # ── POTHOLE REPARATION ───────────────────────────────────────────────────
    (r"pothole|patching|patch\s*(?:work|repair)|filling\s*(?:of\s+)?pothole",
     "POTHOLE REPARATION"),

    # ── OVERLAY ──────────────────────────────────────────────────────────────
    (r"overlay|re[\s\-]?surfac|mill(?:ing)?\s*and\s*(?:re)?pave|rehabilitation\s*overlay",
     "OVERLAY"),

    # ── SIDE DRAINS CLEARANCE (before SIDE DRAINS) ───────────────────────────
    (r"side\s*drain\s*clear|drain(?:age)?\s*clear|clear(?:ing)?\s*(?:of\s+)?drain"
     r"|de[\s\-]?silt|cleaning\s*(?:of\s+)?drain", "SIDE DRAINS CLEARANCE"),

    # ── SIDE DRAINS ──────────────────────────────────────────────────────────
    (r"side\s*drain|kerb\s*drain|edge\s*drain|drainage\s*channel"
     r"|longitudinal\s*drain|u[\s\-]?drain|open\s*drain\s*construct", "SIDE DRAINS"),

    # ── MIX DESIGNS AND MATERIAL SAMPLING ────────────────────────────────────
    (r"sampling|mix\s*design|material\s*(?:test|samp)|lab(?:oratory)?\s*test"
     r"|quality\s*(?:test|control)|trial\s*mix|marshal(?:l)?|proctor"
     r"|sieve|gradation|cbr|atterberg|moisture\s*content|density\s*test"
     r"|core\s*sample|field\s*test", "MIX DESIGNS AND MATERIAL SAMPLING"),
]


def normalise_activity(work_description: str) -> str:
    """
    Derive canonical activity category from raw Work Description text.
    Returns 'OTHER' when no pattern matches.
    Never called with a pre-existing category name — always takes raw text.
    """
    if not work_description:
        return "OTHER"
    wd = str(work_description).strip()
    for pattern, canonical in ACTIVITY_PATTERNS:
        if re.search(pattern, wd, re.IGNORECASE):
            return canonical
    return "OTHER"


# ══════════════════════════════════════════════════════════════════════════════
# STATION / CHAINAGE HELPERS
# ══════════════════════════════════════════════════════════════════════════════

def parse_chainage(s: str) -> float:
    """Parse '12+500' → 12500 (metres). Returns 0.0 on failure."""
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


def calc_length_km(start: str, end: str):
    """Return |end − start| in kilometres, or None if stations are missing."""
    try:
        s, e = parse_chainage(start), parse_chainage(end)
        if s == 0 and e == 0:
            return None
        return round(abs(e - s) / 1000, 4)
    except Exception:
        return None


def merge_intervals(intervals: list) -> list:
    """Merge overlapping/adjacent (start_m, end_m) intervals."""
    if not intervals:
        return []
    ivs = sorted((min(a, b), max(a, b)) for a, b in intervals if not (a == 0 and b == 0))
    if not ivs:
        return []
    merged = [list(ivs[0])]
    for s, e in ivs[1:]:
        if s <= merged[-1][1]:
            merged[-1][1] = max(merged[-1][1], e)
        else:
            merged.append([s, e])
    return [(a, b) for a, b in merged]


def merged_length_km(intervals: list):
    """Total km after merging overlapping intervals."""
    m = merge_intervals(intervals)
    if not m:
        return None
    return round(sum(e - s for s, e in m) / 1000, 4)


def attach_merged_lengths(records: list) -> None:
    """
    Group records by (project, island, rfi_type, road_code, activity_name),
    merge overlapping station intervals, write merged_length_km in-place.
    """
    groups = defaultdict(list)
    for rec in records:
        key = (rec.get("project_km", ""), rec.get("island", ""),
               rec.get("rfi_type", ""), rec.get("road_code", ""),
               rec.get("activity_name", ""))
        groups[key].append(rec)
    for key, grp in groups.items():
        ivs = []
        for r in grp:
            s = parse_chainage(r.get("start_station", ""))
            e = parse_chainage(r.get("end_station", ""))
            if s != 0 or e != 0:
                ivs.append((s, e))
        ml = merged_length_km(ivs)
        for r in grp:
            r["merged_length_km"] = ml if ml is not None else ""


# ══════════════════════════════════════════════════════════════════════════════
# TEXT EXTRACTION
# ══════════════════════════════════════════════════════════════════════════════

def extract_text_from_pdf_bytes(data: bytes) -> str:
    text = ""
    if HAS_PDFPLUMBER:
        try:
            import io
            with pdfplumber.open(io.BytesIO(data)) as pdf:
                for page in pdf.pages:
                    t = page.extract_text()
                    if t:
                        text += t + "\n"
        except Exception as e:
            print(f"    pdfplumber error: {e}")
    if len(text.strip()) < 100 and HAS_OCR:
        try:
            from pdf2image import convert_from_bytes
            images = convert_from_bytes(data, dpi=200)
            for img in images:
                text += pytesseract.image_to_string(img) + "\n"
        except Exception as e:
            print(f"    OCR fallback: {e}")
    return text


def extract_text_from_docx_bytes(data: bytes) -> str:
    if not HAS_DOCX:
        return ""
    import io
    try:
        doc = DocxDocument(io.BytesIO(data))
        parts = [p.text for p in doc.paragraphs]
        for table in doc.tables:
            for row in table.rows:
                parts.append(" | ".join(c.text for c in row.cells))
        return "\n".join(parts)
    except Exception as e:
        print(f"    docx error: {e}")
        return ""


def _ex(pattern, text, flags=re.IGNORECASE | re.DOTALL):
    m = re.search(pattern, text, flags)
    return m.group(1).strip() if m else ""


def _clean(v, maxlen=500):
    if not v:
        return ""
    return re.sub(r"\s+", " ", str(v)).strip()[:maxlen]


# ══════════════════════════════════════════════════════════════════════════════
# FIELD EXTRACTION
# ══════════════════════════════════════════════════════════════════════════════

def extract_fields(text: str, zip_name: str, pdf_name: str,
                   project: str, island_hint: str, date_hint: str,
                   rfi_type: str, source_path: str) -> dict:
    """Extract all fields from one RFI document's text."""
    rec = {}

    rec["rfi_no"] = (
        _ex(r"RFI\s+NO[:\s]+([A-Z0-9\-]+)", text) or
        _ex(r"\b(IRIS[-\s]?[QA]?[-\s]?RFI[-\s]?[PU]?[-\s]?\d{2}[-\s]?\d{4})\b", text) or
        _ex(r"RFI\s*(?:No\.?|Number|#)\s*[:\-]?\s*([A-Z0-9\-/\.]+)", text) or
        os.path.splitext(pdf_name)[0]
    )

    rec["date_received"] = (
        _ex(r"DATE\s+OF\s+SUBMISSION[:\s]+(\d{1,2}[/\-\.]\d{1,2}[/\-\.]\d{2,4})", text) or
        _ex(r"Date\s+(?:Received|Submitted|Issued)[:\s]+(\d{1,2}[/\-\.]\d{1,2}[/\-\.]\d{2,4})", text) or
        _ex(r"Date\s*[:\-]\s*(\d{1,2}[/\-\.]\d{1,2}[/\-\.]\d{4})", text) or
        date_hint
    )

    rec["inspection_date"] = (
        _ex(r"DATE\s+OF\s+INSPECTION[:\s]+(\d{1,2}[/\-\.]\d{1,2}[/\-\.]\d{2,4})", text) or
        _ex(r"(?:Inspection|Proposed\s+Inspection)\s+Date[:\s]+(\d{1,2}[/\-\.]\d{1,2}[/\-\.]\d{2,4})", text)
    )

    loc = (_ex(r"LOCATION[:\s]+([^\n\r]{2,80})", text) or
           _ex(r"Location\s*[:\-]\s*([^\n\r]{5,80})", text))
    rec["road_code"] = _clean(loc)
    rec["location"]  = _clean(loc)

    rec["road_name"] = _clean(
        _ex(r"ROAD\s+NAME[:\s]+([^\n\r]{3,120})", text) or
        _ex(r"Road\s+Name\s*[:\-]?\s*(.{5,100}?)(?:\n|Island|Contract|Start)", text) or
        _ex(r"Road\s*[:\-]\s*([^\n\r]{3,120})", text)
    )

    island_text = _ex(r"ISLAND[:\s]+(PEMBA|UNGUJA|ZANZIBAR)", text, re.IGNORECASE)
    if not island_text:
        fn = (zip_name + " " + pdf_name).upper()
        if "UNGUJA" in fn or "-U-" in fn:
            island_text = "Unguja"
        elif "PEMBA" in fn or "-P-" in fn:
            island_text = "Pemba"
        else:
            island_text = island_hint
    rec["island"] = island_text.capitalize() if island_text else ""

    rec["contract_number"] = CONTRACT_NUMBERS.get(project, "")

    chainage_raw = (
        _ex(r"CHAINAGE[:\s]+([0-9\+\.\s\-–]+)", text) or
        _ex(r"Station\s+Range[:\s]+([^\n\r]{5,60})", text) or
        _ex(r"(?:Ch\.?|STA\.?)\s+([0-9\+\.\s\-–]{4,30})", text)
    )
    rec["station_range"] = _clean(chainage_raw)

    ch_m = re.search(r"([\d\+]+)\s*[-–]\s*([\d\+]+)", rec["station_range"])
    if ch_m:
        rec["start_station"] = ch_m.group(1)
        rec["end_station"]   = ch_m.group(2)
    else:
        rec["start_station"] = _ex(r"Start\s+(?:Station|Chainage|Ch\.?)[:\s]+([\d\+\.]+)", text)
        rec["end_station"]   = _ex(r"End\s+(?:Station|Chainage|Ch\.?)[:\s]+([\d\+\.]+)", text)

    computed = calc_length_km(rec["start_station"], rec["end_station"])
    if computed:
        rec["length_km"] = computed
    else:
        doc_len = _ex(r"Length.*?[:\s]+([\d\.,]+)\s*(?:KM|M\b)", text, re.IGNORECASE)
        try:
            v = float(doc_len.replace(",", "."))
            rec["length_km"] = v if v < 100 else round(v / 1000, 4)
        except Exception:
            rec["length_km"] = ""

    wd = (
        _ex(r"WORK\s+DESCRIPTION\s*[:\s]+(.+?)(?=\n\s*(?:ENCLOSURES|RELATED|Requested|$))", text) or
        _ex(r"(?:Work|Works|Activity)\s*[:\-]\s*(.{5,300}?)(?:\n|$)", text)
    )
    rec["work_description"] = _clean(wd)

    # Activity Name: CATEGORISED from work_description — never verbatim
    rec["activity_name"] = normalise_activity(rec["work_description"])

    status_m = re.search(
        r"(Approved\s+for\s+Further\s+Work|Not\s+Approved[^,\n]*|Approved\s+with\s+Comment)",
        text, re.IGNORECASE)
    rec["status"] = _clean(status_m.group(1)) if status_m else ""

    re_rem = (
        _ex(r"Comments?\s+by\s+(?:the\s+)?RE[:\s]*(.+?)(?=\nRESIDENT ENGINEER|\nRequest:|\Z)", text) or
        _ex(r"RE['\u2019]?s?\s+(?:REMARKS?|COMMENTS?)[:\s]*(.+?)(?=\n\n|Date|\Z)", text)
    )
    rec["re_remarks"] = _clean(re_rem)

    rec["remark_date"] = (
        _ex(r"RE\s+SIGNATURE.*?DATE[:\s]+(\d{1,2}[/\-\.]\d{1,2}[/\-\.]\d{2,4})", text) or
        _ex(r"(?:Remark|Remarks?)\s+Date[:\s]+(\d{1,2}[/\-\.]\d{1,2}[/\-\.]\d{2,4})", text)
    )

    rec["attachment_path"]   = source_path
    rec["project_km"]        = project
    rec["rfi_type"]          = rfi_type
    rec["zip_file"]          = zip_name
    rec["doc_file"]          = pdf_name
    rec["merged_length_km"]  = ""   # filled after all records loaded
    return rec


# ══════════════════════════════════════════════════════════════════════════════
# LOAD ALL RECORDS
# ══════════════════════════════════════════════════════════════════════════════

DOC_EXTENSIONS = {".pdf", ".docx", ".doc"}


def _hints_from_name(name: str) -> tuple:
    n = name.upper()
    project  = "163.22 KM" if "163" in n else ("65.5 KM" if "65" in n else "")
    island   = "Unguja"    if ("UNGUJA" in n or "-U-" in n) else \
               ("Pemba"    if ("PEMBA"  in n or "-P-" in n) else "")
    rfi_type = "LAB" if "LAB" in n else "SITE"
    dm = re.search(r"(\d{2})\.(\d{2})\.(\d{4})", name)
    date_h = f"{dm.group(1)}/{dm.group(2)}/{dm.group(3)}" if dm else ""
    return project, island, rfi_type, date_h


def load_all_records():
    all_files  = sorted(os.listdir(ZIP_FOLDER))
    zip_files  = [f for f in all_files if f.lower().endswith(".zip")]
    standalone = [f for f in all_files
                  if Path(f).suffix.lower() in DOC_EXTENSIONS and not f.startswith("_")]

    print(f"Found {len(zip_files)} ZIPs + {len(standalone)} standalone docs\n")

    rfi_records  = []
    download_log = []
    seen_rfi     = set()

    # ── Process ZIPs ──────────────────────────────────────────────────────────
    for zi, zf_name in enumerate(zip_files, 1):
        zf_path = os.path.join(ZIP_FOLDER, zf_name)
        project, island_hint, rfi_type, date_h = _hints_from_name(zf_name)

        fstat = os.stat(zf_path)
        download_log.append({
            "filename": zf_name, "type": "ZIP",
            "size_kb": round(fstat.st_size/1024, 1),
            "modified": datetime.datetime.fromtimestamp(fstat.st_mtime).strftime("%d/%m/%Y %H:%M"),
            "project": project, "island": island_hint, "rfi_type": rfi_type,
        })

        try:
            with zipfile.ZipFile(zf_path, "r") as zf:
                docs = [x for x in zf.namelist()
                        if Path(x).suffix.lower() in DOC_EXTENSIONS]
                print(f"  [{zi}/{len(zip_files)}] {zf_name[:60]} ({len(docs)} docs)")

                for doc_name in docs:
                    try:
                        data  = zf.read(doc_name)
                        ext   = Path(doc_name).suffix.lower()
                        text  = (extract_text_from_pdf_bytes(data) if ext == ".pdf"
                                 else extract_text_from_docx_bytes(data))
                        if not text.strip():
                            continue
                        inner = os.path.basename(doc_name)
                        src   = rf"D:\HIGHWAY ENGINEERING CAREER\RFI-ZIPs\{zf_name}\{inner}"
                        rec   = extract_fields(text, zf_name, inner,
                                               project, island_hint, date_h,
                                               rfi_type, src)
                        if rec["rfi_no"] in seen_rfi:
                            continue
                        seen_rfi.add(rec["rfi_no"])
                        rfi_records.append(rec)
                    except Exception as e:
                        print(f"      ERROR {doc_name}: {e}")
        except zipfile.BadZipFile as e:
            print(f"  BAD ZIP {zf_name}: {e}")

    # ── Process standalone files ──────────────────────────────────────────────
    if standalone:
        print(f"\nProcessing {len(standalone)} standalone files …")
    for fname in standalone:
        fpath   = os.path.join(ZIP_FOLDER, fname)
        project, island_hint, rfi_type, date_h = _hints_from_name(fname)
        ext     = Path(fname).suffix.lower()
        fstat   = os.stat(fpath)
        download_log.append({
            "filename": fname, "type": ext.upper().lstrip("."),
            "size_kb": round(fstat.st_size/1024, 1),
            "modified": datetime.datetime.fromtimestamp(fstat.st_mtime).strftime("%d/%m/%Y %H:%M"),
            "project": project, "island": island_hint, "rfi_type": rfi_type,
        })
        try:
            with open(fpath, "rb") as f:
                data = f.read()
            text = (extract_text_from_pdf_bytes(data) if ext == ".pdf"
                    else extract_text_from_docx_bytes(data))
            if not text.strip():
                continue
            src = rf"D:\HIGHWAY ENGINEERING CAREER\RFI-ZIPs\{fname}"
            rec = extract_fields(text, fname, fname, project, island_hint, date_h, rfi_type, src)
            if rec["rfi_no"] in seen_rfi:
                continue
            seen_rfi.add(rec["rfi_no"])
            rfi_records.append(rec)
        except Exception as e:
            print(f"  ERROR {fname}: {e}")

    print(f"\nTotal unique RFI records: {len(rfi_records)}")
    return rfi_records, download_log


# ══════════════════════════════════════════════════════════════════════════════
# EXCEL BUILDER
# ══════════════════════════════════════════════════════════════════════════════

DATA_COLS = [
    ("S/No.",               8),
    ("RFI No.",             22),
    ("Date Received",       14),
    ("Inspection Date",     14),
    ("Road Code",           16),
    ("Road Name",           38),
    ("Island",              10),
    ("Contract Number",     22),
    ("Start Station",       14),
    ("End Station",         14),
    ("Length Executed (km)",16),
    ("Merged Length (km)",  16),
    ("Activity Name",       30),
    ("Work Description",    45),
    ("Station Range",       20),
    ("Status",              22),
    ("RE's Remarks",        45),
    ("Remark Date",         14),
    ("Attachment Path",     65),
    ("Project",             14),
    ("RFI Type",            10),
]

DL_COLS = [
    ("S/No.", 8), ("Filename", 55), ("Type", 8), ("Size (KB)", 10),
    ("Date Modified", 16), ("Project", 14), ("Island", 10), ("RFI Type", 10),
]

HEADER_FILLS = {
    "ALL":    PatternFill("solid", fgColor="2F2F2F"),
    "65.5":   PatternFill("solid", fgColor="1F4E79"),
    "163.22": PatternFill("solid", fgColor="1E6B3C"),
    "Pemba":  PatternFill("solid", fgColor="7B3F00"),
    "Unguja": PatternFill("solid", fgColor="4B0082"),
    "LAB":    PatternFill("solid", fgColor="8B0000"),
    "LOG":    PatternFill("solid", fgColor="444444"),
}
ROW_FILLS = {
    "ALL":    [PatternFill("solid", fgColor="F5F5F5"), PatternFill("solid", fgColor="EBEBEB")],
    "65.5":   [PatternFill("solid", fgColor="DDEEFF"), PatternFill("solid", fgColor="C5E0FF")],
    "163.22": [PatternFill("solid", fgColor="DDFFEE"), PatternFill("solid", fgColor="C5FFE0")],
    "Pemba":  [PatternFill("solid", fgColor="FFE8D0"), PatternFill("solid", fgColor="FFD5B0")],
    "Unguja": [PatternFill("solid", fgColor="EEE0FF"), PatternFill("solid", fgColor="DDC5FF")],
    "LAB":    [PatternFill("solid", fgColor="FFE8E8"), PatternFill("solid", fgColor="FFD0D0")],
    "LOG":    [PatternFill("solid", fgColor="F0F0F0"), PatternFill("solid", fgColor="E0E0E0")],
}

thin   = Side(style="thin", color="CCCCCC")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)


def record_to_row(sno: int, rec: dict) -> list:
    return [
        sno, rec.get("rfi_no",""), rec.get("date_received",""), rec.get("inspection_date",""),
        rec.get("road_code",""), rec.get("road_name",""), rec.get("island",""),
        rec.get("contract_number",""), rec.get("start_station",""), rec.get("end_station",""),
        rec.get("length_km",""), rec.get("merged_length_km",""),
        rec.get("activity_name",""), rec.get("work_description",""), rec.get("station_range",""),
        rec.get("status",""), rec.get("re_remarks",""), rec.get("remark_date",""),
        rec.get("attachment_path",""), rec.get("project_km",""), rec.get("rfi_type",""),
    ]


def dl_row(sno: int, e: dict) -> list:
    return [sno, e.get("filename",""), e.get("type",""), e.get("size_kb",""),
            e.get("modified",""), e.get("project",""), e.get("island",""), e.get("rfi_type","")]


def write_sheet(ws, title: str, records_rows: list, col_defs: list, col_key: str):
    from datetime import datetime as dt
    hfill = HEADER_FILLS.get(col_key, HEADER_FILLS["ALL"])
    rfi   = ROW_FILLS.get(col_key, ROW_FILLS["ALL"])
    wf    = Font(color="FFFFFF", bold=True, size=9)
    nf    = Font(size=8)
    wrap_ci = {5, 6, 13, 16, 18}   # 1-based indices of wide-text columns

    # Merge title row
    ws.merge_cells(f"A1:{get_column_letter(len(col_defs))}1")
    c = ws["A1"]
    c.value = f"IRIS RFI LOG — {title}"
    c.font  = Font(size=12, bold=True, color="FFFFFF")
    c.fill  = hfill
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 24

    # Timestamp
    ws.merge_cells(f"A2:{get_column_letter(len(col_defs))}2")
    c = ws["A2"]
    c.value = f"Generated: {dt.now().strftime('%d %B %Y  %H:%M')}   |   Records: {len(records_rows)}"
    c.font  = Font(size=9, italic=True, color="555555")
    c.alignment = Alignment(horizontal="center")
    ws.row_dimensions[2].height = 14

    # Header row
    for ci, (col_name, col_w) in enumerate(col_defs, 1):
        c = ws.cell(row=3, column=ci, value=col_name)
        c.font = wf; c.fill = hfill; c.border = BORDER
        c.alignment = Alignment(horizontal="center", vertical="middle", wrap_text=True)
        ws.column_dimensions[get_column_letter(ci)].width = col_w
    ws.row_dimensions[3].height = 30

    # Data rows
    for rn, row_vals in enumerate(records_rows, 1):
        er   = rn + 3
        fill = rfi[rn % 2]
        for ci, val in enumerate(row_vals, 1):
            c = ws.cell(row=er, column=ci, value=val)
            c.fill = fill; c.font = nf; c.border = BORDER
            c.alignment = Alignment(vertical="top", wrap_text=(ci in wrap_ci))
        ws.row_dimensions[er].height = 18

    ws.auto_filter.ref = f"A3:{get_column_letter(len(col_defs))}{len(records_rows)+3}"
    ws.freeze_panes   = "A4"


def build_excel(all_records: list, download_log: list, output_path: str):
    wb = openpyxl.Workbook()

    def seg(recs, project=None, island=None, rfi_type=None):
        out = recs
        if project:  out = [r for r in out if project in r.get("project_km","")]
        if island:   out = [r for r in out if r.get("island","").lower() == island.lower()]
        if rfi_type: out = [r for r in out if r.get("rfi_type","") == rfi_type]
        return out

    SHEET_PLAN = [
        ("ALL RFIs",                "ALL",    {}),
        ("65.5 KM — Pemba SITE",    "65.5",   {"project":"65.5",   "island":"Pemba",  "rfi_type":"SITE"}),
        ("65.5 KM — Pemba LAB",     "LAB",    {"project":"65.5",   "island":"Pemba",  "rfi_type":"LAB"}),
        ("65.5 KM — Unguja SITE",   "65.5",   {"project":"65.5",   "island":"Unguja", "rfi_type":"SITE"}),
        ("65.5 KM — Unguja LAB",    "LAB",    {"project":"65.5",   "island":"Unguja", "rfi_type":"LAB"}),
        ("163.22 KM — Pemba SITE",  "163.22", {"project":"163.22", "island":"Pemba",  "rfi_type":"SITE"}),
        ("163.22 KM — Pemba LAB",   "LAB",    {"project":"163.22", "island":"Pemba",  "rfi_type":"LAB"}),
        ("163.22 KM — Unguja SITE", "163.22", {"project":"163.22", "island":"Unguja", "rfi_type":"SITE"}),
        ("163.22 KM — Unguja LAB",  "LAB",    {"project":"163.22", "island":"Unguja", "rfi_type":"LAB"}),
    ]

    first = True
    for title, col_key, filters in SHEET_PLAN:
        ws = wb.active if first else wb.create_sheet()
        first = False
        ws.title = title
        recs = seg(all_records, **filters)
        rows = [record_to_row(i+1, r) for i, r in enumerate(recs)]
        write_sheet(ws, title, rows, DATA_COLS, col_key)
        print(f"  Sheet '{title}': {len(recs)} rows")

    # Download log sheet
    ws_dl = wb.create_sheet("Download Log")
    dl_rows = [dl_row(i+1, e) for i, e in enumerate(download_log)]
    write_sheet(ws_dl, "Download Log", dl_rows, DL_COLS, "LOG")
    print(f"  Sheet 'Download Log': {len(download_log)} entries")

    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    wb.save(output_path)
    print(f"\nExcel saved → {output_path}")
    print(f"Total records: {len(all_records)}")


# ══════════════════════════════════════════════════════════════════════════════
# MAIN
# ══════════════════════════════════════════════════════════════════════════════
def main():
    if not os.path.exists(ZIP_FOLDER):
        print(f"ZIP folder not found: {ZIP_FOLDER}")
        return

    print(f"\n{'='*60}")
    print("IRIS RFI Detailed Log Builder")
    print(f"{'='*60}\n")

    all_records, download_log = load_all_records()
    if not all_records:
        print("No RFI records extracted.")
        return

    # Compute merged lengths for each road+activity group
    attach_merged_lengths(all_records)

    # Build Excel workbook
    print(f"\nBuilding Excel …")
    build_excel(all_records, download_log, OUTPUT_XLSX)

    print(f"\n{'='*60}")
    print("DONE! Open IRIS_RFI_Detailed_Log.xlsx in the RFI-ZIPs folder.")
    print(f"{'='*60}\n")


if __name__ == "__main__":
    main()
