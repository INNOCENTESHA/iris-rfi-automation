"""
IRIS Road Code Matcher  (v1 — April 2026)
==========================================
Fuzzy-matches KMZ road names against the canonical Road Code & Name
tables from both Google Sheets projects.

Matching strategy
-----------------
1. Numeric-code match: if the 3-digit number in a KMZ road name is the
   tail of a Physical-Progress code (e.g. "243" inside "MK-243"), that
   pair gets a bonus of +0.30 before final scoring.

2. Token-Jaccard similarity: set of cleaned word tokens, Jaccard index.

3. SequenceMatcher ratio: character-level partial similarity.

4. Combined score = 0.55 × Jaccard + 0.25 × SeqMatcher + 0.20 × code_bonus

Acceptance thresholds
---------------------
  Single candidate  : score ≥ 0.40  (50% − 10%)
  Multiple candidates: best score ≥ 0.70  (50% + 20%), and the best
                       must be at least 0.10 above the runner-up.
  Otherwise         : flagged as AMBIGUOUS or NO_MATCH for manual review.

Output
------
  C:\\AI-Projects\\road_code_mapping.json
  C:\\AI-Projects\\road_code_mapping_report.txt

Run
---
  python C:\\AI-Projects\\match_road_codes.py
"""

from __future__ import annotations
import json, re, sys, os
from difflib   import SequenceMatcher
from pathlib   import Path
from datetime  import datetime
from collections import defaultdict

from project_sheets import PP_ID, RFI_ID, physical_progress_tab

# ── Google auth (same token as other scripts) ─────────────────────────────────
try:
    from google.oauth2.credentials  import Credentials
    from google.auth.transport.requests import Request
    from googleapiclient.discovery  import build
    HAS_GOOGLE = True
except ImportError:
    HAS_GOOGLE = False

BASE_DIR   = Path(r"C:\AI-Projects")
TOKEN_FILE = BASE_DIR / "google_token.json"

# ── Physical Progress sheet IDs ───────────────────────────────────────────────
# ── IRIS RFI sheet IDs ────────────────────────────────────────────────────────

OUTPUT_JSON   = BASE_DIR / "road_code_mapping.json"
OUTPUT_REPORT = BASE_DIR / "road_code_mapping_report.txt"

# ── KMZ road data (embedded — generated from the uploaded KMZ file) ───────────
# Source: 163.5 km (06-11-2025).kmz parsed into groups.
# Each entry: {code, name, island, segments, total_length_km}
KMZ_ROADS = [
    # ── PEMBA ────────────────────────────────────────────────────────────────
    {"num":"243","name":"Chambani (kwa Mbunge) - JKU - Mtangani",                "island":"Pemba"},
    {"num":"240","name":"Mjini Kiuyu msituni - Kiuyu road (Micheweni Access Road)","island":"Pemba"},
    {"num":"240","name":"Kiuyu Mbuyuni Bandari (Micheweni Access Road)",          "island":"Pemba"},
    {"num":"240","name":"Maskani Majenzi Micheweni Access",                        "island":"Pemba"},
    {"num":"",   "name":"Michwewni Mchupwe (Micheweni Access Road)",              "island":"Pemba"},
    {"num":"239","name":"Mtamba - Kitondoo + Tomondo",                             "island":"Pemba"},
    {"num":"239","name":"Mtamba - Kidondoo AL1",                                   "island":"Pemba"},
    {"num":"",   "name":"Dutani",                                                  "island":"Pemba"},
    {"num":"",   "name":"Mtamba - Kidondoo AL2",                                   "island":"Pemba"},
    {"num":"238","name":"Minazini - Shikazako",                                    "island":"Pemba"},
    {"num":"",   "name":"Kichaka - Chwaaka Kubwa + Chwaka Ndogo Spar",            "island":"Pemba"},
    {"num":"235","name":"Kichungwani - Kwa Sanani",                               "island":"Pemba"},
    {"num":"",   "name":"Mauani - School - Kirimdado - Spar",                     "island":"Pemba"},
    {"num":"",   "name":"Kiwani - Sokani - Mchochani",                            "island":"Pemba"},
    {"num":"242","name":"Mizingani - Wambaa + Fundu Lagoon",                      "island":"Pemba"},
    {"num":"244","name":"Konde Sokoni - Kiuyu Kwa Manda",                         "island":"Pemba"},
    {"num":"245","name":"Kuyuni - Kalani",                                         "island":"Pemba"},
    {"num":"252","name":"Pujini (Kilimo) - Mkandandume - Kibaridi",               "island":"Pemba"},
    {"num":"246","name":"Mianzini - Mabatini",                                     "island":"Pemba"},
    {"num":"247","name":"Kengeja - Minazini - Kangani",                            "island":"Pemba"},
    {"num":"248","name":"Chanjaani - Mitamani - Pujuni",                           "island":"Pemba"},
    {"num":"249","name":"Mtambile (Police Post) - Kangaani",                       "island":"Pemba"},
    {"num":"250","name":"Kenya Kwa Mtora",                                         "island":"Pemba"},
    {"num":"237","name":"Mahuduthi - Sipwese + Mto Mkuu Spar",                    "island":"Pemba"},
    {"num":"251","name":"Wete Access Road",                                        "island":"Pemba"},
    {"num":"210","name":"Mwembekungu - Kaskazi",                                  "island":"Pemba"},
    # ── UNGUJA ───────────────────────────────────────────────────────────────
    {"num":"",   "name":"Mikungunini Access Road Matemwe",                         "island":"Unguja"},
    {"num":"217","name":"Donge Mbiji - Karange + Spur",                            "island":"Unguja"},
    {"num":"220","name":"Kisauni Mazurui - Tangi la Maji feeder road and spur",   "island":"Unguja"},
    {"num":"201","name":"Istiqama - Zanzibar University",                          "island":"Unguja"},
    {"num":"223","name":"Ziwani Police access roads",                              "island":"Unguja"},
    {"num":"232","name":"Maungani Vijijini - Shakani",                             "island":"Unguja"},
    {"num":"200","name":"Muyuni Junction - Sunshine Marine Lodge",                 "island":"Unguja"},
    {"num":"203","name":"Muyuni B (Kidodoni) - Ngombeni - Skuli Ndorobo Spur",   "island":"Unguja"},
    {"num":"202","name":"Muyuni A (Kibaoni) - Muyuni C (Dindi)",                  "island":"Unguja"},
    {"num":"204","name":"Dindi - Mchangamle - Wasusanya - Kilimatangi Spar",      "island":"Unguja"},
    {"num":"206","name":"Kisokoni - Kaskazi Kikwecheni",                           "island":"Unguja"},
    {"num":"207","name":"Madete Skuli - Skuli ya Kusini",                          "island":"Unguja"},
    {"num":"208","name":"Mnarani - Kigaeni",                                       "island":"Unguja"},
    {"num":"211","name":"Muembesanda - Mayua",                                     "island":"Unguja"},
    {"num":"212","name":"Nganani - Bungi - Mtegani - Kumbini - Chuini",           "island":"Unguja"},
    {"num":"214","name":"Tasani - Mnarani - Pwani",                                "island":"Unguja"},
    {"num":"205","name":"Wadini - Cheju",                                           "island":"Unguja"},
    {"num":"205","name":"Kidodo Firishi - Kizimachakacha",                         "island":"Unguja"},
    {"num":"209","name":"Mtendemndo - Pwani",                                       "island":"Unguja"},
    {"num":"230","name":"Kisauni Mtundani feeder road",                            "island":"Unguja"},
    {"num":"229","name":"Mitondooni Bakery access road",                           "island":"Unguja"},
    {"num":"224","name":"Kilimani Prison Headquarters access roads",               "island":"Unguja"},
    {"num":"213","name":"Access Road to JKU Dunga camp including car parking",    "island":"Unguja"},
    {"num":"225","name":"Bavuai camp petrol station - Migombani JWTZ camp feeder road","island":"Unguja"},
    {"num":"226","name":"Access Road to Mtoni JWTZ camp",                          "island":"Unguja"},
    {"num":"216","name":"Fuoni Mskiti wa Waarabu - Fuoni Michenzani + Spur",      "island":"Unguja"},
    {"num":"218","name":"Fuoni Mtundani + Jangamizi feeder road",                  "island":"Unguja"},
    {"num":"231","name":"Fuoni Melisita Chunga - Mwera Melisita",                  "island":"Unguja"},
    {"num":"222","name":"Flawa - Suleiman Kondo mskitini access road",             "island":"Unguja"},
    {"num":"221","name":"Pili space Dole - Shamba la Bishafika feeder road",      "island":"Unguja"},
    {"num":"219","name":"Jumbi Makonteni (Meli Tisa) and spur",                    "island":"Unguja"},
]

# Stop-words to exclude from token matching
STOP_WORDS = {
    "road","feeder","access","spur","al1","al2","al3","al4","al5",
    "and","the","to","of","in","at","km","mt","plus","via","including",
    "camp","petrol","station","car","parking","police","post",
}


# ══════════════════════════════════════════════════════════════════════════════
# TEXT UTILITIES
# ══════════════════════════════════════════════════════════════════════════════

def clean_name(s: str) -> str:
    """Remove leading road-number prefix and normalise separators."""
    s = re.sub(r"^\d{2,3}[-\s]+", "", str(s))   # strip leading "243-" or "240 "
    s = re.sub(r"[_\+\(\)]", " ", s)            # replace symbols with space
    s = re.sub(r"\b(al\d|al \d)\b", "", s, flags=re.I)  # remove AL1/AL2
    s = re.sub(r"\d+(\.\d+)?\s*(km|m|mt)\b", "", s, flags=re.I)  # remove lengths
    s = re.sub(r"\s{2,}", " ", s)
    return s.strip().lower()


def tokenise(s: str) -> set[str]:
    tokens = set(re.findall(r"[a-z]{2,}", clean_name(s).lower()))
    return tokens - STOP_WORDS


def jaccard(t1: set, t2: set) -> float:
    if not t1 or not t2:
        return 0.0
    return len(t1 & t2) / len(t1 | t2)


def seq_ratio(s1: str, s2: str) -> float:
    return SequenceMatcher(None, clean_name(s1), clean_name(s2)).ratio()


def numeric_bonus(kmz_num: str, sheet_code: str) -> float:
    """Return 0.30 if the numeric part of sheet_code equals kmz_num."""
    if not kmz_num:
        return 0.0
    m = re.search(r"\d+", sheet_code)
    if m and m.group() == kmz_num:
        return 0.30
    return 0.0


def score(kmz: dict, sheet_code: str, sheet_name: str) -> float:
    t1  = tokenise(kmz["name"])
    t2  = tokenise(sheet_name)
    j   = jaccard(t1, t2)
    s   = seq_ratio(kmz["name"], sheet_name)
    nb  = numeric_bonus(kmz.get("num",""), sheet_code)
    return round(0.55 * j + 0.25 * s + 0.20 * nb, 4)


# ══════════════════════════════════════════════════════════════════════════════
# MATCHING LOGIC
# ══════════════════════════════════════════════════════════════════════════════

THRESHOLD_BASE  = 0.50
THRESHOLD_HIGH  = 0.50 + 0.20   # = 0.70  — accept best of multiple candidates
THRESHOLD_LOW   = 0.50 - 0.10   # = 0.40  — accept single candidate
MIN_GAP         = 0.10           # best must be this much ahead of runner-up


def match_road(kmz: dict, candidates: list[dict]) -> dict:
    """
    candidates: [{code, name, island, project}, ...]
    Returns match result dict.
    """
    scored = sorted(
        [{"code": c["code"], "name": c["name"], "island": c["island"],
          "project": c["project"], "score": score(kmz, c["code"], c["name"])}
         for c in candidates],
        key=lambda x: -x["score"]
    )

    if not scored:
        return {"status": "NO_CANDIDATES", "matches": []}

    best   = scored[0]
    second = scored[1] if len(scored) > 1 else None

    # ── Single candidate ──────────────────────────────────────────────────
    if second is None:
        if best["score"] >= THRESHOLD_LOW:
            return {"status": "MATCHED", "confidence": best["score"],
                    "code": best["code"], "name": best["name"],
                    "island": best["island"], "project": best["project"],
                    "matches": scored[:1]}
        return {"status": "NO_MATCH", "matches": scored[:3]}

    # ── Multiple candidates ───────────────────────────────────────────────
    gap = best["score"] - second["score"]

    if best["score"] >= THRESHOLD_HIGH and gap >= MIN_GAP:
        return {"status": "MATCHED", "confidence": best["score"],
                "code": best["code"], "name": best["name"],
                "island": best["island"], "project": best["project"],
                "matches": scored[:3]}

    if best["score"] >= THRESHOLD_BASE and gap >= MIN_GAP:
        return {"status": "MATCHED", "confidence": best["score"],
                "code": best["code"], "name": best["name"],
                "island": best["island"], "project": best["project"],
                "matches": scored[:3]}

    if best["score"] >= THRESHOLD_LOW and gap >= MIN_GAP:
        return {"status": "LOW_CONFIDENCE", "confidence": best["score"],
                "code": best["code"], "name": best["name"],
                "island": best["island"], "project": best["project"],
                "matches": scored[:3]}

    return {"status": "AMBIGUOUS", "matches": scored[:5]}


# ══════════════════════════════════════════════════════════════════════════════
# GOOGLE SHEETS READER
# ══════════════════════════════════════════════════════════════════════════════

def auth():
    if not HAS_GOOGLE:
        return None
    with open(TOKEN_FILE) as f:
        td = json.load(f)
    creds = Credentials(
        token=td.get("token"), refresh_token=td.get("refresh_token"),
        token_uri=td.get("token_uri"), client_id=td.get("client_id"),
        client_secret=td.get("client_secret"), scopes=td.get("scopes"),
    )
    if creds.expired and creds.refresh_token:
        creds.refresh(Request())
    return build("sheets", "v4", credentials=creds)


def read_road_codes_from_sheet(service, ssid: str, project: str) -> list[dict]:
    """Read Road Code & Road Names tab, return [{code, name, island, project}]."""
    roads = []
    try:
        meta   = service.spreadsheets().get(spreadsheetId=ssid).execute()
        sheets = [s["properties"]["title"] for s in meta["sheets"]]

        # Try canonical tab name first, then scan
        tab = next(
            (t for t in sheets if "road" in t.lower() and (
                "code" in t.lower() or "name" in t.lower())),
            None
        )
        if tab is None:
            tab = next(
                (t for t in sheets if "road" in t.lower()),
                sheets[0]
            )

        result = service.spreadsheets().values().get(
            spreadsheetId=ssid,
            range=f"'{tab}'!A:F",
        ).execute()
        rows = result.get("values", [])
        if not rows:
            return []

        header = [str(h).strip().lower() for h in rows[0]]
        code_col = next((i for i, h in enumerate(header)
                         if "code" in h), None)
        name_col = next((i for i, h in enumerate(header)
                         if "name" in h and "code" not in h), None)

        if code_col is None:
            # Try without header detection — assume col A=code, col B=name
            code_col, name_col = 0, 1

        island_col = next((i for i, h in enumerate(header)
                           if "island" in h), None)

        for row in rows[1:]:
            def cell(c):
                return str(row[c]).strip() if c is not None and c < len(row) else ""
            code   = cell(code_col).upper()
            name   = cell(name_col)
            island = cell(island_col) if island_col else ""
            if code and name and len(code) >= 2:
                roads.append({"code": code, "name": name,
                               "island": island, "project": project})

        print(f"  ✓ {len(roads)} roads from '{tab}' ({project} KM)")
    except Exception as e:
        print(f"  ⚠  Could not read Road Code tab ({project} KM): {e}")
    return roads


def read_from_physical_progress(service, ssid: str, project: str) -> list[dict]:
    """Fallback: scan Physical Progress sheet to harvest road codes + names."""
    roads = []
    seen  = set()
    try:
        meta = service.spreadsheets().get(spreadsheetId=ssid).execute()
        tabs = [s["properties"]["title"] for s in meta["sheets"]]
        pp_tab = physical_progress_tab(tabs)
        result = service.spreadsheets().values().get(
            spreadsheetId=ssid,
            range=f"'{pp_tab}'!A:F",
        ).execute()
        rows = result.get("values", [])

        for row in rows:
            for cell in row[:6]:
                c = str(cell).strip().upper()
                if re.match(r"^[A-Z]{1,3}-\d{3}$", c) and c not in seen:
                    seen.add(c)
                    # Get the name from the same row
                    name = max((str(x).strip() for x in row
                                if len(str(x).strip()) > 8
                                and str(x).strip() != c), key=len, default="")
                    island = ("PEMBA" if any(p in c for p in
                              ("MK","MI","CH","WE","PB")) else "UNGUJA")
                    roads.append({"code": c, "name": name,
                                  "island": island, "project": project})
        print(f"  ✓ {len(roads)} codes harvested from PP sheet ({project} KM)")
    except Exception as e:
        print(f"  ⚠  Could not read PP sheet ({project} KM): {e}")
    return roads


# ══════════════════════════════════════════════════════════════════════════════
# MAIN
# ══════════════════════════════════════════════════════════════════════════════

def main():
    print("=" * 70)
    print("  IRIS Road Code Matcher")
    print(f"  Run at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 70)

    # ── Step 1: collect sheet road data ──────────────────────────────────────
    sheet_roads: list[dict] = []

    if HAS_GOOGLE and TOKEN_FILE.exists():
        print("\n[1] Reading road codes from Google Sheets …")
        try:
            svc = auth()
            for proj, ssid in RFI_ID.items():
                rds = read_road_codes_from_sheet(svc, ssid, proj)
                if not rds:
                    # Fallback to Physical Progress sheet
                    rds = read_from_physical_progress(svc, PP_ID[proj], proj)
                sheet_roads.extend(rds)
            # Also try Physical Progress directly
            for proj, ssid in PP_ID.items():
                rds = read_from_physical_progress(svc, ssid, proj)
                # Merge (deduplicate by code)
                existing = {r["code"] for r in sheet_roads}
                for r in rds:
                    if r["code"] not in existing:
                        sheet_roads.append(r)
                        existing.add(r["code"])
        except Exception as e:
            print(f"  Google Sheets error: {e}")
    else:
        print("\n[1] Google Sheets unavailable — using built-in code list …")

    # ── Step 2: built-in fallback codes (from known Physical Progress data) ──
    BUILTIN_CODES = [
        # 163.22 KM — UNGUJA
        {"code":"NA-200","name":"Matemwe – Sun Shine Hotel (AL1 + AL2) feeder road","island":"UNGUJA","project":"163.22"},
        {"code":"CE-201","name":"Istiqaama - Zanzibar University feeder road","island":"UNGUJA","project":"163.22"},
        {"code":"SO-202","name":"Muyuni A (Kibaoni) – Muyuni C (Dindi) feeder road","island":"UNGUJA","project":"163.22"},
        {"code":"SO-203","name":"Muyuni B (Kidodoni) – Ngombeni + Skuli-Ndorobo Spur feeder road","island":"UNGUJA","project":"163.22"},
        {"code":"SO-204","name":"Dindi – Mchangamle + Wasusanya – Kilimatangi Spar feeder road","island":"UNGUJA","project":"163.22"},
        {"code":"SO-205","name":"Kidodofirishi – Kizimachakacha feeder road","island":"UNGUJA","project":"163.22"},
        {"code":"SO-206","name":"Kisokoni – Kaskazi (Kikwecheni) + Mtende Health Center Spur","island":"UNGUJA","project":"163.22"},
        {"code":"SO-207","name":"Madete Skuli – Skuli ya Kusini access road","island":"UNGUJA","project":"163.22"},
        {"code":"SO-208","name":"Mnarani – Kigaeni feeder road","island":"UNGUJA","project":"163.22"},
        {"code":"SO-209","name":"Mtendemndo – Pwani feeder road","island":"UNGUJA","project":"163.22"},
        {"code":"SO-210","name":"Mwembekungu – Kaskazi Pemba access road","island":"UNGUJA","project":"163.22"},
        {"code":"SO-211","name":"Mwembesanda – Mavua feeder road","island":"UNGUJA","project":"163.22"},
        {"code":"SO-212","name":"Nganani – Bungi – Mtegani + Kumbini + Chuini Spur","island":"UNGUJA","project":"163.22"},
        {"code":"CE-213","name":"Access road to JKU Dunga camp incl. car park","island":"UNGUJA","project":"163.22"},
        {"code":"SO-214","name":"Tasani – Mnarani – Pwani feeder road","island":"UNGUJA","project":"163.22"},
        {"code":"SO-215","name":"Wadini – Cheju feeder road","island":"UNGUJA","project":"163.22"},
        {"code":"WB-216","name":"Fuoni Mskiti wa Waarabu – Fuoni Michenzani + Spur","island":"UNGUJA","project":"163.22"},
        {"code":"NB-217","name":"Donge Mbiji – Karange + Spur","island":"UNGUJA","project":"163.22"},
        {"code":"WB-218","name":"Fuoni Mtundani + Jangamizi feeder road","island":"UNGUJA","project":"163.22"},
        {"code":"CE-219","name":"Jumbi Makonteni (Meli Tisa) and spur","island":"UNGUJA","project":"163.22"},
        {"code":"WB-220","name":"Kisauni Mazurui – Tangi la Maji feeder road + Spur","island":"UNGUJA","project":"163.22"},
        {"code":"WA-221","name":"Pili Space (Dole) – Shamba la Bishafika feeder road","island":"UNGUJA","project":"163.22"},
        {"code":"WA-222","name":"Flawa – Suleiman Kondo mskitini access road","island":"UNGUJA","project":"163.22"},
        {"code":"WA-223","name":"Ziwani Police access roads","island":"UNGUJA","project":"163.22"},
        {"code":"WA-224","name":"Kilimani Prison Headquarters access roads","island":"UNGUJA","project":"163.22"},
        {"code":"WB-225","name":"Bavuai camp (petrol station) – Migombani JWTZ camp feeder road","island":"UNGUJA","project":"163.22"},
        {"code":"WB-226","name":"Access Road to Mtoni JWTZ camp","island":"UNGUJA","project":"163.22"},
        {"code":"AR-227","name":"Mikungunini Access Road (Matemwe)","island":"UNGUJA","project":"163.22"},
        {"code":"WB-229","name":"Mitondooni Bakery access road","island":"UNGUJA","project":"163.22"},
        {"code":"WB-230","name":"Kisauni Mtundani feeder road","island":"UNGUJA","project":"163.22"},
        {"code":"WB-231","name":"Fuoni Melisita Chunga – Mwera Melisita (Jia la Said Ali)","island":"UNGUJA","project":"163.22"},
        {"code":"WB-232","name":"Maungani Vijijini – Shakani","island":"UNGUJA","project":"163.22"},
        # 163.22 KM — PEMBA
        {"code":"MK-233","name":"Kiwani – Sokani – Mchochani feeder road","island":"PEMBA","project":"163.22"},
        {"code":"MK-234","name":"Kichaka – Chwaaka Kubwa + Chwaka Ndogo Spar","island":"PEMBA","project":"163.22"},
        {"code":"MK-235","name":"Kichungwani – Kwa Sanani feeder road","island":"PEMBA","project":"163.22"},
        {"code":"MK-236","name":"Mauani – School – Kirimdado – Spar feeder road","island":"PEMBA","project":"163.22"},
        {"code":"MK-237","name":"Mahuduthi – Sipwese + Mto Mkuu Spar","island":"PEMBA","project":"163.22"},
        {"code":"MK-238","name":"Minazini – Shikazako feeder road","island":"PEMBA","project":"163.22"},
        {"code":"MK-239","name":"Mtamba – Kitondoo + Tomondo feeder road","island":"PEMBA","project":"163.22"},
        {"code":"MI-240","name":"Micheweni Access Road (Kiuyu – Micheweni)","island":"PEMBA","project":"163.22"},
        {"code":"CH-241","name":"Michwewni – Mchupwe (Micheweni Access Road)","island":"PEMBA","project":"163.22"},
        {"code":"MK-242","name":"Mizingani – Wambaa + Fundu Lagoon","island":"PEMBA","project":"163.22"},
        {"code":"MK-243","name":"Chambani (kwa Mbunge) – JKU – Mtangani","island":"PEMBA","project":"163.22"},
        {"code":"MI-244","name":"Konde Sokoni – Kiuyu Kwa Manda","island":"PEMBA","project":"163.22"},
        {"code":"MK-245","name":"Kuyuni – Kalani feeder road","island":"PEMBA","project":"163.22"},
        {"code":"WE-246","name":"Mianzini – Mabatini feeder road","island":"PEMBA","project":"163.22"},
        {"code":"MK-247","name":"Kengeja – Minazini – Kangani feeder road","island":"PEMBA","project":"163.22"},
        {"code":"CH-248","name":"Chanjaani – Mitamani – Pujuni feeder road","island":"PEMBA","project":"163.22"},
        {"code":"MK-249","name":"Mtambile (Police Post) – Kangaani feeder road","island":"PEMBA","project":"163.22"},
        {"code":"MK-250","name":"Kenya Kwa Mtora feeder road","island":"PEMBA","project":"163.22"},
        {"code":"WE-251","name":"Wete Access Road","island":"PEMBA","project":"163.22"},
        {"code":"MK-252","name":"Pujini (Kilimo) – Mkandandume – Kibaridi feeder road","island":"PEMBA","project":"163.22"},
        # 65.5 KM
        {"code":"102","name":"Kianga – Istiqama feeder road","island":"UNGUJA","project":"65.5"},
        {"code":"103","name":"Maskani – Kikaangoni feeder road","island":"UNGUJA","project":"65.5"},
        {"code":"104","name":"Mwanakwerekwe – Bubujiko road","island":"UNGUJA","project":"65.5"},
        {"code":"105","name":"Mwanakwerekwe bypass","island":"UNGUJA","project":"65.5"},
    ]

    # Merge built-ins (only add codes not already loaded from sheets)
    existing_codes = {r["code"] for r in sheet_roads}
    for r in BUILTIN_CODES:
        if r["code"] not in existing_codes:
            sheet_roads.append(r)

    print(f"\n  Total sheet road entries: {len(sheet_roads)}")

    # ── Step 3: run matching ──────────────────────────────────────────────────
    print("\n[2] Running fuzzy matching …")
    print(f"  Thresholds: base={THRESHOLD_BASE}, high={THRESHOLD_HIGH}, "
          f"low={THRESHOLD_LOW}, min_gap={MIN_GAP}")

    results = []
    stats   = defaultdict(int)

    for kmz in KMZ_ROADS:
        # Filter candidates: same island if possible, else allow all
        island_upper = kmz["island"].upper()
        candidates   = [r for r in sheet_roads
                        if r["island"].upper() == island_upper
                        or r["island"].upper() in ("", "UNGUJA", "PEMBA")]

        res = match_road(kmz, candidates)
        res["kmz_name"] = kmz["name"]
        res["kmz_num"]  = kmz.get("num","")
        res["kmz_island"] = kmz["island"]
        results.append(res)
        stats[res["status"]] += 1

    # ── Step 4: print results ─────────────────────────────────────────────────
    print("\n" + "─"*70)
    print(f"  {'KMZ ROAD NAME':<45} {'CODE':<10} {'CONF':>6}  STATUS")
    print("─"*70)

    for r in results:
        code  = r.get("code","—")
        conf  = f"{r.get('confidence',0):.0%}" if r.get("confidence") else "—"
        name  = r["kmz_name"][:45]
        print(f"  {name:<45} {code:<10} {conf:>6}  {r['status']}")

    print("─"*70)
    print(f"\n  MATCHED:         {stats['MATCHED']}")
    print(f"  LOW_CONFIDENCE:  {stats['LOW_CONFIDENCE']}")
    print(f"  AMBIGUOUS:       {stats['AMBIGUOUS']}")
    print(f"  NO_MATCH:        {stats['NO_MATCH']}")
    print(f"  NO_CANDIDATES:   {stats['NO_CANDIDATES']}")

    # ── Step 5: save JSON output ──────────────────────────────────────────────
    mapping_out = {}
    for r in results:
        key = r["kmz_name"]
        mapping_out[key] = {
            "kmz_num":    r["kmz_num"],
            "kmz_island": r["kmz_island"],
            "status":     r["status"],
            "code":       r.get("code",""),
            "pp_name":    r.get("name",""),
            "island":     r.get("island",""),
            "project":    r.get("project",""),
            "confidence": round(r.get("confidence",0), 4),
            "top_matches": [
                {"code": m["code"], "name": m["name"][:60], "score": round(m["score"],4)}
                for m in r.get("matches",[])[:3]
            ],
        }

    # Also include a reverse lookup: code → kmz_name(s)
    code_to_kmz: dict[str, list] = defaultdict(list)
    for kmz_name, info in mapping_out.items():
        if info["code"]:
            code_to_kmz[info["code"]].append(kmz_name)

    output = {
        "generated":   datetime.now().isoformat(),
        "stats":       dict(stats),
        "by_kmz_name": mapping_out,
        "by_code":     {k: v for k, v in code_to_kmz.items()},
    }

    OUTPUT_JSON.write_text(json.dumps(output, indent=2, ensure_ascii=False), encoding="utf-8")
    print(f"\n  ✓ Mapping saved → {OUTPUT_JSON}")

    # ── Step 6: save human-readable report ────────────────────────────────────
    lines = [
        "IRIS ROAD CODE MATCHING REPORT",
        f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
        f"KMZ roads: {len(KMZ_ROADS)}   Sheet codes: {len(sheet_roads)}",
        "=" * 80,
        "",
        "MATCHED",
        "-" * 80,
    ]
    for r in results:
        if r["status"] == "MATCHED":
            lines.append(
                f"  {r['kmz_name']:<55} → {r.get('code',''):<10}"
                f"  [{r.get('confidence',0):.0%}]  {r.get('name','')[:40]}"
            )
    lines += ["", "LOW CONFIDENCE (review recommended)", "-"*80]
    for r in results:
        if r["status"] == "LOW_CONFIDENCE":
            lines.append(
                f"  {r['kmz_name']:<55} → {r.get('code',''):<10}"
                f"  [{r.get('confidence',0):.0%}]"
            )
            for m in r.get("matches",[])[:3]:
                lines.append(f"      option: {m['code']} ({m['score']:.0%}) {m['name'][:50]}")

    lines += ["", "AMBIGUOUS (manual resolution needed)", "-"*80]
    for r in results:
        if r["status"] == "AMBIGUOUS":
            lines.append(f"  {r['kmz_name']}")
            for m in r.get("matches",[])[:5]:
                lines.append(f"      {m['code']:<10} {m['score']:.0%}  {m['name'][:55]}")

    lines += ["", "NO MATCH", "-"*80]
    for r in results:
        if r["status"] in ("NO_MATCH","NO_CANDIDATES"):
            lines.append(f"  {r['kmz_name']}")
            for m in r.get("matches",[])[:3]:
                lines.append(f"      best: {m['code']:<10} {m['score']:.0%}  {m['name'][:55]}")

    OUTPUT_REPORT.write_text("\n".join(lines), encoding="utf-8")
    print(f"  ✓ Report saved  → {OUTPUT_REPORT}")
    print(f"\n  ✓ Done at {datetime.now().strftime('%H:%M:%S')}")


if __name__ == "__main__":
    main()
