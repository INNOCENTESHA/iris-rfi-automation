"""
IRIS RFI Review Agent (Claude Batch API)
========================================
Reviews every new site and laboratory RFI for the three IRIS projects and
writes flags to each project's Google Sheet:

    65.5 KM project          -> 65.5 RFI Detailed Log, tab "AGENT REVIEW"
    163.22 KM Feeder & Rural -> 163.22 RFI Detailed Log, tab "AGENT REVIEW"
    Kizimkazi Road Network   -> Kizimkazi RFI Review sheet, tab "AGENT REVIEW"

It reads the attachments that download_rfi_attachments.py has already saved
to RFI_ZIP_DIR (ZIP, RAR, PDF, DOCX). It never extracts archives to disk and
never changes the RFI Detailed Log tables; it only appends rows to its own
"AGENT REVIEW" and "WEEKLY SUMMARY" tabs.

How a run works
    1. Collect any batch submitted by an earlier run that has now finished,
       and write its results to the sheets.
    2. Find RFI documents not reviewed yet, assign each to a project,
       and submit them to the Message Batches API (50% of normal price).
    3. Wait up to REVIEW_WAIT_MINUTES for that batch. If it is not done,
       its ID is kept in the state file and the next run collects it.
    4. With --weekly, write a summary of the last 7 days per project to
       the "WEEKLY SUMMARY" tab.

ENVIRONMENT VARIABLES
    ANTHROPIC_API_KEY          required  Claude Console API key
    RFI_ZIP_DIR                required  folder the downloader writes to
    GOOGLE_SERVICE_ACCOUNT_PATH required (unless --dry-run) service-account JSON
    REVIEW_STATE_PATH          optional  default <RFI_ZIP_DIR>/../state/review_state.json
    REVIEW_MODEL               optional  default claude-sonnet-5-5
    REVIEW_WAIT_MINUTES        optional  default 40
    REVIEW_MAX_DOCS            optional  cap per run, default 300
    REVIEW_SINCE_DAYS          optional  only files dated in the last N days (default 14, "all" = no limit)
    IRIS_REVIEW_ID_65_5        optional  override workbook ID
    IRIS_REVIEW_ID_163_22      optional  override workbook ID
    IRIS_REVIEW_ID_KIZIMKAZI   optional  override workbook ID

USAGE
    python rfi_review_agent.py               # daily review
    python rfi_review_agent.py --weekly      # daily review + weekly summary
    python rfi_review_agent.py --dry-run     # list what would be sent, no API, no sheets
    python rfi_review_agent.py --list        # show documents found and their project

EXIT CODES
    0 success   1 run finished with errors   2 configuration error
"""

from __future__ import annotations

import argparse
import base64
import hashlib
import io
import json
import os
import re
import sys
import time
import zipfile
from datetime import datetime, timedelta, timezone

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

EAT = timezone(timedelta(hours=3))

# ── Project routing ───────────────────────────────────────────────────────────
PROJECTS = {
    "65.5": {
        "label": "65.5 KM Roads (Unguja and Pemba)",
        "sheet_id": os.environ.get("IRIS_REVIEW_ID_65_5")
        or "1WV58zo-jYgqOTBGOtMAjjlHC3WKdQ83s-XdRFlxZLLU",
    },
    "163.22": {
        "label": "163.22 KM Feeder and Rural Roads",
        "sheet_id": os.environ.get("IRIS_REVIEW_ID_163_22")
        or "1IjN4W79VEZrAjevG5zZVKyJPuCdGbm_xAbcttTNjfOA",
    },
    "kizimkazi": {
        "label": "Design and Build of the Kizimkazi Road Network, Unguja",
        "sheet_id": os.environ.get("IRIS_REVIEW_ID_KIZIMKAZI")
        or "15tSQgfQ0tT45nnD-Z1E6YUAJzyg_IFpV_Kqpt_ofF7A",
    },
}

REVIEW_TAB = "AGENT REVIEW"
SUMMARY_TAB = "WEEKLY SUMMARY"
REVIEW_HEADER = [
    "Reviewed On", "RFI No.", "RFI Date", "Road Code", "Road Name", "Island",
    "Chainage", "Activity", "Type", "Result", "Test Values", "Flag", "Severity",
    "Issue", "Recommended Action", "Source File",
]
SUMMARY_HEADER = [
    "Week Ending", "RFIs Reviewed", "High Flags", "Medium Flags", "Summary",
    "Open Issues", "Generated On",
]

DOC_EXT = (".pdf", ".docx")
ARCHIVE_EXT = (".zip", ".rar")
MAX_TEXT_CHARS = 30000          # per document sent to Claude
MAX_PDF_BYTES = 8 * 1024 * 1024 # larger scanned PDFs are skipped, not sent
MIN_TEXT_CHARS = 200            # below this a PDF is treated as a scan

_RC_RE = re.compile(r"\b([A-Z]{2})[-\s]?([12]\d{2})\b")


def classify_project(archive_name: str, inner_name: str, text: str) -> str:
    """Return '65.5', '163.22', 'kizimkazi' or '' (unknown)."""
    names = f"{archive_name} {inner_name}".upper()
    head = (text or "")[:6000].upper()
    if "KIZIMKAZI" in names or "KIZIMKAZI" in head or re.search(r"\bKZP[-_\s]", names + " " + head):
        return "kizimkazi"
    if re.search(r"163[.,_ ]?22", names):
        return "163.22"
    if re.search(r"\b65[.,_ ]5\b", names):
        return "65.5"
    if re.search(r"163[.,]22\s*KM", head):
        return "163.22"
    if re.search(r"65[.,]5\s*KM", head):
        return "65.5"
    if re.match(r"^IRIS[-_\s]+QAQC", inner_name.upper()):
        return "163.22"
    if re.match(r"^IR[-_\s]+QAQC", inner_name.upper()):
        return "65.5"
    m = _RC_RE.search(names) or _RC_RE.search(head)
    if m:
        return "65.5" if m.group(2).startswith("1") else "163.22"
    m = re.search(r"\bTHE\s+(\d+(?:[.,]\d+)?)\s*KM\b", head)
    if m:
        km = float(m.group(1).replace(",", "."))
        return "163.22" if km > 65.5 else "65.5"
    return ""


# ── Reading documents (in memory only) ────────────────────────────────────────
def pdf_text(data: bytes) -> str:
    import pdfplumber
    try:
        with pdfplumber.open(io.BytesIO(data)) as pdf:
            return "\n".join((p.extract_text() or "") for p in pdf.pages[:15])
    except Exception:
        return ""


def docx_text(data: bytes) -> str:
    import docx
    try:
        d = docx.Document(io.BytesIO(data))
        parts = [p.text for p in d.paragraphs]
        for t in d.tables:
            for row in t.rows:
                parts.append(" | ".join(c.text.strip() for c in row.cells))
        return "\n".join(parts)
    except Exception:
        return ""


def iter_archive(path: str):
    """Yield (inner_name, bytes) for PDF/DOCX members, recursing into nested ZIPs."""
    def from_zip(zf, prefix=""):
        for info in zf.infolist():
            n = info.filename
            if info.is_dir() or "__MACOSX" in n:
                continue
            low = n.lower()
            if low.endswith(DOC_EXT):
                yield prefix + n, zf.read(info)
            elif low.endswith(".zip"):
                try:
                    with zipfile.ZipFile(io.BytesIO(zf.read(info))) as inner:
                        yield from from_zip(inner, prefix + n + "/")
                except zipfile.BadZipFile:
                    continue

    low = path.lower()
    if low.endswith(".zip"):
        try:
            with zipfile.ZipFile(path) as zf:
                yield from from_zip(zf)
        except zipfile.BadZipFile:
            print(f"  ! bad ZIP skipped: {os.path.basename(path)}")
    elif low.endswith(".rar"):
        try:
            import rarfile
            with rarfile.RarFile(path) as rf:
                for info in rf.infolist():
                    if not info.is_dir() and info.filename.lower().endswith(DOC_EXT):
                        yield info.filename, rf.read(info)
        except Exception as e:
            print(f"  ! RAR skipped ({e}): {os.path.basename(path)}")


def collect_documents(zip_dir: str) -> list[dict]:
    """Every RFI document in the download folder, with a content hash."""
    docs = []
    for fname in sorted(os.listdir(zip_dir)):
        path = os.path.join(zip_dir, fname)
        if not os.path.isfile(path):
            continue
        low = fname.lower()
        if low.endswith(ARCHIVE_EXT):
            members = iter_archive(path)
        elif low.endswith(DOC_EXT):
            members = [(fname, open(path, "rb").read())]
            fname = ""
        else:
            continue
        for inner, data in members:
            docs.append({
                "archive": fname,
                "inner": inner,
                "data": data,
                "sha": hashlib.sha256(data).hexdigest(),
            })
    return docs


_NAME_DATE_RE = re.compile(r"(\d{2})[.\-_](\d{2})[.\-_](\d{4})")


def include_by_date(doc: dict, since) -> bool:
    """Keep documents whose file or archive name is dated on/after `since` (undated are kept)."""
    if since is None:
        return True
    m = _NAME_DATE_RE.search(doc["inner"]) or _NAME_DATE_RE.search(doc["archive"])
    if not m:
        return True
    try:
        d = datetime(int(m.group(3)), int(m.group(2)), int(m.group(1))).date()
    except ValueError:
        return True
    return d >= since


# ── Claude request building ───────────────────────────────────────────────────
SYSTEM_PROMPT = """You are a materials and quality engineer reviewing Requests for Inspection (RFIs) \
for road construction in Zanzibar (Unguja and Pemba) under the IRIS-LUPTAN JV. \
The contractor is IRIS Construction and Trade Limited; the supervision consultant is \
Luptan Consult Ltd / Dar Al Handasah; the client is the Ministry of Infrastructure and \
Transportation, Revolutionary Government of Zanzibar. Specifications follow the Tanzanian \
SSRW 2000 and PMDM 1999 manuals and British Standards.

Read the RFI and report only what the document shows. Never invent test values, limits or results.
Flag the RFI when any of these apply:
- the inspection was rejected, not approved, approved with comments, or a re-inspection is requested
- a test result is outside the specified limit stated on the RFI (e.g. compaction below the required %MDD,
  CBR below the required value, layer thickness or level outside tolerance, concrete cube strength below grade)
- required information is missing (no chainage, no test results on a lab RFI, unsigned, no date)
- the same chainage or layer appears to be re-submitted after a failure
Most RFIs are the contractor's submission, sent before inspection. The form always prints the
options "Approved for Further Work / Not Approved, RFI to be re submitted / Approved with Comment";
these printed words are NOT a decision. Only report APPROVED, REJECTED or APPROVED WITH COMMENTS
when a box is ticked, circled or stamped, or the Resident Engineer / consultant has written a decision.
If the consultant sections are blank, the result is PENDING; do not flag a pending RFI for that alone.
"N/A" road name or chainage is normal for laboratory sampling at the camp or a borrow pit.
Severity: HIGH = failed test or rejected work; MEDIUM = comments, re-inspection, missing key data;
LOW = minor clerical issue; NONE = approved and compliant.

Answer with one JSON object and nothing else."""

JSON_SPEC = """Return this JSON (use "" when not on the document):
{
  "rfi_no": "", "rfi_date": "DD/MM/YYYY", "road_code": "", "road_name": "",
  "island": "Unguja|Pemba|", "chainage": "", "activity": "", "type": "SITE|LAB",
  "result": "APPROVED|APPROVED WITH COMMENTS|REJECTED|PENDING|NOT STATED",
  "test_values": "short list of key results vs limits, e.g. 'MDD 97.1% (min 95%); CBR 62 (min 80)'",
  "flag": true, "severity": "HIGH|MEDIUM|LOW|NONE",
  "issue": "one or two sentences, empty if none",
  "recommended_action": "one sentence, empty if none"
}"""


def custom_id_for(sha: str) -> str:
    return "rfi-" + sha[:56]


def build_request(doc: dict, project: str, model: str) -> dict | None:
    name = doc["inner"]
    data = doc["data"]
    context = (f"Project: {PROJECTS[project]['label']}\n"
               f"Archive: {doc['archive'] or '-'}\nFile: {name}\n\n{JSON_SPEC}")
    if name.lower().endswith(".pdf"):
        text = doc.get("text") if doc.get("text") is not None else pdf_text(data)
        if len(data) <= MAX_PDF_BYTES:
            # Send the PDF itself so ticks, stamps, signatures and scans are visible,
            # plus the extracted text when there is any.
            extra = (f"Extracted text (may miss ticks and handwriting):\n{text[:MAX_TEXT_CHARS]}\n\n"
                     if len(text.strip()) >= MIN_TEXT_CHARS else "")
            content = [
                {"type": "document", "source": {
                    "type": "base64", "media_type": "application/pdf",
                    "data": base64.b64encode(data).decode()}},
                {"type": "text", "text": extra + context},
            ]
        elif len(text.strip()) >= MIN_TEXT_CHARS:
            content = [{"type": "text", "text": f"RFI text:\n{text[:MAX_TEXT_CHARS]}\n\n{context}"}]
        else:
            print(f"  ! scanned PDF too large, skipped: {name}")
            return None
    else:
        text = doc.get("text") if doc.get("text") is not None else docx_text(data)
        if not text.strip():
            return None
        content = [{"type": "text", "text": f"RFI text:\n{text[:MAX_TEXT_CHARS]}\n\n{context}"}]
    return {
        "custom_id": custom_id_for(doc["sha"]),
        "params": {
            "model": model,
            "max_tokens": 1024,
            "system": SYSTEM_PROMPT,
            "messages": [{"role": "user", "content": content}],
        },
    }


def parse_review(text: str) -> dict:
    """Pull the JSON object out of Claude's reply."""
    m = re.search(r"\{.*\}", text or "", re.DOTALL)
    if not m:
        return {"flag": True, "severity": "LOW", "issue": "Agent reply could not be read."}
    try:
        return json.loads(m.group(0))
    except json.JSONDecodeError:
        return {"flag": True, "severity": "LOW", "issue": "Agent reply could not be read."}


def _txt(v) -> str:
    """Force a value to stay text when written with USER_ENTERED."""
    v = "" if v is None else str(v).strip()
    return "'" + v if v else ""


def review_to_row(r: dict, meta: dict, reviewed_on: str) -> list:
    sev = str(r.get("severity", "")).upper()
    if sev not in ("HIGH", "MEDIUM", "LOW", "NONE"):
        sev = "LOW"
    flag = r.get("flag")
    flag = "YES" if (flag is True or str(flag).lower() in ("true", "yes")) else "NO"
    src = meta["inner"] if not meta.get("archive") else f"{meta['archive']} > {meta['inner']}"
    fields = [r.get(k, "") for k in ("rfi_no", "rfi_date", "road_code", "road_name",
                                      "island", "chainage", "activity", "type", "result",
                                      "test_values")]
    fields += [flag, sev, r.get("issue", ""), r.get("recommended_action", ""), src]
    return [reviewed_on] + [_txt(f) for f in fields]


# ── State ─────────────────────────────────────────────────────────────────────
def load_state(path: str) -> dict:
    if os.path.exists(path):
        with open(path, encoding="utf-8") as f:
            s = json.load(f)
    else:
        s = {}
    s.setdefault("reviewed", {})        # sha -> {"project", "date"}
    s.setdefault("pending_batches", {}) # batch_id -> {custom_id: meta}
    return s


def save_state(path: str, state: dict) -> None:
    os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(state, f, indent=1)
    os.replace(tmp, path)


# ── Google Sheets ─────────────────────────────────────────────────────────────
def sheets_client(sa_path: str):
    import gspread
    return gspread.service_account(filename=sa_path)


def get_tab(gc, sheet_id: str, title: str, header: list):
    import gspread
    sh = gc.open_by_key(sheet_id)
    try:
        ws = sh.worksheet(title)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=title, rows=1000, cols=len(header))
        ws.update(values=[header], range_name="A1")
        ws.freeze(rows=1)
    if not ws.row_values(1):
        ws.update(values=[header], range_name="A1")
    return ws


def append_rows(gc, project: str, rows: list) -> None:
    if not rows:
        return
    ws = get_tab(gc, PROJECTS[project]["sheet_id"], REVIEW_TAB, REVIEW_HEADER)
    ws.append_rows(rows, value_input_option="USER_ENTERED", insert_data_option="INSERT_ROWS",
                   table_range="A1")
    print(f"  wrote {len(rows)} row(s) to {PROJECTS[project]['label']} / {REVIEW_TAB}")


# ── Batch API ─────────────────────────────────────────────────────────────────
def collect_batch(client, batch_id: str, metas: dict, state: dict, gc, dry: bool) -> bool:
    """Write a finished batch's results. Returns True when the batch is done."""
    batch = client.messages.batches.retrieve(batch_id)
    if batch.processing_status != "ended":
        c = batch.request_counts
        print(f"  batch {batch_id}: {batch.processing_status} "
              f"(processing {c.processing}, done {c.succeeded}, errors {c.errored})")
        return False
    now = datetime.now(EAT)
    reviewed_on = f"=DATE({now.year},{now.month},{now.day})"   # a real date in any locale
    rows_by_project: dict[str, list] = {}
    failed = 0
    for entry in client.messages.batches.results(batch_id):
        meta = metas.get(entry.custom_id)
        if not meta:
            continue
        if entry.result.type != "succeeded":
            failed += 1
            continue  # not marked reviewed, so it is retried next run
        text = "".join(b.text for b in entry.result.message.content if b.type == "text")
        row = review_to_row(parse_review(text), meta, reviewed_on)
        rows_by_project.setdefault(meta["project"], []).append(row)
        state["reviewed"][meta["sha"]] = {"project": meta["project"], "date": now.strftime("%Y-%m-%d")}
    if not dry:
        for project, rows in rows_by_project.items():
            append_rows(gc, project, rows)
    print(f"  batch {batch_id}: {sum(len(r) for r in rows_by_project.values())} reviewed, "
          f"{failed} failed (will retry)")
    return True


def cell_date(v):
    """Sheets serial number or a dd/mm/yyyy / yyyy-mm-dd string -> date."""
    if isinstance(v, (int, float)) and v > 30000:
        return (datetime(1899, 12, 30) + timedelta(days=float(v))).date()
    for fmt in ("%d/%m/%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(str(v).strip(), fmt).date()
        except ValueError:
            pass
    return None


# ── Weekly summary ────────────────────────────────────────────────────────────
def weekly_summary(client, gc, model: str) -> None:
    today = datetime.now(EAT).date()
    week_start = today - timedelta(days=6)
    for project, cfg in PROJECTS.items():
        ws = get_tab(gc, cfg["sheet_id"], REVIEW_TAB, REVIEW_HEADER)
        rows = ws.get_all_values(value_render_option="UNFORMATTED_VALUE")[1:]
        recent = []
        for r in rows:
            d = cell_date(r[0] if r else "")
            if d is None:
                continue
            r = [str(c) for c in r] + [""] * len(REVIEW_HEADER)
            if week_start <= d <= today:
                recent.append(dict(zip(REVIEW_HEADER, r)))
        sws = get_tab(gc, cfg["sheet_id"], SUMMARY_TAB, SUMMARY_HEADER)
        if not recent:
            summary, open_issues = "No new RFIs were reviewed this week.", ""
        else:
            lines = "\n".join(
                f"- {x['RFI No.']} | {x['Road Code']} {x['Road Name']} | {x['Chainage']} | "
                f"{x['Activity']} | {x['Type']} | {x['Result']} | {x['Severity']} | {x['Issue']}"
                for x in recent)
            msg = client.messages.create(
                model=model, max_tokens=1500,
                system=("You write short weekly quality reports for a road construction site "
                        "engineer. Plain English, no headings, no emoji."),
                messages=[{"role": "user", "content": (
                    f"Project: {cfg['label']}\nWeek: {week_start:%d/%m/%Y} to {today:%d/%m/%Y}\n"
                    f"RFIs reviewed this week:\n{lines}\n\n"
                    "Return JSON: {\"summary\": \"4-6 sentences on what was inspected, which roads "
                    "and layers progressed, and the pattern of failures\", \"open_issues\": "
                    "\"numbered list of unresolved HIGH and MEDIUM items with RFI number and road\"}")}],
            )
            out = parse_review("".join(b.text for b in msg.content if b.type == "text"))
            summary = out.get("summary", "")
            open_issues = out.get("open_issues", "")
            if isinstance(open_issues, list):
                open_issues = "\n".join(str(i) for i in open_issues)
        n = len(sws.get_all_values()) + 1
        lo, hi = week_start.strftime("%Y-%m-%d"), today.strftime("%Y-%m-%d")
        rng = (f"'{REVIEW_TAB}'!$A:$A,\">=\"&DATEVALUE(\"{lo}\"),"
               f"'{REVIEW_TAB}'!$A:$A,\"<=\"&DATEVALUE(\"{hi}\")")
        sws.update(values=[[
            f"=DATE({today.year},{today.month},{today.day})",
            f"=COUNTIFS({rng})",
            f"=COUNTIFS({rng},'{REVIEW_TAB}'!$M:$M,\"HIGH\")",
            f"=COUNTIFS({rng},'{REVIEW_TAB}'!$M:$M,\"MEDIUM\")",
            _txt(summary), _txt(open_issues), _txt(datetime.now(EAT).strftime("%d/%m/%Y %H:%M")),
        ]], range_name=f"A{n}", value_input_option="USER_ENTERED")
        print(f"  weekly summary written for {cfg['label']} ({len(recent)} RFIs)")


# ── Main ──────────────────────────────────────────────────────────────────────
def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[1])
    ap.add_argument("--weekly", action="store_true", help="also write the weekly summary")
    ap.add_argument("--dry-run", action="store_true", help="no API calls, no sheet writes")
    ap.add_argument("--list", action="store_true", help="list documents and projects, then exit")
    args = ap.parse_args()

    zip_dir = os.environ.get("RFI_ZIP_DIR", "").strip()
    api_key = os.environ.get("ANTHROPIC_API_KEY", "").strip()
    sa_path = os.environ.get("GOOGLE_SERVICE_ACCOUNT_PATH", "").strip()
    model = os.environ.get("REVIEW_MODEL", "claude-sonnet-5-5").strip()
    wait_min = float(os.environ.get("REVIEW_WAIT_MINUTES", "40"))
    max_docs = int(os.environ.get("REVIEW_MAX_DOCS", "300"))
    since_days = os.environ.get("REVIEW_SINCE_DAYS", "14").strip()
    since = None if since_days in ("", "0", "all") else (
        datetime.now(EAT).date() - timedelta(days=int(since_days)))
    state_path = os.environ.get("REVIEW_STATE_PATH", "").strip() or os.path.join(
        os.path.dirname(os.path.abspath(zip_dir or ".")), "state", "review_state.json")

    missing = [n for n, v in (("RFI_ZIP_DIR", zip_dir),) if not v]
    if not (args.dry_run or args.list):
        missing += [n for n, v in (("ANTHROPIC_API_KEY", api_key),
                                   ("GOOGLE_SERVICE_ACCOUNT_PATH", sa_path)) if not v]
    if missing:
        print("Configuration error - missing: " + ", ".join(missing))
        return 2
    if not os.path.isdir(zip_dir):
        print(f"Configuration error - RFI_ZIP_DIR does not exist: {zip_dir}")
        return 2

    state = load_state(state_path)
    errors = 0

    print(f"Scanning {zip_dir} ...")
    docs = collect_documents(zip_dir)
    pending_shas = {m["sha"] for b in state["pending_batches"].values() for m in b.values()}
    new_docs, unknown = [], []
    for d in docs:
        if d["sha"] in state["reviewed"] or d["sha"] in pending_shas:
            continue
        if not include_by_date(d, since):
            continue
        text = pdf_text(d["data"]) if d["inner"].lower().endswith(".pdf") else docx_text(d["data"])
        d["text"] = text
        d["project"] = classify_project(d["archive"], d["inner"], text)
        (new_docs if d["project"] else unknown).append(d)
    # Same file in two archives -> review once.
    seen, unique = set(), []
    for d in new_docs:
        if d["sha"] not in seen:
            seen.add(d["sha"])
            unique.append(d)
    new_docs = unique[:max_docs]

    counts = {p: sum(1 for d in new_docs if d["project"] == p) for p in PROJECTS}
    print(f"Documents found: {len(docs)}; new: {len(new_docs)} "
          f"(65.5: {counts['65.5']}, 163.22: {counts['163.22']}, Kizimkazi: {counts['kizimkazi']}); "
          f"project unknown: {len(unknown)}")
    for d in unknown[:20]:
        print(f"  ? no project: {d['archive']} > {d['inner']}")

    if args.list:
        for d in new_docs:
            print(f"  [{d['project']:>9}] {d['archive']} > {d['inner']}")
        return 0

    requests, metas = [], {}
    for d in new_docs:
        req = build_request(d, d["project"], model)
        if req:
            requests.append(req)
            metas[req["custom_id"]] = {"sha": d["sha"], "project": d["project"],
                                       "archive": d["archive"], "inner": d["inner"]}
    if args.dry_run:
        print(f"DRY RUN: would submit {len(requests)} request(s) with model {model}.")
        return 0

    import anthropic
    client = anthropic.Anthropic(api_key=api_key)
    gc = sheets_client(sa_path)

    # 1. Collect batches left over from earlier runs.
    for batch_id in list(state["pending_batches"]):
        try:
            if collect_batch(client, batch_id, state["pending_batches"][batch_id], state, gc, False):
                del state["pending_batches"][batch_id]
                save_state(state_path, state)
        except Exception as e:
            errors += 1
            print(f"  ! could not collect batch {batch_id}: {e}")

    # 2. Submit new documents.
    if requests:
        batch = client.messages.batches.create(requests=requests)
        state["pending_batches"][batch.id] = metas
        save_state(state_path, state)
        print(f"Submitted batch {batch.id} with {len(requests)} RFI(s).")

        # 3. Wait for it, within the time budget.
        deadline = time.time() + wait_min * 60
        while time.time() < deadline:
            time.sleep(60)
            try:
                if collect_batch(client, batch.id, metas, state, gc, False):
                    del state["pending_batches"][batch.id]
                    save_state(state_path, state)
                    break
            except Exception as e:
                print(f"  ! polling error (will retry): {e}")
        else:
            print(f"Batch {batch.id} still running; the next run will collect it.")
    else:
        print("No new RFIs to review.")

    # 4. Weekly summary.
    if args.weekly:
        try:
            weekly_summary(client, gc, model)
        except Exception as e:
            errors += 1
            print(f"  ! weekly summary failed: {e}")

    save_state(state_path, state)
    return 1 if errors else 0


if __name__ == "__main__":
    sys.exit(main())
