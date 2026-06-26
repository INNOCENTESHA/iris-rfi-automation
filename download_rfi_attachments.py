"""
IRIS RFI Attachment Downloader
==============================
Downloads ALL attachments from emails sent by:
  Riziki Mangombe  (r.mangombe@irisas.com.tr)
  OBELINE JOLLYSTER (o.jollyster@irisas.com.tr)

Rules:
  - ONLY processes emails from those two senders (strict sender filter)
  - Downloads every attachment type (ZIP, RAR, PDF, DOCX, XLSX, ...)
  - For ZIP and RAR files: logs the contents (inventory) WITHOUT extracting
  - Skips any file already downloaded (dedup by MD5 hash + filename)
  - Phase 1: fetches ONLY email headers to pre-filter at high speed
  - Phase 2: downloads full attachments for matched emails only
  - Retries IMAP connection up to 3x with exponential back-off (5/10/20s)
  - Reports errors to n8n self-healing webhook

SETUP (one-time):
1. Go to https://myaccount.google.com/apppasswords
2. Generate a 16-character App Password for "Mail"
3. Paste it below in GMAIL_APP_PASSWORD
4. Run: python C:\\AI-Projects\\download_rfi_attachments.py
"""

import imaplib
import email
import email.header
import email.utils
import os
import re
import sys
import json
import hashlib
import zipfile
import socket
import ssl
import time
from datetime import datetime, timedelta, timezone

# ── UTF-8 safety on Windows cp1252 / PowerShell terminals ─────────────────────
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

# =============================================================================
# CONFIGURATION — fill in your details
# =============================================================================
GMAIL_ADDRESS      = os.environ.get("GMAIL_USER", "innoptesha@gmail.com")
GMAIL_APP_PASSWORD = os.environ.get("GMAIL_APP_PASSWORD", "ogue yxve arzt zngl")   # override via env in CI   # 16-char App Password (spaces ok)
OUTPUT_FOLDER      = os.environ.get("RFI_ZIP_DIR", r"D:\HIGHWAY ENGINEERING CAREER\RFI-ZIPs")
# =============================================================================

# Extensions to skip — leave empty to download EVERYTHING
SKIP_EXTENSIONS = set()

# Full-scan mode: search all emails, ignore SINCE date optimisation
FULL_SCAN_MODE = (
    "--full-scan" in sys.argv or
    os.environ.get("IRIS_FULL_SCAN", "").lower() in ("1", "true", "yes")
)

# The two senders we care about — ALL other senders are ignored
TARGETED_SENDERS = [
    "r.mangombe@irisas.com.tr",
    "o.jollyster@irisas.com.tr",
]


# ─────────────────────────────────────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────────────────────────────────────

def decode_header_value(value):
    """Decode an email header value (handles multi-charset encoded words)."""
    if value is None:
        return ""
    decoded_parts = email.header.decode_header(value)
    result = []
    for part, charset in decoded_parts:
        if isinstance(part, bytes):
            try:
                result.append(part.decode(charset or "utf-8", errors="replace"))
            except Exception:
                result.append(part.decode("latin-1", errors="replace"))
        else:
            result.append(str(part))
    return " ".join(result)


def is_from_targeted_sender(from_addr: str) -> bool:
    """Return True only when the email is from one of the two targeted senders."""
    from_up = from_addr.upper()
    return any(s.upper() in from_up for s in TARGETED_SENDERS)


def sanitize_filename(name):
    """Remove characters invalid in Windows filenames."""
    return re.sub(r'[\\/:*?"<>|]', '_', name)


def load_existing_log(log_path):
    """Load the JSON download log (returns [] if missing or corrupt)."""
    if os.path.exists(log_path):
        try:
            with open(log_path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            print("WARNING: Existing log file could not be read. Starting fresh.")
    return []


def save_log(log_path, log_entries):
    """Persist the download log to disk."""
    with open(log_path, "w", encoding="utf-8") as f:
        json.dump(log_entries, f, indent=2, ensure_ascii=False)


def get_since_date(log_entries, lookback_days: int = 7) -> str | None:
    """
    Return an IMAP SINCE date string based on the most recent download
    in the log minus lookback_days as a safety buffer.
    Returns None if the log is empty (triggers a full scan).
    """
    dates = [e.get("downloaded_at", "") for e in log_entries if e.get("downloaded_at")]
    if not dates:
        return None
    latest = max(dates)
    try:
        dt = datetime.fromisoformat(latest[:19])
        since_dt = dt - timedelta(days=lookback_days)
        return since_dt.strftime("%d-%b-%Y")
    except Exception:
        return None


# ─────────────────────────────────────────────────────────────────────────────
# Archive inventory  (list contents WITHOUT extracting)
# ─────────────────────────────────────────────────────────────────────────────

def inventory_zip(zip_path: str) -> list[str]:
    """
    Return a list of filenames inside a ZIP archive.
    Does NOT extract anything.
    """
    names = []
    try:
        with zipfile.ZipFile(zip_path, "r") as zf:
            for m in zf.infolist():
                if not m.is_dir():
                    names.append(m.filename)
    except zipfile.BadZipFile:
        print(f"      WARNING: '{os.path.basename(zip_path)}' is not a valid ZIP.")
    except Exception as e:
        print(f"      WARNING: Could not read ZIP inventory for "
              f"'{os.path.basename(zip_path)}': {e}")
    return names


def inventory_rar(rar_path: str) -> list[str]:
    """
    Return a list of filenames inside a RAR archive.
    Does NOT extract anything.
    """
    try:
        import rarfile
        with rarfile.RarFile(rar_path, "r") as rf:
            return [m.filename for m in rf.infolist() if not m.is_dir()]
    except ImportError:
        pass
    except Exception as e:
        print(f"      WARNING: Could not read RAR inventory for "
              f"'{os.path.basename(rar_path)}': {e}")
    return []


def _ensure_rarfile():
    """Auto-install rarfile if needed (for RAR inventory listing only)."""
    try:
        import rarfile  # noqa
    except ImportError:
        import subprocess as _sp
        try:
            _sp.check_call(
                [sys.executable, "-m", "pip", "install", "rarfile", "-q"],
                stdout=_sp.DEVNULL, stderr=_sp.DEVNULL,
            )
        except Exception:
            pass


_ensure_rarfile()


# ─────────────────────────────────────────────────────────────────────────────
# Path helpers
# ─────────────────────────────────────────────────────────────────────────────

def build_unique_output_path(folder, filename, email_date):
    """
    Return (output_path, safe_name) for filename inside folder.

    Duplicate-avoidance: if a file with the same sanitised name already
    exists on disk, return (None, safe_name) so the caller SKIPS it.
    We never append a _dd-mm-yyyy date stamp or _N counter suffix, because
    those create duplicate files (e.g. "...RFI's 24.06.2026_24-06-2026_1.zip").
    """
    safe_name   = sanitize_filename(filename)
    output_path = os.path.join(folder, safe_name)

    if os.path.exists(output_path):
        return None, safe_name      # already on disk -> skip, do not duplicate

    return output_path, safe_name


# ─────────────────────────────────────────────────────────────────────────────
# Pre-run verification
# ─────────────────────────────────────────────────────────────────────────────

def verify_existing_downloads(log_entries, output_folder):
    """
    Summarise the existing download log and check every file still exists.
    Returns (missing_hashes, missing_names) so missing files get re-downloaded.
    """
    if not log_entries:
        print("  No previous downloads found in log.")
        return set(), set()

    total_size  = sum(e.get("file_size", 0) for e in log_entries)
    total_files = len(log_entries)
    dates       = [e.get("downloaded_at", "") for e in log_entries
                   if e.get("downloaded_at")]
    last_date   = max(dates)[:19].replace("T", " ") if dates else "unknown"

    print(f"  Total entries in log : {total_files}")
    print(f"  Total size on disk   : {total_size / (1024*1024):.1f} MB")
    print(f"  Last downloaded at   : {last_date}")

    missing_hashes: set = set()
    missing_names:  set = set()
    missing_count       = 0
    found_count         = 0

    print()
    print("  Verifying files exist on disk ...")
    for entry in log_entries:
        saved_path = entry.get("saved_path", "")
        safe_name  = entry.get("safe_name",  "")
        file_hash  = entry.get("file_hash",  "")

        if saved_path and os.path.exists(saved_path):
            found_count += 1
        else:
            fallback = os.path.join(output_folder, safe_name) if safe_name else ""
            if fallback and os.path.exists(fallback):
                found_count += 1
            else:
                missing_count += 1
                if file_hash:
                    missing_hashes.add(file_hash)
                if safe_name:
                    missing_names.add(safe_name)

    if missing_count == 0:
        print(f"  All {found_count} previously downloaded file(s) are present. [OK]")
    else:
        print(f"  {found_count} file(s) found on disk.")
        print(f"  {missing_count} file(s) in log are MISSING -- will re-download.")

    return missing_hashes, missing_names


# ─────────────────────────────────────────────────────────────────────────────
# n8n error reporter
# ─────────────────────────────────────────────────────────────────────────────

def _post_connection_error_to_n8n(error_detail: str) -> None:
    import urllib.request
    webhook = "https://innocentesha.app.n8n.cloud/webhook/rfi-status"
    payload = json.dumps({
        "script":    "download_rfi_attachments",
        "status":    "error",
        "stage":     "IMAP_CONNECT",
        "message":   f"Connection failed after 3 retries: {error_detail}",
        "timestamp": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
    }).encode("utf-8")
    try:
        req = urllib.request.Request(
            webhook,
            data=payload,
            headers={"Content-Type": "application/json"},
            method="POST",
        )
        with urllib.request.urlopen(req, timeout=10) as resp:
            _ = resp.read()
        print("  (Connection failure reported to n8n self-healing webhook.)")
    except Exception as _e:
        print(f"  (Could not reach n8n webhook: {_e})")


# ─────────────────────────────────────────────────────────────────────────────
# Main downloader
# ─────────────────────────────────────────────────────────────────────────────

def download_attachments():
    os.makedirs(OUTPUT_FOLDER, exist_ok=True)

    app_password = GMAIL_APP_PASSWORD.replace(" ", "")
    if len(app_password) < 10:
        print("ERROR: Please set your Gmail App Password in the script.")
        return

    SEP  = "=" * 60
    SEP2 = "-" * 60

    print()
    print(SEP)
    print("IRIS RFI Attachment Downloader  (Sender-strict / No-extract mode)")
    print(SEP)
    print(f"Gmail   : {GMAIL_ADDRESS}")
    print(f"Output  : {OUTPUT_FOLDER}")
    print(f"Senders : {' | '.join(TARGETED_SENDERS)}")
    print(f"Note    : ZIPs and RARs are inventoried only — never extracted")
    print()

    # ── Load log + pre-run verification ───────────────────────────────────────
    log_path    = os.path.join(OUTPUT_FOLDER, "_download_log.json")
    log_entries = load_existing_log(log_path)

    print(SEP2)
    print("PRE-RUN VERIFICATION")
    print(SEP2)
    missing_hashes, missing_names = verify_existing_downloads(log_entries, OUTPUT_FOLDER)
    print()

    # ── Build skip-sets  (exclude missing files so they get re-downloaded) ────
    processed_emails = {
        str(e.get("msg_id"))
        for e in log_entries if e.get("msg_id")
    }
    downloaded_hashes = {
        e.get("file_hash")
        for e in log_entries
        if e.get("file_hash") and e.get("file_hash") not in missing_hashes
    }
    downloaded_names = {
        e.get("safe_name")
        for e in log_entries
        if e.get("safe_name") and e.get("safe_name") not in missing_names
    }
    missing_email_ids = {
        str(e.get("msg_id"))
        for e in log_entries
        if (e.get("file_hash") in missing_hashes or
            e.get("safe_name") in missing_names)
    }
    skip_emails = processed_emails - missing_email_ids

    # ── Connect to Gmail  (with retry + exponential back-off) ─────────────────
    print(SEP2)
    print("CONNECTING TO GMAIL")
    print(SEP2)

    _IMAP_RETRIES = 3
    _IMAP_BACKOFF = [5, 10, 20]
    _CONN_TIMEOUT = 30
    _FETCH_TIMEOUT = 180   # longer timeout for large email bodies
    mail          = None

    def _connect_imap(quiet=False):
        """
        Open a fresh IMAP connection to Gmail, log in, select All Mail.
        Returns the connected mail object, or None on failure.
        """
        for _att in range(1, _IMAP_RETRIES + 1):
            _m = None
            try:
                if not quiet:
                    print(f"Connecting to Gmail IMAP  (attempt {_att}/{_IMAP_RETRIES}) ...")
                socket.setdefaulttimeout(_CONN_TIMEOUT)
                _ctx = ssl.create_default_context()
                _m = imaplib.IMAP4_SSL("imap.gmail.com", 993, ssl_context=_ctx, timeout=_CONN_TIMEOUT)
                _m.login(GMAIL_ADDRESS, app_password)
                _m.select('"[Gmail]/All Mail"', readonly=True)
                if not quiet:
                    print("Connected successfully!")
                    print()
                return _m

            except imaplib.IMAP4.error as e:
                print(f"LOGIN FAILED: {e}")
                print("Use an App Password, not your regular Gmail password.")
                return None

            except (ConnectionResetError, ConnectionRefusedError,
                    ConnectionAbortedError, OSError, socket.error,
                    ssl.SSLError, TimeoutError) as e:
                print(f"  Network error (attempt {_att}): {type(e).__name__}: {e}")
                if _m:
                    try:
                        _m.logout()
                    except Exception:
                        pass
                if _att < _IMAP_RETRIES:
                    _wait = _IMAP_BACKOFF[_att - 1]
                    print(f"  Retrying in {_wait}s ...")
                    time.sleep(_wait)
                else:
                    _post_connection_error_to_n8n(str(e))
        return None

    mail = _connect_imap()
    if mail is None:
        print(f"\nERROR: Could not connect to Gmail IMAP after {_IMAP_RETRIES} attempts.")
        return

    # ── SINCE optimisation  (disabled in full-scan mode) ──────────────────────
    _auto_full    = len(missing_hashes) > 0 or len(missing_names) > 0
    _use_full_scan = FULL_SCAN_MODE or _auto_full

    since_date = None if _use_full_scan else get_since_date(log_entries, lookback_days=7)

    if _use_full_scan:
        reason = "--full-scan flag" if FULL_SCAN_MODE else \
                 f"{len(missing_hashes)} missing file(s) detected"
        print(f"  FULL SCAN MODE  ({reason})")

    def _since(term: str) -> str:
        if since_date:
            return f"(SINCE {since_date} {term[1:-1]})"
        return term

    # ── IMAP search: ONLY search by the two targeted senders ──────────────────
    # This ensures we never accidentally mark unrelated emails as processed.
    _base_terms = [
        f'(FROM "{s}")'
        for s in TARGETED_SENDERS
    ]
    search_terms = [_since(t) for t in _base_terms]

    print(SEP2)
    print("SEARCHING GMAIL  (sender-strict)")
    print(SEP2)
    if since_date:
        print(f"  Fast mode: since {since_date} (last download minus 7 days)")
    else:
        print("  Full scan: all emails from targeted senders")

    all_msg_ids: set = set()
    for term in search_terms:
        try:
            status, data = mail.search(None, term)
            if status == "OK" and data and data[0]:
                ids = data[0].split()
                all_msg_ids.update(ids)
                print(f"  {len(ids):>5} message(s) matched: {term}")
        except Exception as e:
            print(f"  Search error for '{term}': {e}")

    total_msgs      = len(all_msg_ids)
    skippable_count = sum(
        1 for mid in all_msg_ids
        if (mid.decode() if isinstance(mid, bytes) else str(mid)) in skip_emails
    )

    print()
    print(f"Total unique emails found    : {total_msgs}")
    print(f"Already fully processed      : {skippable_count}  (will skip)")
    print(f"To check / (re-)download     : {total_msgs - skippable_count}")
    print(SEP2)
    print()

    downloaded = skipped_dup = skipped_ext = errors = 0

    # ── Phase 1: fetch headers ONLY to pre-filter ─────────────────────────────
    sorted_ids = sorted(all_msg_ids)
    to_process = []   # (msg_id, msg_id_str, subject, from_addr, date_str)

    print(f"Phase 1 — header pre-scan ({len(sorted_ids)} emails) ...")
    header_errors = 0
    for i, msg_id in enumerate(sorted_ids, 1):
        try:
            msg_id_str = msg_id.decode() if isinstance(msg_id, bytes) else str(msg_id)
            if msg_id_str in skip_emails:
                continue

            if (i % 50) == 0 or i == len(sorted_ids):
                print(f"  Header scan: {i}/{len(sorted_ids)} "
                      f"({len(to_process)} matched so far) ...", end="\r", flush=True)

            status, hdata = mail.fetch(msg_id, "(BODY.PEEK[HEADER])")
            if status != "OK" or not hdata or not hdata[0]:
                continue

            raw_hdr = hdata[0][1] if isinstance(hdata[0], tuple) else None
            if not raw_hdr:
                continue

            hdr_msg   = email.message_from_bytes(raw_hdr)
            subject   = decode_header_value(hdr_msg.get("Subject", ""))
            from_addr = decode_header_value(hdr_msg.get("From",    ""))
            date_str  = hdr_msg.get("Date", "")

            # Only process emails from our two targeted senders
            if is_from_targeted_sender(from_addr):
                to_process.append((msg_id, msg_id_str, subject, from_addr, date_str))

        except Exception:
            header_errors += 1

    print()
    print(f"Phase 1 complete — {len(to_process)} emails from targeted senders "
          f"(scanned {len(sorted_ids)}, {header_errors} header errors).")
    print(f"Phase 2 — full download of {len(to_process)} emails ...")
    print()

    total_to_process = len(to_process)

    # ── Phase 2: full RFC822 fetch for each targeted-sender email ─────────────
    # Use a longer socket timeout during full-body fetches (large ZIPs can be slow)
    socket.setdefaulttimeout(_FETCH_TIMEOUT)

    for i, (msg_id, msg_id_str, subject, from_addr, date_str) in enumerate(to_process, 1):
        try:
            print(f"Fetching email {i}/{total_to_process} ...")  # \n so watchdog pipe sees it

            status, data = mail.fetch(msg_id, "(RFC822)")
            if status != "OK" or not data or not data[0]:
                continue

            raw_email = data[0][1] if isinstance(data[0], tuple) else data[0]
            if not isinstance(raw_email, bytes):
                for item in data:
                    if isinstance(item, tuple) and len(item) > 1:
                        raw_email = item[1]
                        break
            if not raw_email:
                continue

            msg = email.message_from_bytes(raw_email)

            subject   = decode_header_value(msg.get("Subject", subject))
            from_addr = decode_header_value(msg.get("From",    from_addr))
            date_str  = msg.get("Date", date_str)

            # Double-check sender (in case of forwarding / envelope mismatch)
            if not is_from_targeted_sender(from_addr):
                continue

            try:
                parsed_date = email.utils.parsedate_to_datetime(date_str)
                email_date  = parsed_date.strftime("%d-%m-%Y")
            except Exception:
                email_date = date_str[:10] if date_str else "unknown"

            print(f"\n[{i}/{total_to_process}] {subject[:70]}")
            print(f"          From : {from_addr[:55]}")
            print(f"          Date : {email_date}")

            email_had_attachment = False

            for part in msg.walk():
                filename = part.get_filename()
                if filename:
                    filename = decode_header_value(filename)
                if not filename:
                    continue

                # Accept all parts with a filename (inline or attached)
                file_ext = os.path.splitext(filename.lower())[1]

                if SKIP_EXTENSIONS and file_ext in SKIP_EXTENSIONS:
                    print(f"    SKIP (ext filtered): {filename}")
                    skipped_ext += 1
                    continue

                email_had_attachment = True

                try:
                    payload = part.get_payload(decode=True)
                    if not payload:
                        print(f"    WARNING: Empty payload for {filename}")
                        continue

                    file_hash = hashlib.md5(payload).hexdigest()
                    safe_name = sanitize_filename(filename)

                    # Dedup: skip if hash OR name already saved
                    if file_hash in downloaded_hashes:
                        print(f"    SKIP (dup hash): {safe_name}")
                        skipped_dup += 1
                        continue
                    if safe_name in downloaded_names:
                        print(f"    SKIP (dup name): {safe_name}")
                        skipped_dup += 1
                        continue

                    output_path, safe_name = build_unique_output_path(
                        OUTPUT_FOLDER, filename, email_date)

                    # Duplicate name already on disk -> skip (never date-suffix)
                    if output_path is None:
                        print(f"    SKIP (already on disk): {safe_name}")
                        skipped_dup += 1
                        downloaded_names.add(safe_name)
                        continue

                    with open(output_path, "wb") as f:
                        f.write(payload)

                    file_size = len(payload)
                    print(f"    SAVED ({file_size/1024:.1f} KB): "
                          f"{os.path.basename(output_path)}")

                    # ── ZIP: log inventory only, do NOT extract ────────────────
                    zip_contents: list[str] = []
                    if file_ext == ".zip":
                        zip_contents = inventory_zip(output_path)
                        if zip_contents:
                            print(f"      ZIP contains {len(zip_contents)} file(s) "
                                  f"[inventory only, not extracted]")

                    # ── RAR: log inventory only, do NOT extract ────────────────
                    rar_contents: list[str] = []
                    if file_ext == ".rar":
                        rar_contents = inventory_rar(output_path)
                        if rar_contents:
                            print(f"      RAR contains {len(rar_contents)} file(s) "
                                  f"[inventory only, not extracted]")

                    log_entries.append({
                        "msg_id":         msg_id_str,
                        "subject":        subject,
                        "from":           from_addr,
                        "email_date":     email_date,
                        "filename":       filename,
                        "safe_name":      safe_name,
                        "file_hash":      file_hash,
                        "file_size":      file_size,
                        "saved_path":     output_path,
                        "downloaded_at":  datetime.now().isoformat(),
                        "zip_contents":   zip_contents   or None,
                        "rar_contents":   rar_contents   or None,
                    })
                    downloaded_hashes.add(file_hash)
                    downloaded_names.add(safe_name)
                    downloaded += 1

                except Exception as e:
                    print(f"    ERROR saving {filename}: {e}")
                    errors += 1

            # Only mark as processed in skip set if we found attachments
            if email_had_attachment:
                processed_emails.add(msg_id_str)

        except (imaplib.IMAP4.abort, socket.timeout, TimeoutError,
                ConnectionResetError, ConnectionAbortedError, OSError) as e:
            # Connection-level error — socket is desynced; must reconnect
            print(f"\n  CONNECTION ERROR on message {msg_id}: {type(e).__name__}: {e}")
            print(f"  Reconnecting to Gmail ...")
            try:
                mail.logout()
            except Exception:
                pass
            mail = None
            save_log(log_path, log_entries)
            _reconnect_waits = [30, 60, 120]   # longer waits for real network dropouts
            for _ratt in range(1, 4):
                _rwait = _reconnect_waits[_ratt - 1]
                print(f"  Waiting {_rwait}s before reconnect attempt {_ratt}/3 ...")
                time.sleep(_rwait)
                mail = _connect_imap(quiet=True)
                if mail is not None:
                    print(f"  Reconnected (attempt {_ratt}). Continuing from email {i+1}.")
                    socket.setdefaulttimeout(_FETCH_TIMEOUT)
                    break
            if mail is None:
                print("  Could not reconnect after 3 attempts. Stopping early.")
                break
            errors += 1

        except Exception as e:
            print(f"  ERROR processing message {msg_id}: {e}")
            errors += 1

        # Persist log after every email
        save_log(log_path, log_entries)

    # ── Final summary ─────────────────────────────────────────────────────────
    try:
        mail.logout()
    except Exception:
        pass

    save_log(log_path, log_entries)

    print()
    print(SEP)
    print("DOWNLOAD COMPLETE")
    print(SEP)
    print(f"  Scan mode         : {'FULL (all emails)' if _use_full_scan else f'Fast (since {since_date})'}")
    print(f"  Senders scanned   : {' | '.join(TARGETED_SENDERS)}")
    print(f"  Phase 1 scanned   : {len(sorted_ids)} emails (headers only)")
    print(f"  Phase 2 processed : {total_to_process} emails (full fetch)")
    print(f"  Downloaded (new)  : {downloaded} file(s)")
    print(f"  Skipped (dup)     : {skipped_dup} file(s)")
    print(f"  Skipped (ext)     : {skipped_ext} file(s)")
    print(f"  Errors            : {errors}")
    print(f"  Log saved to      : {log_path}")
    print(f"  Output folder     : {OUTPUT_FOLDER}")
    print(SEP)
    print()


if __name__ == "__main__":
    download_attachments()
    sys.exit(0)
