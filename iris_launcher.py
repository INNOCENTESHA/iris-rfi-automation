#!/usr/bin/env python3
"""
IRIS RFI Runner  (v5 — June 2026)
══════════════════════════════════
Simple sequential launcher — no watchdog, no TUI, no timeouts.
Each script runs to natural completion; output goes straight to the terminal.

Usage:
  python C:\\AI-Projects\\iris_launcher.py

Scripts run in order:
  1. download_rfi_attachments.py
  2. update_rfi_google_sheets.py
  3. update_physical_progress.py
"""

import os
import sys
import time
import subprocess
from datetime import datetime

# ── Script definitions ────────────────────────────────────────────────────────
SCRIPTS = [
    {
        "name": "Download RFI Attachments",
        "path": r"C:\AI-Projects\download_rfi_attachments.py",
    },
    {
        "name": "RFI → Google Sheets",
        "path": r"C:\AI-Projects\update_rfi_google_sheets.py",
    },
    {
        "name": "Update Physical Progress",
        "path": r"C:\AI-Projects\update_physical_progress.py",
    },
]

SEP  = "─" * 65
DSEP = "═" * 65

# ── Helpers ───────────────────────────────────────────────────────────────────

def banner(msg: str):
    print(f"\n{SEP}")
    print(f"  {msg}")
    print(SEP)


def run_script(idx: int, name: str, path: str) -> tuple[int, float]:
    """Run a single script, stream its output to terminal, return (returncode, elapsed_secs)."""
    if not os.path.exists(path):
        banner(f"✗  MISSING  [{idx}/{len(SCRIPTS)}]  {name}")
        print(f"  Script not found: {path}")
        return -1, 0.0

    banner(f"START  [{idx}/{len(SCRIPTS)}]  {name}  @ {datetime.now().strftime('%H:%M:%S')}")

    t0 = time.time()
    rc = subprocess.run(
        [sys.executable, "-u", path],   # -u: unbuffered stdout in child
    ).returncode
    elapsed = time.time() - t0

    icon   = "✓" if rc == 0 else "✗"
    status = "OK" if rc == 0 else f"FAILED (exit {rc})"
    banner(
        f"{icon}  DONE  [{idx}/{len(SCRIPTS)}]  {name}  "
        f"[{elapsed / 60:.1f} min]  →  {status}"
    )
    return rc, elapsed


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    print(f"\n{DSEP}")
    print(f"  IRIS RFI RUNNER  ·  {datetime.now().strftime('%Y-%m-%d  %H:%M:%S')}")
    print(f"  Scripts: {len(SCRIPTS)}  ·  Running sequentially, no timeouts")
    print(DSEP)

    results: list[tuple[str, int, float]] = []

    for idx, s in enumerate(SCRIPTS, start=1):
        rc, elapsed = run_script(idx, s["name"], s["path"])
        results.append((s["name"], rc, elapsed))

        if rc != 0:
            print(f"\n  ⚠  Script {idx} exited with code {rc} — continuing to next script …")
            if rc == 1 and "Sheets" in s["name"]:
                print(
                    "\n  ─── Google Auth Note ─────────────────────────────────────────\n"
                    "  If the Sheets script printed 'token revoked/expired':\n"
                    "    1. Delete C:\\AI-Projects\\google_token.json\n"
                    "    2. Run:  python C:\\AI-Projects\\update_rfi_google_sheets.py\n"
                    "    3. Complete the browser OAuth flow.\n"
                    "    4. Then re-run this launcher.\n"
                    "  ──────────────────────────────────────────────────────────────"
                )

    # ── Summary ───────────────────────────────────────────────────────────────
    print(f"\n{DSEP}")
    print(f"  FINAL SUMMARY")
    print(SEP)

    total = 0.0
    any_failed = False
    for name, rc, elapsed in results:
        icon   = "✓" if rc == 0 else "✗"
        status = "OK" if rc == 0 else f"FAILED (exit {rc})"
        print(f"  {icon}  {status:<20s}  {elapsed / 60:5.1f} min   {name}")
        total += elapsed
        if rc != 0:
            any_failed = True

    print(SEP)
    total_ok  = sum(1 for _, rc, _ in results if rc == 0)
    total_fail = len(results) - total_ok
    print(
        f"  {'Passed':<22s}  {total_ok}/{len(results)}  scripts\n"
        f"  {'Total elapsed':<22s}  {total / 60:.1f} min\n"
        f"  {'Completed':<22s}  {datetime.now().strftime('%H:%M:%S')}"
    )
    print(DSEP)

    sys.exit(1 if any_failed else 0)


if __name__ == "__main__":
    main()
