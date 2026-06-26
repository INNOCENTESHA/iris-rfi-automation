"""
delete_extracted_files.py
=========================
Permanently deletes all files that were created by the old ZIP/RAR
extraction code — files ending in  _rar.<ext>  or  _unzipped.<ext>
(and their counter variants  _rar_N.<ext>  /  _unzipped_N.<ext>).

These are flat files dumped into D:\\HIGHWAY ENGINEERING CAREER\\RFI-ZIPs
by previous versions of the downloader. The archive files themselves
(.zip, .rar) are NOT touched.

Run once:
    python C:\\AI-Projects\\delete_extracted_files.py
"""

import os
import re
import sys

RFI_ZIPS_FOLDER = r"D:\HIGHWAY ENGINEERING CAREER\RFI-ZIPs"

# Pattern matches: *_rar.pdf  *_rar_1.pdf  *_unzipped.pdf  *_unzipped_3.pdf
_EXTRACTED_PATTERN = re.compile(
    r"(_rar(\.\w+|_\d+\.\w+)|_unzipped(\.\w+|_\d+\.\w+))$",
    re.IGNORECASE,
)


def find_extracted_files(folder: str) -> list[str]:
    """Walk folder (non-recursively) and return paths matching the pattern."""
    hits = []
    if not os.path.isdir(folder):
        print(f"ERROR: Folder not found: {folder}")
        return hits
    for entry in os.scandir(folder):
        if entry.is_file():
            name_no_ext, ext = os.path.splitext(entry.name)
            # Check if the name (without final ext) ends with _rar or _unzipped suffix
            if _EXTRACTED_PATTERN.search(name_no_ext + ext):
                hits.append(entry.path)
    return hits


def main():
    print()
    print("=" * 60)
    print("  IRIS — Delete extracted files (_rar, _unzipped)")
    print(f"  Folder: {RFI_ZIPS_FOLDER}")
    print("=" * 60)
    print()

    targets = find_extracted_files(RFI_ZIPS_FOLDER)

    if not targets:
        print("  No _rar or _unzipped files found. Nothing to delete.")
        print()
        return

    total_size = 0
    print(f"  Found {len(targets)} file(s) to delete:\n")
    for path in sorted(targets):
        size = 0
        try:
            size = os.path.getsize(path)
        except OSError:
            pass
        total_size += size
        print(f"    {os.path.basename(path)}  ({size/1024:.1f} KB)")

    print()
    print(f"  Total size to free: {total_size / (1024*1024):.2f} MB")
    print()

    deleted = 0
    failed  = 0
    for path in targets:
        try:
            os.remove(path)
            deleted += 1
        except Exception as e:
            print(f"  ERROR deleting {path}: {e}")
            failed += 1

    print(f"  Deleted : {deleted} file(s)")
    if failed:
        print(f"  Failed  : {failed} file(s)")
    print()
    print("  Done.")
    print()


if __name__ == "__main__":
    main()
