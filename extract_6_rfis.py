"""
Quick script to extract text from 6 unresolved IRIS-Q-RFI PDFs.
Saves output to C:\AI-Projects\rfi_6_extract.txt for review.
"""
import sys
import pathlib

pdf_dir = pathlib.Path(r"D:\HIGHWAY ENGINEERING CAREER\RFI-ZIPs")
files = [
    "IRIS-Q-RFI-P-25-0001.pdf",
    "IRIS-Q-RFI-P-25-0004.pdf",
    "IRIS-Q-RFI-U-25-0001.pdf",
    "IRIS-Q-RFI-U-25-0002.pdf",
    "IRIS-Q-RFI-U-25-0003.pdf",
    "IRIS-Q-RFI-U-25-0005.pdf",
]

output_lines = []

try:
    import pdfplumber
except ImportError:
    sys.exit("pdfplumber not installed")

for fname in files:
    fpath = pdf_dir / fname
    output_lines.append(f"\n{'='*70}")
    output_lines.append(f"FILE: {fname}")
    output_lines.append(f"{'='*70}")
    try:
        with pdfplumber.open(fpath) as pdf:
            for i, page in enumerate(pdf.pages, 1):
                text = page.extract_text() or ""
                output_lines.append(f"--- Page {i} ---")
                output_lines.append(text[:4000] if text else "(no text extracted)")
    except Exception as e:
        output_lines.append(f"ERROR: {e}")

out_path = pathlib.Path(r"C:\AI-Projects\rfi_6_extract.txt")
out_path.write_text("\n".join(output_lines), encoding="utf-8")
print(f"Saved to {out_path}")
