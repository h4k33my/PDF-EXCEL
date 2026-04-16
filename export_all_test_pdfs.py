#!/usr/bin/env python3
"""
Convert every PDF in the project root to Excel under test_excel_outputs/.
Run from project root: python export_all_test_pdfs.py
"""
from __future__ import annotations

import os
import re
import sys

ROOT = os.path.dirname(os.path.abspath(__file__))
OUT_DIR = os.path.join(ROOT, "test_excel_outputs")
sys.path.insert(0, os.path.join(ROOT, "src"))

from converter import extract_all_tables_from_pdf, export_to_excel  # noqa: E402


def _safe_xlsx_name(pdf_name: str) -> str:
    base = os.path.splitext(pdf_name)[0]
    base = re.sub(r'[^\w\-]+', "_", base, flags=re.UNICODE).strip("_")
    return (base or "output") + ".xlsx"


def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    pdfs = sorted(
        f
        for f in os.listdir(ROOT)
        if f.lower().endswith(".pdf") and os.path.isfile(os.path.join(ROOT, f))
    )
    if not pdfs:
        print("No PDF files found in project root.")
        return 1
    ok = 0
    for name in pdfs:
        pdf_path = os.path.join(ROOT, name)
        out_path = os.path.join(OUT_DIR, _safe_xlsx_name(name))
        try:
            sheets = extract_all_tables_from_pdf(pdf_path)
            export_to_excel(sheets, out_path)
            rows = len(sheets[0]["data"]) if sheets and sheets[0].get("data") else 0
            print(f"OK  {name} -> {os.path.basename(out_path)} ({rows} rows)")
            ok += 1
        except Exception as e:
            print(f"ERR {name}: {e}")
    print(f"\nExported {ok}/{len(pdfs)} files to {OUT_DIR}")
    return 0 if ok == len(pdfs) else 2


if __name__ == "__main__":
    sys.exit(main())
