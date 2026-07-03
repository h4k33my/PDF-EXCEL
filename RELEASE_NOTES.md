# Release v1.2.0

Release date: 2026-07-03

Summary
-------
- UBA Bank PDFs now convert correctly: all 7 columns extracted with original header names, footer text excluded.
- Account summary data is now output on a separate **Summary** sheet, keeping it from interfering with the transaction columns.
- Excel files exported by bank portals as HTML (`.xls` extension) now load correctly via a built-in HTML table parser.
- PyQt6 updated to 6.11.0 to resolve DLL load errors on Windows.

Files included in this release
-----------------------------
- `dist_package/GAC-PDF-EXCEL-CONVERTER.exe` (built locally)
- `dist_package/GAC-PDF-EXCEL-CONVERTER.exe.sha256` (updated hash)

Notable changes
---------------

**UBA PDF extraction (OCR path)**
- pdfplumber cannot extract UBA statement tables (no visible grid lines). The app now routes UBA PDFs through an OCR fallback (pytesseract + pdf2image) that correctly reconstructs the 7-column transaction table.
- Column headers use the original PDF wording: `Transaction Date`, `Value Date`, `Cheque Number`, `Transaction Remarks`, `Withdrawal`, `Deposit`, `Balance`.
- Footer logo text ("The Virtual Banker", "United Bank for Africa") is excluded from the transaction data using a gap-threshold filter.

**Account Summary — separate sheet**
- Previously, account metadata (account number, opening/closing balance, statement period, etc.) was prepended to the transaction rows on the same sheet, causing column-count mismatches.
- The summary block is now always placed on a separate **Summary** sheet.
- For UBA-style PDFs where the summary arrives as a single concatenated text blob, a field-name splitter reconstructs proper `[Label, Value]` rows.
- For banks where pdfplumber already extracts a structured key/value table, the raw rows are preserved as-is on the Summary sheet.

**HTML-as-.xls support**
- Bank portals (including some UBA and GTBank download options) export HTML tables with a `.xls` file extension. Loading these previously produced "Unsupported format" or "File is not a zip file" errors.
- The Excel loader now detects HTML content by sniffing the first bytes, and falls back to a pure-Python HTML table parser when the file is not a genuine binary XLS.
- Multiple `<table>` elements in the HTML become separate sheets.

**Dependency update**
- PyQt6 bumped from 6.6.1 to 6.11.0. The 6.6.1 wrapper was mismatched with the 6.11.0 Qt6 DLLs installed on the build machine, causing a `DLL load failed while importing QtCore` error when launching the previous exe.

Upgrade notes
-------------
- Run `pip install -r requirements.txt` to pull in the updated PyQt6 and new OCR dependencies (pytesseract, pdf2image, Pillow).
- Tesseract OCR must be installed separately on the machine for UBA PDF extraction to work. If Tesseract is not found, the app falls back to the coordinate-based extractor.
