# Coordinate fallback — status report

**Last updated:** 2026-04-10  

## Scope

- `src/coordinate_fallback.py` — word geometry, header detection, column assignment, canonical rows.
- `src/converter.py` — conservative gate, coordinate replace, period labels (`Balance at Period Start` / `Balance at Period End`), PDF column order (Reference | Description), bank details from page-1 text when tables are thin.
- `test_coordinate_zapphaire.py` — Zapphaire + regression row counts.
- `export_all_test_pdfs.py` — batch Excel under `test_excel_outputs/`.
- `build.py` — PyInstaller build; copies `GAC-PDF-EXCEL-CONVERTER.exe` to `dist_package/`.

## Gate fired (per file)

| PDF | Gate |
|-----|------|
| ZAPPHAIRE EVENTS LIMITED STATEMENT OF ACCOUNT.pdf | Yes |
| MARCH BANK STATEMENT.pdf | No |
| 1a70169f-341a-482a-9554-07aa599aebe3 (1).pdf | No |
| 40d4a207-1ca1-41d7-ab82-ec3b8061230a (1).pdf | No |

## Zapphaire output (current)

- **Bank block:** First-page text lines above the ledger (account number, name, currency, branch, etc.).
- **Header:** Single PDF row (`Book Date`, `Reference`, `Description`, …); no duplicate synthetic `Transaction Date` row.
- **Period rows:** `Balance at Period Start` / `Balance at Period End` with amounts from `extract_period_balances_from_text`.
- **Transactions:** 14 data rows; Reference and Description columns match PDF order after swap.
- **Artifacts:** `_strip_cell_artifacts` removes trailing `Page N of M`, statement dates, and merged `Balance at Period End` fragments from references where applicable.

## Regression (row count after ledger header)

| File | Rows after header |
|------|-------------------|
| MARCH BANK STATEMENT.pdf | 50 |
| 1a70169f-341a-482a-9554-07aa599aebe3 (1).pdf | 77 |
| 40d4a207-1ca1-41d7-ab82-ec3b8061230a (1).pdf | 116 |

## Commands

```text
python test_coordinate_zapphaire.py
python export_all_test_pdfs.py   # writes test_excel_outputs/*.xlsx for each PDF in project root
python build.py                  # dist/ + dist_package/GAC-PDF-EXCEL-CONVERTER.exe
```

**Executable:** `dist/GAC-PDF-EXCEL-CONVERTER.exe` and copy at `dist_package/GAC-PDF-EXCEL-CONVERTER.exe` (rebuilt with coordinate fallback and `main.py` frozen-path fix).

## Notes

- `src/main.py` sets `sys.path` from `src/` in development and from `_MEIPASS` when frozen so `converter` / `coordinate_fallback` resolve in the `.exe`.
- Gate remains conservative (pattern count + primary row shortfall) so Access/GTB-style PDFs stay on the table path.
- **Page-1 text bank block** (`extract_bank_details_from_pdf_text`) is only used to **replace** existing `bank_details` when the coordinate gate fires; otherwise PDFs without a `Book Date` / `Reference` / `Description` header line would pull the entire first page into the bank section.
