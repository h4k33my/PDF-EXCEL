# Release v1.1.4

Release date: 2026-05-15

Summary
-------
- Add support for legacy Excel formats: `.xls` and `.xlsm` in addition to `.xlsx`.
- UI: allow selecting `.xls` and `.xlsm` files in the Excel browse dialog.
- Add `xlrd` dependency to read `.xls` files.
- Update auto-update asset verification: refreshed `GAC-PDF-EXCEL-CONVERTER.exe.sha256` to match newly built exe.
- Misc: various parsing and UI improvements (table preservation for invoice-style PDFs, full-text sheet fallback, improved header selection and preview state preservation).

Files included in this release
-----------------------------
- `dist_package/GAC-PDF-EXCEL-CONVERTER.exe` (built locally)
- `dist_package/GAC-PDF-EXCEL-CONVERTER.exe.sha256` (updated hash)

Notable changes
---------------
- Excel loading: the app now reads `.xls` files using `xlrd` and `.xlsx/.xlsm` via `openpyxl`. The loader returns the same sheet/dict structure used by the converter.
- Invoice conversion: invoice-style 3-column tables are preserved as sheets and a `Full Text` sheet is added to retain non-table details.
- UI and UX: header selection behavior and preview refreshes were improved to preserve view state and avoid losing the user's current position.
- Updater: asset lookup and SHA verification were hardened to avoid false negatives during auto-update checks.

Upgrade notes
-------------
- After updating, install Python dependencies listed in `requirements.txt` (new `xlrd==2.0.1` entry).

How to publish (manual steps)
----------------------------
1. Create a GitHub release for tag `v1.1.4` (or use the web UI). Use this `RELEASE_NOTES.md` content as the release description.
2. Upload the built executable `dist_package/GAC-PDF-EXCEL-CONVERTER.exe` and `dist_package/GAC-PDF-EXCEL-CONVERTER.exe.sha256` as release assets.
3. Mark the release as published.

If you want me to publish the release automatically, provide a GitHub personal access token with `repo` scope and I can call the GitHub releases API to create the release and upload the assets.
