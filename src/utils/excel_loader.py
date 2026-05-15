"""
Load Excel workbooks into the same sheet dict structure used by the PDF converter.

Supports: .xlsx, .xlsm (via openpyxl) and legacy .xls (via xlrd).
"""
from __future__ import annotations

from typing import Any, List
import os


def _cell(v: Any) -> str:
    if v is None:
        return ""
    return str(v)


def load_xlsx_to_sheets_data(path: str) -> List[dict]:
    """
    Read all worksheets from an Excel file (.xlsx, .xlsm, .xls).
    Returns [{'name': str, 'data': list[list[str]], 'is_table': True}, ...]
    """
    ext = os.path.splitext(path)[1].lower()
    out: List[dict] = []
    if ext in ('.xlsx', '.xlsm'):
        from openpyxl import load_workbook

        wb = load_workbook(path, read_only=True, data_only=True)
        try:
            for ws in wb.worksheets:
                rows = []
                for row in ws.iter_rows(values_only=True):
                    rows.append([_cell(c) for c in row])
                # Trim trailing completely empty rows
                while rows and not any(str(c).strip() for c in rows[-1]):
                    rows.pop()
                if not rows:
                    rows = [[""]]
                # Normalize row widths
                max_cols = max(len(r) for r in rows) if rows else 1
                rows = [r + [""] * (max_cols - len(r)) for r in rows]
                name = ws.title[:31] or "Sheet"
                out.append({"name": name, "data": rows, "is_table": True})
        finally:
            wb.close()
        return out

    if ext == '.xls':
        # xlrd is used for legacy .xls files
        try:
            import xlrd  # type: ignore
        except Exception as e:
            raise ImportError("xlrd is required to read .xls files; please install it (pip install xlrd)") from e

        wb = xlrd.open_workbook(path, formatting_info=False)
        for sheet in wb.sheets():
            rows = []
            for ri in range(sheet.nrows):
                row = sheet.row_values(ri)
                rows.append([_cell(c) for c in row])
            # Trim trailing completely empty rows
            while rows and not any(str(c).strip() for c in rows[-1]):
                rows.pop()
            if not rows:
                rows = [[""]]
            max_cols = max(len(r) for r in rows) if rows else 1
            rows = [r + [""] * (max_cols - len(r)) for r in rows]
            name = (sheet.name[:31] or "Sheet")
            out.append({"name": name, "data": rows, "is_table": True})
        return out

    raise ValueError(f"Unsupported Excel extension: {ext}")
