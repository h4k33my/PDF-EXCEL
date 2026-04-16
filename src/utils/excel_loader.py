"""
Load .xlsx workbooks into the same sheet dict structure used by the PDF converter.
"""
from __future__ import annotations

from typing import Any, List


def _cell(v: Any) -> str:
    if v is None:
        return ""
    return str(v)


def load_xlsx_to_sheets_data(path: str) -> List[dict]:
    """
    Read all worksheets from an .xlsx file.
    Returns [{'name': str, 'data': list[list[str]], 'is_table': True}, ...]
    """
    from openpyxl import load_workbook

    wb = load_workbook(path, read_only=True, data_only=True)
    out: List[dict] = []
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
