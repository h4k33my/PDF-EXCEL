"""Row/column helpers for sheet tools (primary column filter)."""
from __future__ import annotations

from typing import List

from converter import to_float


def validate_numeric_primary_column(data: List[list], col_idx: int) -> bool:
    """
    True if every non-empty cell in the column (data rows only) parses as a number.
    Empty cells are allowed (e.g. sparse columns).
    """
    if col_idx < 0:
        return False
    for row in data[1:]:
        if col_idx >= len(row):
            continue
        cell = row[col_idx]
        s = str(cell).strip() if cell is not None else ""
        if not s:
            continue
        if to_float(cell) is None:
            return False
    return True


def filter_rows_by_positive_primary(data: List[list], col_idx: int) -> List[list]:
    """Keep header and rows where primary column parses to float > 0."""
    if not data or col_idx < 0:
        return data
    header = data[0]
    out = [list(header)]
    for row in data[1:]:
        if col_idx >= len(row):
            continue
        cell = row[col_idx]
        s = str(cell).strip() if cell is not None else ""
        if not s:
            continue
        v = to_float(cell)
        if v is not None and v > 0:
            out.append(list(row))
    return out
