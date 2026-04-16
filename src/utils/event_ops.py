"""
Helpers for Inflow/Outflow event mapping on sheet grids.
"""
from __future__ import annotations

import re
from typing import Dict, Iterable, List, Tuple


def clone_grid(data: List[List[object]]) -> List[List[object]]:
    return [list(row) for row in data] if data else []


def normalize_header(value: object) -> str:
    text = str(value or "").strip().lower()
    text = re.sub(r"\s+", " ", text)
    return text


def ensure_header_columns(
    data: List[List[object]], wanted_headers: Iterable[str]
) -> Tuple[List[List[object]], Dict[str, int], List[str]]:
    if not data:
        return [[""]], {}, []

    out = clone_grid(data)
    header = out[0]
    existing: Dict[str, int] = {}
    for idx, cell in enumerate(header):
        key = normalize_header(cell)
        if key and key not in existing:
            existing[key] = idx

    created: List[str] = []
    wanted_map: Dict[str, int] = {}
    for name in wanted_headers:
        raw_name = str(name or "").strip()
        if not raw_name:
            continue
        key = normalize_header(raw_name)
        if not key:
            continue
        if key in existing:
            wanted_map[key] = existing[key]
            continue
        new_idx = len(header)
        header.append(raw_name)
        for r in range(1, len(out)):
            out[r].append("")
        existing[key] = new_idx
        wanted_map[key] = new_idx
        created.append(raw_name)

    return out, wanted_map, created


def apply_event_amount_mapping(
    data: List[List[object]],
    *,
    amount_col_idx: int,
    event_col_idx: int,
    options: Iterable[str],
) -> Tuple[List[List[object]], Dict[str, int], List[str]]:
    """
    For each row, write amount into the column whose header matches selected event.
    Returns mapped_data, stats, created_headers.
    """
    option_names = [str(v or "").strip() for v in options if str(v or "").strip()]
    out, header_map, created = ensure_header_columns(data, option_names)
    if not out:
        return out, {"rows_updated": 0, "rows_skipped": 0, "events_missing": 0}, created

    rows_updated = 0
    rows_skipped = 0
    events_missing = 0
    for row_idx in range(1, len(out)):
        row = out[row_idx]
        if event_col_idx >= len(row) or amount_col_idx >= len(row):
            rows_skipped += 1
            continue
        event_value = str(row[event_col_idx] or "").strip()
        if not event_value:
            rows_skipped += 1
            continue
        key = normalize_header(event_value)
        dest_idx = header_map.get(key)
        if dest_idx is None:
            events_missing += 1
            rows_skipped += 1
            continue
        while len(row) <= dest_idx:
            row.append("")
        row[dest_idx] = row[amount_col_idx]
        rows_updated += 1

    stats = {
        "rows_updated": rows_updated,
        "rows_skipped": rows_skipped,
        "events_missing": events_missing,
    }
    return out, stats, created


def to_number(value: object) -> float:
    text = str(value or "").strip()
    if not text:
        return 0.0
    negative = False
    if text.startswith("(") and text.endswith(")"):
        negative = True
        text = text[1:-1]
    text = text.replace(",", "").replace("$", "").strip()
    try:
        num = float(text)
    except ValueError:
        return 0.0
    return -num if negative else num


def sum_column_values(data: List[List[object]], col_idx: int) -> float:
    if not data or col_idx < 0:
        return 0.0
    total = 0.0
    for row in data[1:]:
        if col_idx < len(row):
            total += to_number(row[col_idx])
    return total


def summarize_totals_for_headers(
    data: List[List[object]], *, amount_col_idx: int, headers: Iterable[str]
) -> Tuple[float, Dict[str, float], float]:
    if not data:
        return 0.0, {}, 0.0
    header_row = data[0]
    header_map: Dict[str, int] = {}
    for idx, value in enumerate(header_row):
        key = normalize_header(value)
        if key and key not in header_map:
            header_map[key] = idx

    amount_total = sum_column_values(data, amount_col_idx)
    per_header: Dict[str, float] = {}
    mapped_total = 0.0
    for name in headers:
        raw = str(name or "").strip()
        if not raw:
            continue
        key = normalize_header(raw)
        col = header_map.get(key)
        total = sum_column_values(data, col) if col is not None else 0.0
        per_header[raw] = total
        mapped_total += total
    return amount_total, per_header, mapped_total
