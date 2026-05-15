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
    row_event_keys: Dict[int, str] | None = None,
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
    destination_indices = list(header_map.values())
    for row_idx in range(1, len(out)):
        row = out[row_idx]
        if event_col_idx >= len(row) or amount_col_idx >= len(row):
            rows_skipped += 1
            continue
        # Keep exactly one mapped destination per row among mapping headers.
        for dest_col in destination_indices:
            while len(row) <= dest_col:
                row.append("")
            row[dest_col] = ""

        event_value = ""
        if row_event_keys is not None:
            event_value = str(row_event_keys.get(row_idx, "") or "").strip()
        text_value = str(row[event_col_idx] or "").strip()
        if not event_value:
            text_key = normalize_header(text_value) if text_value else ""
            if text_key in header_map:
                event_value = text_value
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


def detect_description_column(header_row: List[object]) -> int | None:
    """Return index of the first column whose header matches description-like keywords."""
    keywords = (
        "transaction details",
        "description",
        "narration",
        "particulars",
        "remarks",
        "details",
    )
    if not header_row:
        return None
    for idx, cell in enumerate(header_row):
        text = str(cell or "").strip().lower()
        if not text:
            continue
        for k in keywords:
            if k in text:
                return idx
    return None


def _alias_candidates(option_text: str, aliases: Dict[str, List[str]]) -> List[str]:
    """Build the lowercase substring candidates for matching one event option."""
    candidates: List[str] = []
    for part in str(option_text).split("|"):
        token = part.strip().lower()
        if len(token) >= 2:
            candidates.append(token)
    for alias in aliases.get(option_text, []) or []:
        token = str(alias).strip().lower()
        if len(token) >= 2:
            candidates.append(token)
    # Deduplicate while preserving order
    seen = set()
    unique = []
    for c in candidates:
        if c not in seen:
            seen.add(c)
            unique.append(c)
    return unique


def auto_assign_events_by_description(
    data: List[List[object]],
    *,
    event_col_idx: int,
    description_col_idx: int,
    options: Iterable[str],
    aliases: Dict[str, List[str]] | None = None,
    prefilled_rows: Iterable[int] | None = None,
) -> Dict[int, str]:
    """
    Scan rows; for each row whose event cell is empty (and not prefilled),
    find the first option whose name parts (split on '|') or aliases appear
    as case-insensitive substrings of the description. Returns
    {row_idx: matched_option}. Does NOT modify `data`.
    """
    if not data or event_col_idx < 0 or description_col_idx < 0:
        return {}
    aliases = aliases or {}
    prefilled = set(prefilled_rows or [])
    option_list = [opt for opt in options if str(opt or "").strip()]
    if not option_list:
        return {}

    candidates_per_option = [(opt, _alias_candidates(opt, aliases)) for opt in option_list]

    matches: Dict[int, str] = {}
    for row_idx in range(1, len(data)):
        if row_idx in prefilled:
            continue
        row = data[row_idx]
        # Skip rows whose event cell already has content.
        existing = ""
        if event_col_idx < len(row):
            existing = str(row[event_col_idx] or "").strip()
        if existing:
            continue
        if description_col_idx >= len(row):
            continue
        description = str(row[description_col_idx] or "").strip().lower()
        if not description:
            continue
        for option, tokens in candidates_per_option:
            if any(tok in description for tok in tokens):
                matches[row_idx] = option
                break
    return matches


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
