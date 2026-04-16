"""
Coordinate-based transaction reconstruction using pdfplumber word geometry.
Mirrors the web app's pdfjs positional logic for borderless / weak-table PDFs.
"""
from __future__ import annotations

import re
from dataclasses import dataclass
from typing import List, Optional, Tuple

import pdfplumber

# --- Header aliases (order: specific before short), same semantics as web pdfParser.ts ---
HEADER_ALIASES: List[Tuple[re.Pattern, str]] = [
    (re.compile(r"^transaction\s+date$", re.I), "Date"),
    (re.compile(r"^book\s+date$", re.I), "Date"),
    (re.compile(r"^posting\s+date$", re.I), "Date"),
    (re.compile(r"^tran(?:s(?:action)?)?\s+date$", re.I), "Date"),
    (re.compile(r"^date$", re.I), "Date"),
    (re.compile(r"^transaction\s+details$", re.I), "Description"),
    (re.compile(r"^transaction\s+description$", re.I), "Description"),
    (re.compile(r"^narration$", re.I), "Description"),
    (re.compile(r"^description$", re.I), "Description"),
    (re.compile(r"^details$", re.I), "Description"),
    (re.compile(r"^particulars$", re.I), "Description"),
    (re.compile(r"^remarks$", re.I), "Description"),
    (re.compile(r"^debit\s+amount$", re.I), "Debit"),
    (re.compile(r"^debit$", re.I), "Debit"),
    (re.compile(r"^withdrawal(?:s)?$", re.I), "Debit"),
    (re.compile(r"^dr$", re.I), "Debit"),
    (re.compile(r"^credit\s+amount$", re.I), "Credit"),
    (re.compile(r"^credit$", re.I), "Credit"),
    (re.compile(r"^deposit(?:s)?$", re.I), "Credit"),
    (re.compile(r"^cr$", re.I), "Credit"),
    (re.compile(r"^current\s+balance$", re.I), "Balance"),
    (re.compile(r"^running\s+balance$", re.I), "Balance"),
    (re.compile(r"^balance$", re.I), "Balance"),
    (re.compile(r"^value\s+date$", re.I), "Value Date"),
    (re.compile(r"^reference(?:\s+no(?:\.|umber)?)?$", re.I), "Reference"),
    (re.compile(r"^ref(?:erence)?\.?$", re.I), "Reference"),
]

NUMERIC_COL_RE = re.compile(r"debit|credit|balance|amount|withdrawal|deposit", re.I)
AMOUNT_RE = re.compile(r"^-?[\d,]+\.\d{2}$")
_PAGE_TAIL_RE = re.compile(r"\s+Page\s+\d+\s+of\s+\d+.*$", re.I | re.DOTALL)
_STMT_DATE_TAIL_RE = re.compile(
    r"\s+\d{1,2}\s+(?:January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{4}\s*$",
    re.I,
)


def _strip_cell_artifacts(s: str) -> str:
    if not s:
        return ""
    t = s.strip()
    t = _PAGE_TAIL_RE.sub("", t)
    t = _STMT_DATE_TAIL_RE.sub("", t)
    # Footer text merged into Reference (Providus): "... Balance at Period End"
    t = re.sub(r"\s*Balance\s+at\s+Period\s+E\s*nd\b.*$", "", t, flags=re.I)
    t = re.sub(r"\s*Balance\s+at\s+Period\s+End\b.*$", "", t, flags=re.I)
    parts = t.split()
    if len(parts) >= 2 and parts[-1] == parts[-2]:
        t = " ".join(parts[:-1])
    return t.strip()


@dataclass
class PWord:
    text: str
    x: float
    width: float
    height: float
    top: float
    page: int


@dataclass
class ColumnDef:
    name: str
    x: float
    right_edge: float
    header_right_x: float
    is_numeric: bool


def _match_header_alias(text: str) -> Optional[str]:
    t = text.strip()
    for pat, canonical in HEADER_ALIASES:
        if pat.match(t):
            return canonical
    return None


def _words_from_pdf(pdf: pdfplumber.PDF) -> List[PWord]:
    out: List[PWord] = []
    for page in pdf.pages:
        pno = page.page_number
        for w in page.extract_words(use_text_flow=False) or []:
            x0 = float(w.get("x0", 0))
            x1 = float(w.get("x1", x0))
            top = float(w.get("top", 0))
            bottom = float(w.get("bottom", top + 10))
            text = (w.get("text") or "").strip()
            if not text:
                continue
            out.append(
                PWord(
                    text=text,
                    x=x0,
                    width=max(x1 - x0, 0.1),
                    height=max(bottom - top, 3),
                    top=top,
                    page=pno,
                )
            )
    return out


def _group_into_lines(words: List[PWord], line_tolerance: float = 3.5) -> List[List[PWord]]:
    raw = sorted(words, key=lambda w: (w.page, w.top, w.x))
    lines: List[List[PWord]] = []
    current: List[PWord] = []
    current_top = -1e9
    current_page = -1

    for w in raw:
        t = w.top
        if w.page != current_page or abs(t - current_top) > line_tolerance:
            if current:
                lines.append(sorted(current, key=lambda z: z.x))
            current = [w]
            current_top = t
            current_page = w.page
        else:
            current.append(w)
    if current:
        lines.append(sorted(current, key=lambda z: z.x))
    return lines


def _try_extract_columns_from_header_lines(header_lines: List[List[PWord]]) -> List[ColumnDef]:
    Tagged = Tuple[PWord, int, bool]  # word, line_idx, used
    items: List[Tagged] = []
    for li, line in enumerate(header_lines):
        for it in line:
            items.append((it, li, False))

    candidates: List[Tuple[str, float, float]] = []
    used_canonical: set = set()

    def mark_used(i: int, j: int) -> None:
        a, la, _ = items[i]
        b, lb, _ = items[j]
        items[i] = (a, la, True)
        items[j] = (b, lb, True)

    i = 0
    while i < len(items):
        item, li, used = items[i]
        if used:
            i += 1
            continue

        matched = False
        # Adjacent on same line
        for j in range(i + 1, len(items)):
            wj, lj, uj = items[j]
            if uj or lj != li:
                break
            gap = wj.x - (item.x + item.width)
            if gap > 20:
                break
            phrase = f"{item.text.strip()} {wj.text.strip()}"
            can = _match_header_alias(phrase)
            if can and can not in used_canonical:
                candidates.append((can, item.x, wj.x + wj.width))
                used_canonical.add(can)
                mark_used(i, j)
                matched = True
                break

        if not matched and li + 1 < len(header_lines):
            best_j = -1
            best_dx = 25.0
            for j in range(len(items)):
                wj, lj, uj = items[j]
                if uj or lj != li + 1:
                    continue
                dx = abs(wj.x - item.x)
                if dx < best_dx:
                    best_dx = dx
                    best_j = j
            if best_j >= 0:
                wj, _, uj = items[best_j]
                phrase = f"{item.text.strip()} {wj.text.strip()}"
                can = _match_header_alias(phrase)
                if can and can not in used_canonical:
                    candidates.append((can, item.x, wj.x + wj.width))
                    used_canonical.add(can)
                    items[i] = (item, li, True)
                    items[best_j] = (wj, items[best_j][1], True)
                    matched = True

        if not matched:
            can = _match_header_alias(item.text.strip())
            if can and can not in used_canonical:
                candidates.append((can, item.x, item.x + item.width))
                used_canonical.add(can)
                items[i] = (item, li, True)

        i += 1

    if len(candidates) < 3:
        return []

    candidates.sort(key=lambda c: c[1])
    cols: List[ColumnDef] = []
    for idx, (name, x, hrx) in enumerate(candidates):
        right = candidates[idx + 1][1] if idx + 1 < len(candidates) else float("inf")
        cols.append(
            ColumnDef(
                name=name,
                x=x,
                right_edge=right,
                header_right_x=hrx,
                is_numeric=bool(NUMERIC_COL_RE.search(name)),
            )
        )
    return cols


def _detect_header(lines: List[List[PWord]]) -> Optional[Tuple[int, List[ColumnDef]]]:
    max_scan = min(len(lines), 100)
    for start_i in range(max_scan):
        best: Optional[Tuple[int, List[ColumnDef]]] = None
        for span in range(1, 4):
            if start_i + span > len(lines):
                break
            header_lines = lines[start_i : start_i + span]
            cols = _try_extract_columns_from_header_lines(header_lines)
            if len(cols) >= 3:
                end_idx = start_i + span - 1
                if best is None or len(cols) > len(best[1]):
                    best = (end_idx, cols)
        if best is not None:
            return best
    return None


def _assign_line_to_cells(line: List[PWord], columns: List[ColumnDef]) -> List[str]:
    cells = [""] * len(columns)
    for w in line:
        text = w.text.strip()
        if not text:
            continue
        looks_amount = bool(AMOUNT_RE.match(text))
        best_col = -1
        best_dist = float("inf")

        if looks_amount:
            item_right = w.x + w.width
            for i, col in enumerate(columns):
                if not col.is_numeric:
                    continue
                dist = abs(item_right - col.header_right_x)
                if dist < best_dist:
                    best_dist = dist
                    best_col = i
            if best_dist > 40:
                best_col = -1
                best_dist = float("inf")

        if best_col < 0:
            item_center = w.x + w.width / 2
            for i, col in enumerate(columns):
                col_width = columns[i + 1].x - col.x if i + 1 < len(columns) else 200
                col_center = col.x + col_width / 2
                dist = abs(item_center - col_center)
                if w.x >= col.x - 15 and (i == len(columns) - 1 or w.x < columns[i + 1].x + 10):
                    if dist < best_dist:
                        best_dist = dist
                        best_col = i

        if best_col < 0:
            min_d = float("inf")
            for i, col in enumerate(columns):
                col_end = columns[i + 1].x if i + 1 < len(columns) else col.x + 200
                if col.x - 20 <= w.x <= col_end + 20:
                    d = abs(w.x - col.x)
                    if d < min_d:
                        min_d = d
                        best_col = i

        if best_col >= 0:
            cur = cells[best_col]
            cells[best_col] = f"{cur} {text}".strip() if cur else text
    return cells


def _is_data_row(cells: List[str], columns: List[ColumnDef]) -> bool:
    all_text = " ".join(cells)
    has_date_mdy = bool(
        re.search(r"\b\d{1,2}\s+(?:JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)\s+\d{2,4}\b", all_text, re.I)
    )
    has_date_slash = bool(re.search(r"\b\d{2}/\d{2}/\d{4}\b", all_text))
    has_date_iso = bool(re.search(r"\b\d{4}-\d{2}-\d{2}\b", all_text))
    has_date = has_date_mdy or has_date_slash or has_date_iso
    has_amount = bool(re.search(r"[\d,]+\.\d{2}", all_text))
    has_ref = bool(re.search(r"\b[A-Z0-9]{8,}\b", all_text))
    return has_date or (has_amount and has_ref)


SKIP_LINE_PATTERNS = [
    re.compile(r"Balance\s+at\s+Period", re.I),
    re.compile(r"Opening\s+Balance", re.I),
    re.compile(r"Account\s+Statement", re.I),
    re.compile(r"Account\s+Number", re.I),
    re.compile(r"Account\s+Name", re.I),
    re.compile(r"Currency\s*:", re.I),
    re.compile(r"Branch\s+:", re.I),
    re.compile(r"Short\s+Name", re.I),
    re.compile(r"Account\s+Type", re.I),
    re.compile(r"^\s*Page\s+\d+", re.I),
    re.compile(r"^\d{1,2}:\d{2}:\d{2}$"),
    re.compile(
        r"^\d{1,2}\s+(?:January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{4}$",
        re.I,
    ),
    re.compile(r"Transaction\s+Description", re.I),
    re.compile(r"Number\s+of\s+(?:Debit|Credit)\s+Transaction", re.I),
    re.compile(r"Total\s+(?:Debit|Credit)\s+Amount", re.I),
    re.compile(r"Period\s+Opening\s+Balance", re.I),
    re.compile(r"Period\s+Closing\s+Balance", re.I),
    re.compile(r"Statement\s+of\s+Account", re.I),
    re.compile(r"IBAN\s*:", re.I),
    re.compile(r"Account\s+Nickname", re.I),
    re.compile(r"Transaction\s+Type", re.I),
    re.compile(r"DR/\s*CR", re.I),
    re.compile(r"^\s*NGN\s*$", re.I),
    re.compile(r"P\.O\.Box", re.I),
    re.compile(r"PROVIDUS\s+BANK", re.I),
    re.compile(r"IKEJA\s+BRANCH", re.I),
]


def _line_to_string(line: List[PWord]) -> str:
    if not line:
        return ""
    parts: List[str] = []
    prev_right = line[0].x
    for w in line:
        gap = w.x - prev_right
        if gap > 5:
            parts.append(" " * max(1, min(int(round(gap / 7)), 20)))
        parts.append(w.text)
        prev_right = w.x + w.width
    return "".join(parts)


def _parse_rows_after_header(
    lines: List[List[PWord]], header_line_idx: int, columns: List[ColumnDef]
) -> List[List[str]]:
    rows: List[List[str]] = []
    current: Optional[List[str]] = None

    desc_col = next((i for i, c in enumerate(columns) if re.search(r"description|narration|details|particulars", c.name, re.I)), -1)

    def flush() -> None:
        nonlocal current
        if current is not None:
            rows.append(current)
            current = None

    for i in range(header_line_idx + 1, len(lines)):
        line = lines[i]
        line_str = _line_to_string(line)

        if any(p.search(line_str) for p in SKIP_LINE_PATTERNS):
            bal_m = re.search(r"([\d,]+\.\d{2})\s*$", line_str)
            if bal_m and re.search(r"Balance\s+at\s+Period", line_str, re.I):
                flush()
                cells = [""] * len(columns)
                bi = next((j for j, c in enumerate(columns) if re.search(r"balance", c.name, re.I)), -1)
                if bi >= 0:
                    cells[bi] = bal_m.group(1)
                if desc_col >= 0:
                    m2 = re.search(r"Balance\s+at\s+Period\s+\w+", line_str, re.I)
                    cells[desc_col] = m2.group(0) if m2 else "Balance"
                rows.append(cells)
            else:
                flush()
            continue

        if not line:
            continue

        cells = _assign_line_to_cells(line, columns)
        filled = sum(1 for c in cells if c.strip())

        if filled == 0:
            continue

        if _is_data_row(cells, columns):
            flush()
            current = cells
        elif current is not None and filled <= 2:
            for ci, val in enumerate(cells):
                v = val.strip()
                if not v:
                    continue
                if current[ci]:
                    current[ci] = f"{current[ci]} {v}".strip()
                else:
                    current[ci] = v
        else:
            flush()

    flush()
    return rows


def _cells_to_canonical(cells: List[str], columns: List[ColumnDef]) -> List[str]:
    """Map named columns to 7-field output matching converter.py."""
    by_name = {columns[i].name: cells[i].strip() if i < len(cells) else "" for i in range(len(columns))}

    date_v = by_name.get("Date", "")
    desc_v = by_name.get("Description", "")
    ref_v = by_name.get("Reference", "")
    vd_v = by_name.get("Value Date", "")
    deb_v = by_name.get("Debit", "")
    cred_v = by_name.get("Credit", "")
    bal_v = by_name.get("Balance", "")

    # Single Amount column -> heuristic
    if "Amount" in by_name and not deb_v and not cred_v:
        amt = by_name["Amount"]
        if amt:
            deb_v = amt

    return [date_v, desc_v, ref_v, vd_v, deb_v, cred_v, bal_v]


def reconstruct_transactions_coordinate(pdf_path: str) -> List[List[str]]:
    """
    Extract transactions using word positions. Returns canonical rows:
    [Transaction Date, Details, Reference, Value Date, Debit, Credit, Balance]
    """
    with pdfplumber.open(pdf_path) as pdf:
        words = _words_from_pdf(pdf)
        if not words:
            return []

        lines = _group_into_lines(words, line_tolerance=3.5)
        hdr = _detect_header(lines)
        if not hdr:
            return []

        header_idx, columns = hdr
        raw_rows = _parse_rows_after_header(lines, header_idx, columns)
        out: List[List[str]] = []
        for rc in raw_rows:
            if len(rc) < len(columns):
                rc = rc + [""] * (len(columns) - len(rc))
            canon = _cells_to_canonical(rc[: len(columns)], columns)
            canon = [_strip_cell_artifacts(c) for c in canon]
            if any(c.strip() for c in canon):
                out.append(canon)
    return out


def count_probable_transaction_lines_in_text(pdf_path: str) -> int:
    """Heuristic: lines matching DD MON YY + ref + amounts (for gating)."""
    pat = re.compile(
        r"\d{2}\s+[A-Z]{3}\s+\d{2}\s+[A-Z0-9]{8,}.+\d{2}\s+[A-Z]{3}\s+\d{2}\s+[\d,]+\.\d{2}\s+[\d,]+\.\d{2}",
        re.I,
    )
    n = 0
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            for ln in text.split("\n"):
                if pat.search(ln):
                    n += 1
    return n


def should_use_coordinate_fallback(
    pdf_path: str,
    primary_transaction_count: int,
) -> bool:
    """
    Conservative gate: use coordinate path when table extraction likely under-filled
    but page text shows many structured transaction lines.
    """
    try:
        expected = count_probable_transaction_lines_in_text(pdf_path)
    except Exception:
        return False

    if expected >= 12 and primary_transaction_count < expected - 3:
        return True

    with pdfplumber.open(pdf_path) as pdf:
        first = pdf.pages[0] if pdf.pages else None
        if not first:
            return False
        text0 = (first.extract_text() or "").upper()
        words0 = " ".join(w.get("text", "") for w in (first.extract_words() or [])).upper()
        has_book_date = "BOOK DATE" in text0 or "BOOK DATE" in words0
        has_balance_hdr = bool(re.search(r"\bBALANCE\b", words0)) and "DEBIT" in words0

    if has_book_date and has_balance_hdr and primary_transaction_count < 12:
        return True

    return False
