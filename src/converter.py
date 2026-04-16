"""
PDF-to-Excel Converter Module
Simple, reliable extraction for MARCH bank statements.
"""
import pdfplumber
from coordinate_fallback import reconstruct_transactions_coordinate, should_use_coordinate_fallback
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
import re

PAGE_MARKER_RE = re.compile(r"\b(?:page|pg|p|pag)\.?\s*\d+\b", re.I)
INLINE_PAGE_ARTIFACT_RE = re.compile(r"\b(?:pag|page)\b\s*\d*\b", re.I)
NUMERIC_RE = re.compile(r"^-?\d{1,3}(?:,\d{3})*(?:\.\d{1,2})?$|^-?\d+(?:\.\d{1,2})?$")
DATE_RE = re.compile(
    r"^(\d{1,2}[/-]\d{1,2}[/-]\d{2,4}|\d{1,2}\s+(?:JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)\s+\d{2,4}|\d{1,2}-(?:JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)-\d{2,4})$",
    re.I
)


def clean_cell(cell):
    """Remove newlines, page markers, and normalize whitespace."""
    if cell is None:
        return ""
    s = str(cell).replace('\n', ' ').strip()
    s = PAGE_MARKER_RE.sub('', s).strip()
    s = INLINE_PAGE_ARTIFACT_RE.sub('', s).strip()
    s = re.sub(r'\s+', ' ', s).strip()
    return s


def is_date_like(value):
    """Return True if value resembles a statement date."""
    if value is None:
        return False
    return bool(DATE_RE.match(str(value).strip()))


def to_float(value):
    """Parse numeric values with thousand separators."""
    if value is None:
        return None
    s = str(value).strip().replace(',', '')
    if not s:
        return None
    try:
        return float(s)
    except ValueError:
        return None


def first_numeric_token(value):
    """Extract first number token from mixed artifacts like '74.26 995,412.22'."""
    if value is None:
        return ""
    s = str(value).strip()
    if not s:
        return ""
    # Keep full clean number values as-is.
    if NUMERIC_RE.match(s):
        return s
    # Split tokens and return first numeric-looking token.
    for token in re.split(r"\s+", s):
        token = token.strip()
        if NUMERIC_RE.match(token):
            return token
    return s


def is_page_marker_row(row):
    """Return True if a row contains only page-number markers."""
    if not row:
        return False
    non_empty = [cell.strip() for cell in row if cell and cell.strip()]
    if not non_empty:
        return False
    return all(PAGE_MARKER_RE.fullmatch(cell) for cell in non_empty)


def drop_empty_columns(table_rows):
    """Remove columns that are blank in every row."""
    if not table_rows:
        return []
    max_cols = max(len(row) for row in table_rows)
    normalized = [row + [""] * (max_cols - len(row)) for row in table_rows]
    keep_columns = [col_idx for col_idx in range(max_cols)
                    if any(normalized[row_idx][col_idx].strip() for row_idx in range(len(normalized)))]
    if not keep_columns:
        return []
    return [[row[col_idx] for col_idx in keep_columns] for row in normalized]


def is_total_or_closing_row(row):
    """Check if a row is a total/closing summary row, not a transaction."""
    if not row:
        return False
    row_text = ' '.join(str(c).lower() for c in row if c)
    summary_keywords = ['total', 'closing balance', 'balance b/f', 'summary', 'grand total']
    return any(keyword in row_text for keyword in summary_keywords)


def extract_bank_details_from_raw_table(raw_table):
    """Extract bank details from the beginning of a raw table before normalization."""
    if not raw_table:
        return []
    
    bank_details = []
    for row_idx, row in enumerate(raw_table):
        if not row:
            continue
        row_text = ' '.join(str(c).lower() for c in row if c)
        
        # Stop when we hit transaction headers or a row with clear transaction pattern
        if any(keyword in row_text for keyword in ['transaction date', 'date', 'transaction details']):
            # Check if this row has the standard transaction header structure
            if 'date' in row_text and 'details' in row_text:
                break
        
        # Include rows with bank/account information
        has_bank_info = any(keyword in row_text for keyword in [
            'account', 'summary', 'opening balance', 'closing balance', 'currency', 
            'total debit', 'total credit', 'withdrawals', 'lodgements', 'branch'
        ])
        
        if has_bank_info or row_idx < 15:  # Include first 15 rows to be safe
            cleaned = [clean_cell(c) for c in row]
            # Only add if has some content
            if any(c.strip() for c in cleaned):
                bank_details.append(cleaned)
        
        # Stop after we've collected enough header rows and see actual transaction content
        if row_idx > 20 and any(keyword in row_text for keyword in ['date', 'balance', 'amount']):
            break
    
    # Filter out rows that are only continuation of previous or just whitespace
    filtered = []
    for row in bank_details:
        if not any(c.strip() for c in row):
            continue
        filtered.append(row)
    
    return filtered


def is_transaction_table(table_rows):
    """Detect whether a table is a transaction ledger table across bank formats."""
    if not table_rows:
        return False
    header_variants = {
        'transaction date', 'book date', 'date', 'value date', 'details', 'description', 'narration',
        'particulars', 'reference', 'ref', 'debit', 'withdrawal', 'dr', 'credit', 'deposit', 'cr', 'balance'
    }

    # Strong signal: header-like row in first few rows with multiple transaction keywords.
    for row in table_rows[:3]:
        row_text = ' '.join(str(cell).lower() for cell in row if cell).strip()
        hits = sum(1 for k in header_variants if k in row_text)
        if hits >= 3:
            return True

    # Fallback: rows that look like transaction data (date + numeric amount/balance pattern).
    data_like_rows = 0
    for row in table_rows[:10]:
        if not row:
            continue
        first = str(row[0]).strip() if row else ""
        if not is_date_like(first):
            continue
        numeric_cells = sum(1 for cell in row if to_float(cell) is not None)
        if numeric_cells >= 2 and len(row) >= 5:
            data_like_rows += 1

    return data_like_rows >= 2


def looks_like_transaction_header_row(row):
    """Return True when a row looks like a transaction header, not data."""
    if not row:
        return False
    cells = [clean_cell(cell).lower() for cell in row if clean_cell(cell)]
    if not cells:
        return False
    text = ' '.join(cells)
    header_keywords = [
        'transaction', 'book date', 'date', 'details', 'description', 'narration', 'particulars',
        'reference', 'ref', 'debit', 'withdrawal', 'dr', 'credit', 'deposit', 'cr', 'balance', 'value date'
    ]
    keyword_hits = sum(1 for keyword in header_keywords if keyword in text)
    has_header_keyword = keyword_hits >= 2
    first_cell = cells[0] if cells else ''
    starts_like_date = is_date_like(first_cell)
    # Avoid classifying closing-balance narrative rows as headers.
    if 'closing balance' in text and 'transaction' not in text:
        return False
    return has_header_keyword and not starts_like_date


def detect_transaction_header_index(table_rows):
    """Find the header row index in a transaction table; return None when absent."""
    for idx, row in enumerate(table_rows[:5]):
        if looks_like_transaction_header_row(row):
            row_text = ' '.join(str(c).lower() for c in row if c)
            keyword_hits = sum(
                1 for k in ['date', 'details', 'description', 'debit', 'credit', 'balance', 'reference', 'value date']
                if k in row_text
            )
            if keyword_hits >= 3:
                return idx
    return None


def map_transaction_columns(header_row):
    """Map normalized transaction header columns to canonical fields."""
    mapped = {
        'date': None,
        'details': None,
        'reference': None,
        'value_date': None,
        'debit': None,
        'credit': None,
        'balance': None,
    }
    for idx, raw in enumerate(header_row):
        text = clean_cell(raw).lower()
        compact = re.sub(r'[^a-z]', '', text)
        if not text:
            continue
        if mapped['date'] is None and any(k in text for k in ['transaction date', 'book date', 'date']):
            mapped['date'] = idx
        elif mapped['details'] is None and any(k in text for k in ['details', 'description', 'narration', 'particular']):
            mapped['details'] = idx
        elif mapped['reference'] is None and any(k in text for k in ['reference', 'ref']):
            mapped['reference'] = idx
        elif mapped['value_date'] is None and 'value date' in text:
            mapped['value_date'] = idx
        elif mapped['debit'] is None and (
            any(k in text for k in ['debit', 'withdrawal']) or compact in {'dr', 'debitamount'}
        ):
            mapped['debit'] = idx
        elif mapped['credit'] is None and (
            any(k in text for k in ['credit', 'deposit', 'lodgement']) or compact in {'cr', 'creditamount'}
        ):
            mapped['credit'] = idx
        elif mapped['balance'] is None and 'balance' in text:
            mapped['balance'] = idx
    return mapped


def infer_transaction_columns_from_rows(table_rows):
    """Infer transaction column positions when a header row is missing."""
    if not table_rows:
        return {}
    sample_rows = [row for row in table_rows[:8] if row]
    if not sample_rows:
        return {}

    # Common compact transaction layout:
    # [date, details, reference, value_date, debit, credit, balance]
    for row in sample_rows:
        if len(row) < 7:
            continue
        first = clean_cell(row[0])
        val_date = clean_cell(row[3])
        debit = first_numeric_token(clean_cell(row[4]))
        credit = first_numeric_token(clean_cell(row[5]))
        balance = first_numeric_token(clean_cell(row[6]))
        if is_date_like(first) and (is_date_like(val_date) or not val_date):
            if any(to_float(v) is not None for v in [debit, credit, balance]):
                return {
                    'date': 0,
                    'details': 1,
                    'reference': 2,
                    'value_date': 3,
                    'debit': 4,
                    'credit': 5,
                    'balance': 6,
                }
    return {}


def map_fits_rows(col_map, table_rows):
    """Check whether mapped indices are valid for a table shape."""
    if not col_map or not table_rows:
        return False
    max_idx = max((idx for idx in col_map.values() if idx is not None), default=-1)
    min_len = min(len(row) for row in table_rows if row)
    return max_idx < min_len


def sanitize_transaction_row(row, col_map):
    """Clean page/total artifacts and keep numeric/date columns parseable."""
    cleaned = [clean_cell(cell) for cell in row]
    for key in ('date', 'value_date'):
        idx = col_map.get(key)
        if idx is not None and idx < len(cleaned):
            value = cleaned[idx]
            value = re.sub(r"\btotal\b.*$", "", value, flags=re.I).strip()
            value = INLINE_PAGE_ARTIFACT_RE.sub('', value).strip()
            cleaned[idx] = value
    for key in ('debit', 'credit', 'balance'):
        idx = col_map.get(key)
        if idx is not None and idx < len(cleaned):
            value = re.sub(r"\btotal\b.*$", "", cleaned[idx], flags=re.I).strip()
            value = INLINE_PAGE_ARTIFACT_RE.sub('', value).strip()
            cleaned[idx] = first_numeric_token(value)
    return cleaned


def canonicalize_transaction_row(row, col_map):
    """Normalize transaction row into consistent output column order."""
    cleaned = sanitize_transaction_row(row, col_map)

    def pick(mapped_key):
        idx = col_map.get(mapped_key)
        if idx is not None and idx < len(cleaned):
            return cleaned[idx]
        return ""

    date_val = pick('date')
    details_val = pick('details')
    ref_val = pick('reference')
    value_date_val = pick('value_date')
    debit_val = pick('debit')
    credit_val = pick('credit')
    balance_val = pick('balance')

    date_cells = [cell for cell in cleaned if is_date_like(cell)]
    if not date_val and date_cells:
        date_val = date_cells[0]
    if not value_date_val and len(date_cells) > 1:
        value_date_val = date_cells[1]

    money_like = []
    for cell in cleaned:
        token = first_numeric_token(cell)
        if to_float(token) is None:
            continue
        # Avoid using long ID-like integers as money.
        if '.' in token or ',' in token or len(token.replace('-', '').replace('.', '').replace(',', '')) <= 7:
            money_like.append(token)
    if not balance_val and money_like:
        balance_val = money_like[-1]
    # Infer debit/credit only when both are unmapped and row clearly carries both amounts.
    if col_map.get('debit') is None and col_map.get('credit') is None and not debit_val and not credit_val:
        if len(money_like) >= 3:
            debit_val = money_like[-3]
            credit_val = money_like[-2]
    if credit_val and balance_val and credit_val == balance_val and debit_val:
        credit_val = ""

    if not details_val:
        candidates = []
        for cell in cleaned:
            text = str(cell).strip()
            if not text:
                continue
            if is_date_like(text):
                continue
            if to_float(first_numeric_token(text)) is not None:
                continue
            candidates.append(text)
        if candidates:
            details_val = max(candidates, key=len)

    if not ref_val:
        for cell in cleaned:
            text = str(cell).strip()
            if re.search(r"[A-Za-z]", text) and re.search(r"\d", text) and len(text) >= 8:
                if text != details_val:
                    ref_val = text
                    break

    # If reference looks like a date and value date doesn't, this is likely a shifted row with no reference.
    if is_date_like(ref_val) and (not value_date_val or not is_date_like(value_date_val)):
        value_date_val = ref_val
        ref_val = ""

    # Opening balance rows should not fabricate references.
    ref_looks_dateish = bool(re.match(r'^\d{1,2}[/-]\d{1,2}[/-]\d{2,4}', str(ref_val).strip()))
    if details_val and 'opening balance' in details_val.lower() and (is_date_like(ref_val) or ref_looks_dateish):
        ref_val = ""
    if details_val and 'closing balance' in details_val.lower() and (is_date_like(ref_val) or ref_looks_dateish):
        ref_val = ""

    return [date_val, details_val, ref_val, value_date_val, debit_val, credit_val, balance_val]


def should_keep_transaction_row(row):
    """Drop non-transaction noise rows that often appear in fragmented extracts."""
    date_val, details_val, _ref, _value_date, debit_val, credit_val, balance_val = row
    if is_date_like(date_val):
        return True
    if any(str(v).strip() for v in [debit_val, credit_val, balance_val]):
        details_text = str(details_val).lower().strip()
        if details_text and any(k in details_text for k in [
            'number of debit transaction', 'number of credit transaction', 'total credit amount',
            'total debit amount', 'period opening balance', 'period closing balance', 'transaction description'
        ]):
            return False
        return True
    details_text = str(details_val).lower().strip()
    if details_text and any(k in details_text for k in ['opening balance', 'closing balance', 'balance b/f']):
        return True
    return False


def extract_total_row_candidate(raw_row, col_map):
    """Extract a separate totals row when totals leak into transaction fields."""
    cleaned = [clean_cell(cell) for cell in raw_row]
    row_text = ' '.join(cleaned).lower()
    if 'total' not in row_text:
        return None

    debit_idx = col_map.get('debit')
    credit_idx = col_map.get('credit')
    balance_idx = col_map.get('balance')

    debit_text = cleaned[debit_idx] if debit_idx is not None and debit_idx < len(cleaned) else ""
    credit_text = cleaned[credit_idx] if credit_idx is not None and credit_idx < len(cleaned) else ""
    balance_text = cleaned[balance_idx] if balance_idx is not None and balance_idx < len(cleaned) else ""

    debit_tokens = [t for t in re.split(r"\s+", debit_text) if NUMERIC_RE.match(t)]
    total_debit = debit_tokens[1] if len(debit_tokens) > 1 else ""
    total_credit = first_numeric_token(credit_text) if credit_text else ""
    total_balance = first_numeric_token(balance_text) if balance_text else ""

    if not any([total_debit, total_credit]):
        return None

    return ["", "Total", "", "", total_debit, total_credit, total_balance]


def extract_text_fragments(table_rows):
    """Collect text-only fragments that likely continue previous transaction descriptions."""
    fragments = []
    for row in table_rows:
        cells = [clean_cell(cell) for cell in row if clean_cell(cell)]
        if not cells:
            continue
        text = ' '.join(cells).strip()
        if not text:
            continue
        has_date = any(is_date_like(cell) for cell in cells)
        has_numeric = any(to_float(first_numeric_token(cell)) is not None for cell in cells)
        if has_date or has_numeric:
            continue
        # Skip obvious statement chrome/meta rows.
        low = text.lower()
        if any(k in low for k in ['account state', 'summary statement', 'private & confidential']):
            continue
        fragments.append(text)
    return fragments


def append_fragments_to_last_transaction(transaction_data, fragments):
    """Attach orphan text fragments to the last transaction description."""
    if not transaction_data or not fragments:
        return
    # Last transaction details column is index 1 in canonical schema.
    last = transaction_data[-1]
    if len(last) < 2:
        return
    additions = [frag for frag in fragments if frag]
    if not additions:
        return
    detail = str(last[1]).strip()
    extra = ' '.join(additions).strip()
    if extra and extra.lower() not in detail.lower():
        last[1] = f"{detail} {extra}".strip()


def extract_period_balances_from_text(pdf):
    """Extract opening/closing period balances from page text when tables are weak."""
    opening = ""
    closing = ""
    for page in pdf.pages:
        text = page.extract_text() or ""
        if not text:
            continue
        # Normalize line breaks so labels like "Balance at Period S\ntart" match.
        flat = re.sub(r"\s+", " ", text)
        if not opening:
            m = re.search(
                r"balance\s+at\s+period\s+start[^\d-]*(-?\d{1,3}(?:,\d{3})*(?:\.\d{1,2})?)",
                flat,
                re.I,
            )
            if m:
                opening = m.group(1)
            if not opening:
                m = re.search(
                    r"opening\s+balance\s*[:\s]+(-?\d{1,3}(?:,\d{3})*(?:\.\d{1,2})?)",
                    flat,
                    re.I,
                )
                if m:
                    opening = m.group(1)
        if not closing:
            m = re.search(
                r"balance\s+at\s+period\s+end[^\d-]*(-?\d{1,3}(?:,\d{3})*(?:\.\d{1,2})?)",
                flat,
                re.I,
            )
            if m:
                closing = m.group(1)
            if not closing:
                m = re.search(
                    r"(?:balance at end of period|closing balance)[^\d-]*(-?\d{1,3}(?:,\d{3})*(?:\.\d{1,2})?)",
                    flat,
                    re.I,
                )
                if m:
                    closing = m.group(1)
    return opening, closing


def extract_bank_details_from_pdf_text(pdf):
    """
    Capture header lines above the transaction table when extract_tables() yields no usable bank block
    (common for borderless / Providus-style PDFs).
    """
    rows = []
    for page in pdf.pages[:1]:
        text = page.extract_text() or ""
        for raw in text.split("\n"):
            s = clean_cell(raw.strip())
            if not s:
                continue
            if re.search(r"Book\s+Date", s, re.I) and "Reference" in s and "Description" in s:
                break
            rows.append([s])
    return rows


def _pairify_bank_detail_text(line):
    """
    Split dense account-summary text into key/value pairs so details don't remain
    in one giant cell.
    """
    s = clean_cell(line)
    if not s:
        return []
    labels = [
        "Account Number",
        "Opening Balance",
        "Account Currency",
        "Withdrawal",
        "Account Type",
        "Deposit",
        "Account Nickname",
        "Closing Balance",
        "Branch",
        "Available Balance",
        "Account Summary Statement Period",
    ]
    pattern = re.compile("|".join(re.escape(lbl) for lbl in labels), re.I)
    matches = list(pattern.finditer(s))
    if len(matches) < 2:
        return []
    pairs = []
    for i, m in enumerate(matches):
        key = m.group(0).strip()
        start = m.end()
        end = matches[i + 1].start() if i + 1 < len(matches) else len(s)
        val = s[start:end].strip(" :;,-")
        if val:
            pairs.append([key, val])
    return pairs


def normalize_bank_details_rows(bank_details):
    """Convert single-cell bank-summary lines into two-column key/value rows when possible."""
    normalized = []
    for row in bank_details or []:
        if len(row) == 1:
            split_pairs = _pairify_bank_detail_text(row[0])
            if split_pairs:
                normalized.extend(split_pairs)
                continue
        normalized.append(row)
    return normalized


def extract_transaction_header_from_pdf_text(pdf):
    """
    Extract an "original-ish" display header from page text.

    We still map internal parsing to the canonical output schema elsewhere; this
    function only aims to preserve the header wording for what the user sees.
    """
    for page in pdf.pages[:1]:
        text = page.extract_text() or ""
        lines = [clean_cell(ln) for ln in text.split("\n") if clean_cell(ln)]
        header_line_idx = None
        for i, ln in enumerate(lines):
            low = ln.lower()
            if ("transaction" in low) and ("withdrawal" in low) and ("deposit" in low) and ("balance" in low):
                header_line_idx = i
                break
        if header_line_idx is None:
            continue

        raw = lines[header_line_idx]
        # Attempt to split the header line into canonical 7 cells by keyword boundaries.
        # Note: some formats merge "reference" into other columns; we preserve wording
        # as best-effort while keeping a stable 7-column shape for the UI.
        low = raw.lower()

        def cut_between(start_pat, end_pats):
            """Return substring starting at start_pat up to first occurrence of any end_pats."""
            m = re.search(start_pat, raw, flags=re.I)
            if not m:
                return ""
            start = m.start()
            end = len(raw)
            for ep in end_pats:
                em = re.search(ep, raw[m.end() :], flags=re.I)
                if em:
                    end = min(end, m.end() + em.start())
            return raw[start:end].strip()

        # Canonical columns in output order:
        # [Transaction Date, Details, Reference, Value Date, Debit, Credit, Balance]
        date_cell = cut_between(r"transaction\s+(value\s+date|date)", [r"cheque", r"remarks", r"reference", r"value\s+date"])
        debit_cell = cut_between(r"withdrawal", [r"deposit", r"balance", r"credit"])
        credit_cell = cut_between(r"deposit", [r"balance", r"debit", r"credit"])
        balance_cell = cut_between(r"balance", [])

        # Everything between date and debit is treated as details/reference wording.
        before_debit = raw[: re.search(r"withdrawal", raw, flags=re.I).start()] if re.search(r"withdrawal", raw, flags=re.I) else raw
        after_date = before_debit

        # Try to split details/reference if we can see either "reference" or "cheque"/"remarks".
        ref_seen = re.search(r"reference", after_date, flags=re.I)
        cheque_seen = re.search(r"cheque", after_date, flags=re.I)
        remarks_seen = re.search(r"remarks", after_date, flags=re.I)

        details_cell = ""
        reference_cell = ""
        if ref_seen:
            # Split around the word "Reference"
            details_cell = after_date[: ref_seen.start()].strip()
            reference_cell = after_date[ref_seen.start() :].strip()
        elif cheque_seen and remarks_seen:
            # UBA-style: "Cheque Transaction Remarks"
            # Keep both phrases to avoid the reference column becoming empty and
            # later being dropped by `drop_empty_columns`.
            details_cell = after_date[cheque_seen.start() : remarks_seen.start()].strip()
            reference_cell = after_date[remarks_seen.start() : remarks_seen.end()].strip()
        else:
            # Best-effort: keep whatever remains as Details.
            details_cell = after_date.strip()

        # Value date sometimes reuses the same wording as transaction date in these formats.
        value_date_cell = re.search(r"value\s+date", raw, flags=re.I)
        value_date_cell = value_date_cell.group(0).title() if value_date_cell else (date_cell or "Value Date")

        # Cleanup to avoid giant merged phrases in cells.
        def normalize_cell(s):
            s = clean_cell(s)
            # Remove obvious repeated "Transaction"
            return re.sub(r"^transaction\s+", "", s, flags=re.I).strip()

        out = [
            normalize_cell(date_cell or "Transaction Date"),
            normalize_cell(details_cell),
            normalize_cell(reference_cell),
            normalize_cell(value_date_cell),
            normalize_cell(debit_cell or "Withdrawal"),
            normalize_cell(credit_cell or "Deposit"),
            normalize_cell(balance_cell or "Balance"),
        ]
        return [out]
    return []


def _header_is_sane_for_display(header_row: list[str]) -> bool:
    """Return True if a header row looks like a transaction header (not just a random bank/details line)."""
    if not header_row:
        return False
    non_empty_count = sum(1 for c in header_row if str(c).strip())
    text = " ".join(str(c).strip().lower() for c in header_row if str(c).strip())
    if not text:
        return False
    # Require at least two "transaction-ish" keywords to avoid accidental display headers.
    keywords = [
        "transaction",
        "value date",
        "details",
        "reference",
        "withdrawal",
        "deposit",
        "debit",
        "credit",
        "balance",
        "remarks",
        "cheque",
    ]
    hits = sum(1 for k in keywords if k in text)
    # Also accept if it has money column terms strongly.
    if hits < 2:
        return False
    # If the header is effectively a single merged cell (very common with weak
    # table extraction), avoid selecting it for display. We prefer headers
    # that actually split into multiple cells.
    if len(header_row) >= 5:
        return True
    return non_empty_count >= 3


def _ledger_header_ref_before_description(header_row):
    """True when header cells follow Book Date | Reference | Description (Reference left of Description)."""
    if not header_row:
        return False
    idx_ref = idx_desc = None
    for i, cell in enumerate(header_row):
        if cell is None:
            continue
        t = str(cell).strip().lower()
        if t == "reference" or (t.startswith("reference") and "description" not in t):
            idx_ref = i
        if "description" in t and "reference" not in t:
            idx_desc = i
    return idx_ref is not None and idx_desc is not None and idx_ref < idx_desc


def _is_transaction_ledger_header_row(row):
    if not row:
        return False
    t = " ".join(str(c).lower() for c in row if c)
    return "book date" in t and "reference" in t and "description" in t


def _row_has_period_opening_label(row):
    if not row or len(row) < 2:
        return False
    blob = " ".join(str(c).lower() for c in row[:4] if c)
    return "opening balance" in blob or "balance at period start" in blob


def _row_has_period_closing_label(row):
    if not row or len(row) < 2:
        return False
    blob = " ".join(str(c).lower() for c in row[:4] if c)
    return "closing balance" in blob or "balance at period end" in blob


def parse_transactions_from_page_text(pdf):
    """Parse transaction lines directly from page text for weak/no-border statements."""
    tx_rows = []
    money_re = re.compile(r'\d{1,3}(?:,\d{3})*(?:\.\d{2})')
    row_start_re = re.compile(
        r'^(?P<date>\d{1,2}/\d{1,2}/\d{4})\s+(?P<value_date>\d{1,2}/\d{1,2}/\d{4})\s+(?P<rest>.+)$',
        re.I,
    )
    stop_markers = (
        "account summary statement",
        "your transactions",
        "transaction value date",
        "date number",
        "page ",
    )
    for page in pdf.pages:
        text = page.extract_text() or ""
        if not text:
            continue
        in_transactions = False
        for raw_line in text.split('\n'):
            line = clean_cell(raw_line)
            if not line:
                continue
            low_line = line.lower()
            if "your transactions" in low_line:
                in_transactions = True
                continue
            if not in_transactions:
                continue
            if any(marker in low_line for marker in stop_markers):
                continue

            m = row_start_re.match(line)
            if not m:
                # Continuation text line: append to previous transaction details.
                if tx_rows and not re.search(r'\d{1,3}(?:,\d{3})*(?:\.\d{2})', line):
                    if not any(k in low_line for k in ("balance", "opening", "closing", "account", "statement")):
                        tx_rows[-1][1] = f"{tx_rows[-1][1]} {line}".strip()
                continue
            rest = m.group('rest')
            money_hits = list(money_re.finditer(rest))
            if len(money_hits) < 2:
                continue
            balance_val = money_hits[-1].group(0)
            amount_val = money_hits[-2].group(0)
            details_val = clean_cell(rest[: money_hits[-2].start()].rstrip(" -"))
            if not details_val:
                continue

            lower_details = details_val.lower()
            is_credit_like = any(
                k in lower_details for k in ['transfer in', 'trf from', 'from ', 'lodgement', 'deposit', 'credit']
            )
            debit_val = "" if is_credit_like else amount_val
            credit_val = amount_val if is_credit_like else ""
            tx_rows.append([
                m.group('date'),
                details_val,
                "",
                m.group('value_date'),
                debit_val,
                credit_val,
                balance_val,
            ])
    return tx_rows


def merge_missing_text_transactions(transaction_data, text_transactions):
    """Merge text-parsed transactions that are not already present in extracted table rows."""
    if not text_transactions:
        return transaction_data

    def ref_key(ref):
        cleaned = clean_cell(ref).upper()
        if not cleaned:
            return ""
        m = re.match(r'[A-Z0-9]{8,}', cleaned)
        if m:
            return m.group(0)
        return re.sub(r'[^A-Z0-9]', '', cleaned)

    existing_keys = set()
    for row in transaction_data:
        key = (
            clean_cell(row[0]) if len(row) > 0 else "",
            ref_key(row[2]) if len(row) > 2 else "",
            clean_cell(row[3]) if len(row) > 3 else "",
            clean_cell(row[6]) if len(row) > 6 else "",
        )
        existing_keys.add(key)

    merged = list(transaction_data)
    for row in text_transactions:
        key = (
            clean_cell(row[0]),
            ref_key(row[2]),
            clean_cell(row[3]),
            clean_cell(row[6]),
        )
        if key in existing_keys:
            continue
        merged.append(row)
        existing_keys.add(key)

    # Keep transactions in chronological/text order by date then reference where possible.
    def sort_key(r):
        date_val = clean_cell(r[0]) if len(r) > 0 else ""
        ref_val = clean_cell(r[2]) if len(r) > 2 else ""
        return (date_val, ref_val)

    return sorted(merged, key=sort_key)


def deduplicate_transactions(transaction_data):
    """Remove duplicate transactions while keeping the richest row variant."""
    if not transaction_data:
        return transaction_data

    def ref_key(ref):
        cleaned = clean_cell(ref).upper()
        if not cleaned:
            return ""
        m = re.match(r'[A-Z0-9]{8,}', cleaned)
        if m:
            return m.group(0)
        return re.sub(r'[^A-Z0-9]', '', cleaned)

    by_key = {}
    order = []
    for row in transaction_data:
        date_val = clean_cell(row[0]) if len(row) > 0 else ""
        details_val = clean_cell(row[1]) if len(row) > 1 else ""
        ref_val = clean_cell(row[2]) if len(row) > 2 else ""
        value_date_val = clean_cell(row[3]) if len(row) > 3 else ""
        balance_val = clean_cell(row[6]) if len(row) > 6 else ""
        key = (date_val, ref_key(ref_val), value_date_val, balance_val)
        if not key[1]:
            key = (date_val, details_val.lower(), value_date_val, balance_val)
        if key not in by_key:
            by_key[key] = row
            order.append(key)
            continue
        existing = by_key[key]
        existing_score = len(clean_cell(existing[1])) + (5 if clean_cell(existing[2]) else 0)
        new_score = len(details_val) + (5 if ref_val else 0)
        if new_score > existing_score:
            by_key[key] = row

    return [by_key[k] for k in order]


def is_summary_table(table_rows):
    """Detect whether a table is a summary block, not the main transaction sheet."""
    if not table_rows:
        return False
    text = ' '.join(str(cell).lower() for row in table_rows for cell in row if cell)
    keywords = [
        'number of debit transaction', 'period opening balance', 'total debit amount', 'total credit amount',
        'opening balance', 'closing balance', 'total withdrawals', 'total lodgements', 'account statement',
        'summary details', 'account no.', 'alt. account no.', 'currency', 'transaction description'
    ]
    return any(k in text for k in keywords)


def is_continuation_row(row):
    """Check if a row is a continuation (no date-like pattern in first cell)."""
    if not row or not row[0]:
        return True
    first_cell = str(row[0]).strip()
    
    # If this is a closing/total row, it's NOT a continuation - it's a standalone summary
    if is_total_or_closing_row(row):
        return False
    
    # Check if it's a header row (contains common header keywords)
    # Look across the row for header patterns
    row_text = ' '.join(str(c).lower() for c in row if c)
    header_keywords = ['transaction date', 'transaction details', 'reference', 'value date', 
                      'date', 'debit amount', 'credit amount', 'balance', 'withdrawals', 'lodgements']
    if any(keyword in row_text for keyword in header_keywords):
        # Check if this looks like a header (first cell matches common header patterns)
        if re.match(r'^(date|transaction|reference|value|debit|credit|balance|withdrawal|lodgement|actual)', 
                   first_cell, re.IGNORECASE):
            return False  # This is a header, not a continuation
    
    # Date patterns: d/d/yyyy or dd/dd/yyyy OR d MMM d OR DD MMM YY or similar text dates
    # Match numeric: 1/1/2026 or 01/01/2026
    # Match text months: 05 FEB 26, 05 Feb 2026, 1-Jan-2026, etc.
    date_pattern = r'^(\d{1,2}[/-]\d{1,2}[/-]\d{2,4}|\d{1,2}\s+(?:JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)\s+\d{2,4}|\d{1,2}-(?:JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)-\d{4})'
    return not re.match(date_pattern, first_cell, re.IGNORECASE)


def merge_continuation_rows(table_rows):
    """Merge rows that are continuations of previous transactions."""
    if not table_rows:
        return []
    merged = []
    current_row = None
    for row in table_rows:
        if is_continuation_row(row):
            if current_row is not None:
                # Append this continuation row's non-empty cells to current row
                for i, cell in enumerate(row):
                    cell_str = str(cell).strip() if cell else ""
                    if cell_str and i < len(current_row):
                        # Append to existing cell with a space
                        current_part = str(current_row[i]).strip() if current_row[i] else ""
                        if current_part:
                            current_row[i] = f"{current_part} {cell_str}"
                        else:
                            current_row[i] = cell_str
        else:
            if current_row is not None:
                merged.append(current_row)
            current_row = [clean_cell(cell) for cell in row]
    if current_row is not None:
        merged.append(current_row)
    return merged


def normalize_table(table_rows):
    """Clean rows, merge continuations (for date-based tables), remove page markers, and drop empty columns."""
    cleaned = []
    for row in table_rows:
        if not row:
            continue
        cleaned_row = [clean_cell(cell) for cell in row]
        if is_page_marker_row(cleaned_row):
            continue
        if not any(cell.strip() for cell in cleaned_row):
            continue
        cleaned.append(cleaned_row)
    
    # Check if this table has date-based rows (transaction-like structure)
    has_dates = any(re.match(r'^\d{1,2}[/-]', str(row[0]).strip()) for row in cleaned if row and row[0])
    
    # Only merge continuation rows for date-based tables (transactions)
    if has_dates:
        merged = merge_continuation_rows(cleaned)
    else:
        # For non-date tables (summaries, bank details), keep all rows as-is
        merged = cleaned
    
    return drop_empty_columns(merged)


def ensure_transaction_headers(transaction_headers, transaction_data, preferred_header=None):
    """
    Ensure transaction headers exist.

    `preferred_header` is the *display* header row chosen by header source priority.
    We only fall back to the canonical/default header if `preferred_header` is missing
    or looks non-informative.
    """
    if not transaction_data and not transaction_headers and not preferred_header:
        return transaction_headers

    # Standardized header used when bank PDFs omit/garble a clear header row in extracted tables.
    standard_header = [
        'Transaction Date',
        'Transaction Details',
        'Reference',
        'Value Date',
        'Debit Amount',
        'Credit Amount',
        'Current Balance'
    ]

    max_cols = len(standard_header)
    if transaction_data:
        max_cols = max(max_cols, max(len(row) for row in transaction_data))

    def pad_header(header_row):
        return header_row + [''] * max(0, max_cols - len(header_row))

    if preferred_header and _header_is_sane_for_display(preferred_header):
        return [pad_header(list(preferred_header))]
    # If caller already provided a transaction header row, keep it when sane.
    if transaction_headers:
        hdr = transaction_headers[0]
        if hdr and _header_is_sane_for_display(list(hdr)):
            return [pad_header(list(hdr))]

    synthesized = pad_header(list(standard_header))
    return [synthesized]


def extract_all_tables_from_pdf(pdf_path, combine_tables=False):
    """Extract bank details, headers, and all transaction data into a single sheet."""
    bank_details = []
    transaction_headers = []  # display header (chosen near the end)
    transaction_data = []
    transaction_col_map = {}
    
    # Track if we've found the first transaction table to extract its header
    found_first_trans_table = False
    first_transaction_table_header_row = None
    text_header_row = None

    with pdfplumber.open(pdf_path) as pdf:
        opening_from_text, closing_from_text = extract_period_balances_from_text(pdf)
        text_transactions = parse_transactions_from_page_text(pdf)
        text_header_rows = extract_transaction_header_from_pdf_text(pdf)
        if text_header_rows and text_header_rows[0]:
            text_header_row = list(text_header_rows[0])
        for page in pdf.pages:
            page_tables = page.extract_tables()
            for idx, table in enumerate(page_tables):
                if not table:
                    continue
                
                # Try to extract bank details from raw table first (before normalization)
                if not bank_details:
                    raw_bank_details = extract_bank_details_from_raw_table(table)
                    if raw_bank_details:
                        bank_details.extend(raw_bank_details)
                
                normalized = normalize_table(table)
                if not normalized:
                    continue
                
                # Check table classification
                is_trans = is_transaction_table(normalized)
                is_summ = is_summary_table(normalized)
                
                # If small table but has consistent structure with nearby tables, might be fragment
                is_fragment = (len(normalized) < 5 and 
                              idx > 0 and 
                              len(normalized[0]) >= 6 and 
                              any(cell for row in normalized for cell in row))
                
                if is_summ and not bank_details and not found_first_trans_table:
                    # Store summary/bank details at the top (for cases where normalize_table preserves them)
                    bank_details.extend(normalized)
                elif is_trans or (is_fragment and transaction_data):
                    # For first transaction table, capture header
                    if not found_first_trans_table and is_trans:
                        header_idx = detect_transaction_header_index(normalized)
                        if header_idx is not None:
                            table_header_row = normalized[header_idx]
                            if first_transaction_table_header_row is None and table_header_row:
                                first_transaction_table_header_row = list(table_header_row)
                            transaction_col_map = map_transaction_columns(table_header_row)
                            for row in normalized[header_idx + 1:]:
                                if looks_like_transaction_header_row(row):
                                    continue
                                total_row = extract_total_row_candidate(row, transaction_col_map)
                                canon = canonicalize_transaction_row(row, transaction_col_map)
                                if total_row and canon[5] == total_row[5]:
                                    canon[5] = ""
                                if should_keep_transaction_row(canon):
                                    transaction_data.append(canon)
                                if total_row:
                                    transaction_data.append(total_row)
                        else:
                            # Some PDFs start directly with data rows; keep all rows as transaction data.
                            inferred = infer_transaction_columns_from_rows(normalized)
                            if inferred and (not transaction_col_map or not map_fits_rows(transaction_col_map, normalized)):
                                transaction_col_map = inferred
                            for row in normalized:
                                total_row = extract_total_row_candidate(row, transaction_col_map)
                                canon = canonicalize_transaction_row(row, transaction_col_map)
                                if total_row and canon[5] == total_row[5]:
                                    canon[5] = ""
                                if should_keep_transaction_row(canon):
                                    transaction_data.append(canon)
                                if total_row:
                                    transaction_data.append(total_row)
                        found_first_trans_table = True
                    elif is_fragment and transaction_data:
                        # Fragment: might have header-like first row, skip if we already have headers
                        if normalized and transaction_headers:
                            # Skip header-like rows in fragments
                            start_row = 0
                            if looks_like_transaction_header_row(normalized[0]):
                                start_row = 1
                            local_map = transaction_col_map or infer_transaction_columns_from_rows(normalized)
                            for row in normalized[start_row:]:
                                total_row = extract_total_row_candidate(row, local_map)
                                canon = canonicalize_transaction_row(row, local_map)
                                if total_row and canon[5] == total_row[5]:
                                    canon[5] = ""
                                if should_keep_transaction_row(canon):
                                    transaction_data.append(canon)
                                if total_row:
                                    transaction_data.append(total_row)
                        else:
                            local_map = infer_transaction_columns_from_rows(normalized)
                            if local_map and (not transaction_col_map or not map_fits_rows(transaction_col_map, normalized)):
                                transaction_col_map = local_map
                            for row in normalized:
                                total_row = extract_total_row_candidate(row, transaction_col_map)
                                canon = canonicalize_transaction_row(row, transaction_col_map)
                                if total_row and canon[5] == total_row[5]:
                                    canon[5] = ""
                                if should_keep_transaction_row(canon):
                                    transaction_data.append(canon)
                                if total_row:
                                    transaction_data.append(total_row)
                    elif is_trans and found_first_trans_table:
                        header_idx = detect_transaction_header_index(normalized)
                        if header_idx is not None:
                            table_header_row = normalized[header_idx]
                            if first_transaction_table_header_row is None and table_header_row:
                                first_transaction_table_header_row = list(table_header_row)
                            transaction_col_map = map_transaction_columns(table_header_row)
                        else:
                            inferred = infer_transaction_columns_from_rows(normalized)
                            if inferred and (not transaction_col_map or not map_fits_rows(transaction_col_map, normalized)):
                                transaction_col_map = inferred
                        start_row = header_idx + 1 if header_idx is not None else 0
                        for row in normalized[start_row:]:
                            if looks_like_transaction_header_row(row):
                                continue
                            total_row = extract_total_row_candidate(row, transaction_col_map)
                            canon = canonicalize_transaction_row(row, transaction_col_map)
                            if total_row and canon[5] == total_row[5]:
                                canon[5] = ""
                            if should_keep_transaction_row(canon):
                                transaction_data.append(canon)
                            if total_row:
                                transaction_data.append(total_row)
                elif transaction_data:
                    # Weak-border PDFs can split long descriptions into separate text-only mini tables.
                    fragments = extract_text_fragments(normalized)
                    append_fragments_to_last_transaction(transaction_data, fragments)

        text_bank = extract_bank_details_from_pdf_text(pdf)
        if not bank_details:
            bank_details = text_bank
        # Never replace table-derived bank_details with page-1 text unless the coordinate gate
        # will run: extract_bank_details_from_pdf_text only stops at "Book Date…Reference…Description";
        # PDFs without that line (e.g. Access) would otherwise ingest the whole page as fake "bank" rows.
        elif should_use_coordinate_fallback(pdf_path, len(transaction_data)) and len(text_bank) > len(
            bank_details
        ):
            bank_details = text_bank

    primary_transaction_count = len(transaction_data)
    used_coordinate_fallback = False
    if should_use_coordinate_fallback(pdf_path, primary_transaction_count):
        try:
            coord_rows = reconstruct_transactions_coordinate(pdf_path)
        except Exception:
            coord_rows = []
        if coord_rows:
            transaction_data = [list(row) for row in coord_rows]
            used_coordinate_fallback = True

    # Only synthesize period rows when we already have at least one transaction row.
    has_real_transaction_row = any(
        is_date_like(row[0]) for row in transaction_data if row
    )
    # Ensure opening/closing are present when found in plain text extraction.
    has_opening_row = any(_row_has_period_opening_label(row) for row in transaction_data)
    if has_real_transaction_row and opening_from_text and not has_opening_row:
        if used_coordinate_fallback:
            # Label lives in Reference column (index 2); swap below aligns with PDF header order.
            transaction_data.insert(0, ["", "", "Balance at Period Start", "", "0.00", "0.00", opening_from_text])
        else:
            transaction_data.insert(0, ["", "Balance at Period Start", "", "", "0.00", "0.00", opening_from_text])

    has_closing_row = any(_row_has_period_closing_label(row) for row in transaction_data)
    if has_real_transaction_row and closing_from_text and not has_closing_row:
        if used_coordinate_fallback:
            transaction_data.append(["", "", "Balance at Period End", "", "", "", closing_from_text])
        else:
            transaction_data.append(["", "Balance at Period End", "", "", "", "", closing_from_text])

    # If statement text does not expose an explicit closing-balance line, synthesize from the last balance value.
    if has_real_transaction_row and not has_closing_row and transaction_data:
        last_balance = ""
        for row in reversed(transaction_data):
            if len(row) > 6 and clean_cell(row[6]):
                last_balance = clean_cell(row[6])
                break
        has_opening_row = any(_row_has_period_opening_label(row) for row in transaction_data)
        if last_balance and has_opening_row:
            if used_coordinate_fallback:
                transaction_data.append(["", "", "Balance at Period End", "", "", "", last_balance])
            else:
                transaction_data.append(["", "Balance at Period End", "", "", "", "", last_balance])

    if not used_coordinate_fallback:
        transaction_data = merge_missing_text_transactions(transaction_data, text_transactions)
    transaction_data = deduplicate_transactions(transaction_data)

    # Align with PDF column order Reference | Description (cols 2–3) for coordinate extraction.
    if used_coordinate_fallback:
        for row in transaction_data:
            if len(row) >= 3:
                row[1], row[2] = row[2], row[1]

    preferred_header_row = None
    if first_transaction_table_header_row and _header_is_sane_for_display(first_transaction_table_header_row):
        preferred_header_row = first_transaction_table_header_row
    elif text_header_row and _header_is_sane_for_display(text_header_row):
        preferred_header_row = text_header_row

    transaction_headers = ensure_transaction_headers(
        transaction_headers,
        transaction_data,
        preferred_header=preferred_header_row,
    )

    # Build the final sheet: bank details → headers → transactions
    sheet_data = []
    if bank_details:
        bank_details = normalize_bank_details_rows(bank_details)
        filtered_bank = []
        th0 = transaction_headers[0] if transaction_headers else None
        for r in bank_details:
            if th0 and r == th0:
                continue
            if th0 and _is_transaction_ledger_header_row(r) and _is_transaction_ledger_header_row(th0):
                continue
            filtered_bank.append(r)
        bank_details = filtered_bank
    if bank_details:
        sheet_data.extend(bank_details)
    
    if transaction_headers:
        sheet_data.extend(transaction_headers)
    
    if transaction_data:
        sheet_data.extend(transaction_data)

    # Normalize rows to a common width, then drop fully empty columns from the final sheet.
    if sheet_data:
        max_cols = max(len(row) for row in sheet_data)
        sheet_data = [row + [''] * (max_cols - len(row)) for row in sheet_data]
        sheet_data = drop_empty_columns(sheet_data)

    sheets = []
    if sheet_data:
        sheets.append({'name': 'Statement', 'data': sheet_data, 'is_table': True})

    return sheets


def export_to_excel(sheets_data, output_path):
    """Write sheets to Excel file with basic formatting."""
    wb = Workbook()
    wb.remove(wb.active)
    header_fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
    header_font = Font(bold=True, color='FFFFFF')
    center_alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

    for sheet_info in sheets_data:
        sheet_name = sheet_info['name']
        data = sheet_info['data']
        sheet_name = sheet_name[:31].replace('[', '').replace(']', '').replace(':', '').replace('*', '').replace('?', '').replace('/', '')
        ws = wb.create_sheet(sheet_name)
        for row_idx, row in enumerate(data, 1):
            for col_idx, cell_value in enumerate(row, 1):
                cell = ws.cell(row=row_idx, column=col_idx, value=cell_value)
                if row_idx == 1:
                    cell.fill = header_fill
                    cell.font = header_font
                    cell.alignment = center_alignment
        for col_idx, col in enumerate(ws.columns, 1):
            max_length = 0
            column_letter = get_column_letter(col_idx)
            for cell in col:
                try:
                    if len(str(cell.value or '')) > max_length:
                        max_length = len(str(cell.value or ''))
                except:
                    pass
            ws.column_dimensions[column_letter].width = min(max_length + 2, 50)
    wb.save(output_path)


def get_sheet_preview(sheets_data, max_rows=10):
    """Get preview of sheets for display."""
    previews = []
    for sheet_info in sheets_data:
        data = sheet_info['data']
        previews.append({'name': sheet_info['name'], 'total_rows': len(data), 'columns': len(data[0]) if data else 0, 'preview': data[:max_rows]})
    return previews
