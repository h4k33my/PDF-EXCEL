#!/usr/bin/env python3
"""
Regression and Zapphaire validation for coordinate-based PDF fallback.
Run from project root: python test_coordinate_zapphaire.py
"""
from __future__ import annotations

import os
import sys

ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(ROOT, "src"))

from coordinate_fallback import (  # noqa: E402
    count_probable_transaction_lines_in_text,
    should_use_coordinate_fallback,
)
from converter import extract_all_tables_from_pdf  # noqa: E402


def _pdf(name: str) -> str:
    return os.path.join(ROOT, name)


def _find_ledger_header_index(data: list) -> int:
    for i, row in enumerate(data):
        if not row:
            continue
        if str(row[0]).strip().lower() == "book date":
            return i
        if str(row[0]).strip() == "Transaction Date":
            return i
    return -1


def _transaction_body_rows(sheets: list) -> list:
    data = sheets[0]["data"] if sheets else []
    start = _find_ledger_header_index(data)
    if start < 0:
        return data
    return data[start + 1 :]


def test_zapphaire_coordinate_path():
    path = _pdf("ZAPPHAIRE EVENTS LIMITED STATEMENT OF ACCOUNT.pdf")
    assert os.path.isfile(path), f"Missing PDF: {path}"

    primary_like = 8  # table+fragment count before merge (documented baseline)
    assert should_use_coordinate_fallback(path, primary_like) is True
    exp = count_probable_transaction_lines_in_text(path)
    assert exp >= 12, exp

    sheets = extract_all_tables_from_pdf(path)
    assert sheets, "expected non-empty sheet"
    data = sheets[0]["data"]

    # Single PDF-style ledger header (Book Date …); no duplicate synthetic header row.
    book_hdr = [r for r in data if r and str(r[0]).strip().lower() == "book date"]
    txn_hdr = [r for r in data if r and str(r[0]).strip() == "Transaction Date"]
    assert len(book_hdr) == 1, "expected exactly one Book Date header row"
    assert len(txn_hdr) == 0, "synthetic Transaction Date header should not appear for Zapphaire"

    top_blob = " ".join(str(c) for row in data[:12] for c in row if c)
    assert "Account" in top_blob or "Statement" in top_blob, "expected bank details lines at top"

    body = _transaction_body_rows(sheets)
    txn_rows = [
        r
        for r in body
        if r
        and "balance at period end" not in str(r[1]).lower()
        and "balance at period start" not in str(r[1]).lower()
    ]
    assert len(txn_rows) >= 14, f"expected >= 14 transaction rows, got {len(txn_rows)}"


def test_regression_no_gate_on_known_good_pdfs():
    """Gate must stay false; row counts must match pre-fallback baselines."""
    cases = [
        ("MARCH BANK STATEMENT.pdf", 50),
        ("1a70169f-341a-482a-9554-07aa599aebe3 (1).pdf", 77),
        ("40d4a207-1ca1-41d7-ab82-ec3b8061230a (1).pdf", 116),
    ]
    for name, expected_rows in cases:
        path = _pdf(name)
        if not os.path.isfile(path):
            continue
        sheets = extract_all_tables_from_pdf(path)
        body = _transaction_body_rows(sheets)
        assert len(body) == expected_rows, f"{name}: expected {expected_rows} rows after header, got {len(body)}"
        # Conservative gate: pattern-based expected line count is 0 for these formats
        assert count_probable_transaction_lines_in_text(path) == 0
        assert should_use_coordinate_fallback(path, len(body)) is False


def main():
    test_zapphaire_coordinate_path()
    test_regression_no_gate_on_known_good_pdfs()
    print("All coordinate fallback tests passed.")


if __name__ == "__main__":
    main()
