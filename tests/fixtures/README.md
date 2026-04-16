# Test Fixtures

This folder is for safe, dummy, or fully sanitized fixtures that can be committed.

Recommended layout:

- `pdfs/`: sample input PDFs used for regression testing
- `expected/`: expected parsed output for each sample input

Rules:

- Do not place real customer or bank statement files here.
- Only add dummy or sanitized files that are safe to share.
- Match fixture names where possible, for example:
  - `pdfs/uba_dummy_statement.pdf`
  - `expected/uba_dummy_statement_expected.json`
