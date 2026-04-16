#!/usr/bin/env python3
"""
Test script to convert MARCH BANK STATEMENT.pdf and inspect the output
"""
import sys
import os
sys.path.append(os.path.join(os.path.dirname(__file__), '..'))

from src.converter import extract_all_tables_from_pdf, export_to_excel, get_sheet_preview

def test_conversion():
    pdf_path = r"c:\Users\DELL\Desktop\ZAPPHAIRE EVENTS LIMITED STATEMENT OF ACCOUNT.pdf"
    output_path = r"c:\Users\DELL\Desktop\ZAPPHAIRE_TEST_CONVERSION_OUTPUT.xlsx"

    print(f"Testing conversion of: {pdf_path}")
    print(f"Output will be saved to: {output_path}")

    # Extract tables
    sheets_data = extract_all_tables_from_pdf(pdf_path, combine_tables=True)
    print(f"\nExtracted {len(sheets_data)} sheets:")

    for sheet in sheets_data:
        print(f"- {sheet['name']}: {len(sheet['data'])} rows, {len(sheet['data'][0]) if sheet['data'] else 0} columns")

    # Get preview
    previews = get_sheet_preview(sheets_data, max_rows=5)
    print("\nSheet Previews:")
    for preview in previews:
        print(f"\n{preview['name']} (Total: {preview['total_rows']} rows, {preview['columns']} columns):")
        for row in preview['preview']:
            print("  " + " | ".join(str(cell)[:50] for cell in row))

    # Export to Excel
    export_to_excel(sheets_data, output_path)
    print(f"\nExported to Excel: {output_path}")

    return sheets_data

if __name__ == "__main__":
    test_conversion()