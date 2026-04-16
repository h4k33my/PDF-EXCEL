# Bank Statement PDF-to-Excel Converter

A lightweight desktop application that extracts all data from bank statement PDFs and exports to organized Excel spreadsheets. Designed for easy use in small offices.

## Features

- **Extract All Tables**: Automatically detects and extracts all tables from PDF bank statements
- **Interactive Preview**: View extracted data before saving
- **Selective Export**: Choose which sheets to include in the final Excel file
- **Professional Formatting**: Auto-formatted Excel output with proper headers and column widths
- **No Installation Required**: Standalone `.exe` for Windows (no Python needed)

## System Requirements

- **Windows 7 or later** (for distributed .exe version)
- **Minimum RAM**: 2GB
- **Disk Space**: ~150MB for executable

## Installation & Usage

### For End Users (Using .exe)

1. Download `GAC-PDF-EXCEL-CONVERTER.exe` from the project folder
2. Double-click to launch the application
3. **Select PDF**: Click "Browse..." and choose your bank statement PDF
4. **Review Preview**: Check the extracted tables in the tabs
5. **Select Sheets**: Check/uncheck sheets you want to include
6. **Save as Excel**: Specify output filename and location
7. **Convert & Save**: Click the button to generate Excel file

### For Developers (Source Code)

#### Prerequisites
- Python 3.11 or higher
- pip (Python package manager)

#### Setup

```bash
# Navigate to project directory
cd BankStatementConverter

# Install dependencies
pip install -r requirements.txt

# Run the application
python src/main.py
```

#### Building Standalone .exe

```bash
# Install PyInstaller
pip install pyinstaller

# Build executable
python build.py

# Output: dist/GAC-PDF-EXCEL-CONVERTER.exe (copy also in dist_package/)
```

## Project Structure

```
BankStatementConverter/
├── src/
│   ├── main.py                 # Application entry point
│   ├── converter.py            # PDF extraction & Excel export logic
│   ├── coordinate_fallback.py # Coordinate-based fallback (borderless / weak tables)
│   ├── ui/
│   │   ├── __init__.py
│   │   └── main_window.py      # PyQt6 GUI
│   └── utils/
│       ├── __init__.py
│       └── file_handler.py     # File utilities
├── test_coordinate_zapphaire.py  # Zapphaire + regression tests
├── export_all_test_pdfs.py       # Batch-export all root PDFs → test_excel_outputs/
├── COORDINATE_FALLBACK_REPORT.md # Fallback behaviour & metrics
├── requirements.txt            # Python dependencies
├── build.py                    # PyInstaller → dist/ + dist_package/
└── README.md                   # This file
```

## How It Works

1. **PDF Parsing**: Uses `pdfplumber` to extract tables; for some Providus-style statements a **coordinate fallback** (word positions) replaces weak table output when a conservative gate fires.
2. **Data Cleaning**: Removes page artifacts, normalizes formatting, and cleans whitespace
3. **Preview**: Displays extracted tables in tabbed interface for review
4. **Selection**: Users can choose which sheets/tables to include in final output
5. **Excel Export**: Uses `openpyxl` to write clean, formatted Excel files with multiple sheets

## Features Explained

### Preview Section
- Shows all extracted tables/sheets from the PDF
- Each sheet appears in a separate tab
- Headers are highlighted in blue
- Shows total row count and column count for each sheet

### Sheet Selection
- Checkboxes let you include/exclude sheets
- All sheets selected by default
- Uncheck sheets you don't want in the final Excel file

### Output File
- Default saves to Desktop as `bank_statements.xlsx`
- Change location and filename as needed
- Will create `.xlsx` extension automatically if not provided

## Troubleshooting

### "Error extracting PDF"
- Verify the PDF is a valid, uncorrupted bank statement
- Try opening the PDF in Adobe Reader to confirm it's readable
- Check that PDF is not password-protected

### "No tables found in PDF"
- The PDF may not contain standard tables
- Some bank statements use image-based PDFs (not text-based)
- Contact support for non-standard formats

### Application crashes on launch
- Ensure Windows has the latest updates
- If running from source, verify Python 3.11+ is installed
- Try running `python src/main.py` from command line to see error details

### Excel file is blank or incomplete
- Check that at least one sheet is selected before conversion
- Verify the output file path is writable
- Ensure sufficient disk space is available

## Data Privacy & Security

- **All processing is local**: No data is sent to external servers
- **No data retention**: PDFs are not stored or cached
- **Secure handling**: Uses industry-standard libraries for PDF and Excel processing

## Support & Updates

For issues or feature requests, contact the development team or review the application logs.

## Version History

### v1.0 (Initial Release)
- Extract all tables from bank statement PDFs
- Interactive preview with sheet selection
- Professional Excel export with formatting
- Standalone Windows executable

---

**Created**: April 2026  
**Python Version**: 3.11+  
**License**: Internal Use Only
