"""
File handling utilities
"""
from pathlib import Path
import os


def get_default_output_path():
    """Get default output path (Desktop)"""
    return str(Path.home() / 'Desktop' / 'bank_statements.xlsx')


def ensure_directory_exists(file_path):
    """Ensure directory exists for given file path"""
    directory = os.path.dirname(file_path)
    if directory and not os.path.exists(directory):
        os.makedirs(directory)


def is_valid_pdf(file_path):
    """Check if file is a valid PDF"""
    if not os.path.exists(file_path):
        return False
    return file_path.lower().endswith('.pdf')


def is_valid_excel_output_path(file_path):
    """Check if output path is valid for Excel"""
    if not file_path:
        return False
    return file_path.lower().endswith('.xlsx') or file_path.lower().endswith('.xls')
