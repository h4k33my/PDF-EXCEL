"""
PyInstaller Build Script
Converts the Python application to a standalone Windows .exe
"""
import PyInstaller.__main__
import os
import shutil

# Get script directory
script_dir = os.path.dirname(os.path.abspath(__file__))

# Clean PyInstaller work folder only (full dist clean can be done manually).
work_root = os.path.join(script_dir, 'build', 'work')
if os.path.exists(work_root):
    shutil.rmtree(work_root)

# Run PyInstaller
PyInstaller.__main__.run([
    os.path.join(script_dir, 'src', 'main.py'),
    '--onefile',                                    # Single executable
    '--windowed',                                   # No console window
    '--name=GAC-PDF-EXCEL-CONVERTER',              # Application name
    '--distpath=' + os.path.join(script_dir, 'dist'),
    '--specpath=' + os.path.join(script_dir, 'build'),
    '--workpath=' + os.path.join(script_dir, 'build', 'work'),
    # pdfplumber pulls in pdfminer and pypdfium2 binaries; openpyxl is pure Python.
    '--collect-all=pdfplumber',
    '--hidden-import=converter',
    '--hidden-import=coordinate_fallback',
    '--hidden-import=utils.excel_loader',
    '--hidden-import=utils.sheet_ops',
    '--hidden-import=ui.sheet_tools_dialogs',
    '--hidden-import=PyQt6.QtCore',
    '--hidden-import=PyQt6.QtGui',
    '--hidden-import=PyQt6.QtWidgets',
])

dist_exe = os.path.join(script_dir, "dist", "GAC-PDF-EXCEL-CONVERTER.exe")
pkg_dir = os.path.join(script_dir, "dist_package")
os.makedirs(pkg_dir, exist_ok=True)
pkg_exe = os.path.join(pkg_dir, "GAC-PDF-EXCEL-CONVERTER.exe")
if os.path.isfile(dist_exe):
    shutil.copy2(dist_exe, pkg_exe)
    print(f"Copied to: {pkg_exe}")

print("\n" + "="*60)
print("Build complete!")
print(f"Executable location: {dist_exe}")
print("="*60)
