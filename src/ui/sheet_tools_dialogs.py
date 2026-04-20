"""
Dialogs: import columns from another sheet, filter by numeric primary column.
"""
from __future__ import annotations

from typing import List, Literal, Optional, Tuple

from PyQt6.QtWidgets import (
    QDialog,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QComboBox,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QMessageBox,
    QAbstractItemView,
    QDialogButtonBox,
    QLineEdit,
    QFileDialog,
    QRadioButton,
)
from PyQt6.QtGui import QColor
from openpyxl import load_workbook


class ImportColumnsDialog(QDialog):
    def __init__(self, parent, sheets_data: List[dict]):
        super().__init__(parent)
        self.setWindowTitle("Import columns")
        self._sheets = sheets_data
        self._selected_source_cols: List[int] = []
        self._source_idx = 0
        self._dest_idx = 0

        layout = QVBoxLayout(self)

        row1 = QHBoxLayout()
        row1.addWidget(QLabel("Source sheet:"))
        self._source_combo = QComboBox()
        for i, s in enumerate(sheets_data):
            self._source_combo.addItem(s["name"], i)
        self._source_combo.currentIndexChanged.connect(self._on_source_changed)
        row1.addWidget(self._source_combo)
        layout.addLayout(row1)

        self._preview = QTableWidget()
        self._preview.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self._preview.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectColumns)
        self._preview.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self._preview.horizontalHeader().setVisible(True)
        layout.addWidget(QLabel("Select columns (click column headers; Ctrl+click for multiple):"))
        layout.addWidget(self._preview, stretch=1)

        row2 = QHBoxLayout()
        row2.addWidget(QLabel("Destination sheet:"))
        self._dest_combo = QComboBox()
        for i, s in enumerate(sheets_data):
            self._dest_combo.addItem(s["name"], i)
        self._dest_combo.addItem("New sheet (created on OK)", -1)
        row2.addWidget(self._dest_combo)
        layout.addLayout(row2)

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(self._on_ok)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

        self._reload_source_preview()
        self.resize(720, 420)

    def _on_source_changed(self, _index: int):
        self._reload_source_preview()

    def _reload_source_preview(self):
        idx = self._source_combo.currentData()
        if idx is None:
            idx = 0
        self._source_idx = idx
        data = self._sheets[self._source_idx]["data"]
        if not data:
            self._preview.clear()
            self._preview.setRowCount(0)
            self._preview.setColumnCount(0)
            return
        ncols = len(data[0])
        nrows = min(len(data), 500)
        self._preview.setColumnCount(ncols)
        self._preview.setRowCount(nrows)
        hdr_bg = QColor(68, 114, 196)
        hdr_fg = QColor(255, 255, 255)
        for r in range(nrows):
            row = data[r]
            for c in range(ncols):
                val = row[c] if c < len(row) else ""
                item = QTableWidgetItem(str(val))
                if r == 0:
                    item.setBackground(hdr_bg)
                    item.setForeground(hdr_fg)
                self._preview.setItem(r, c, item)
        self._preview.resizeColumnsToContents()

    def _on_ok(self):
        sel = self._preview.selectionModel().selectedColumns()
        cols = sorted({idx.column() for idx in sel})
        if not cols:
            QMessageBox.warning(self, "Import columns", "Select at least one column (use column headers).")
            return
        self._selected_source_cols = cols
        self._dest_idx = self._dest_combo.currentData()
        self.accept()

    def result(self) -> Tuple[int, List[int], int]:
        """source_sheet_idx, column_indices (sorted), dest_sheet_idx or -1 for new sheet."""
        return self._source_idx, list(self._selected_source_cols), self._dest_idx


class PrimaryColumnDialog(QDialog):
    def __init__(self, parent, header_row: List[str]):
        super().__init__(parent)
        self.setWindowTitle("Filter by primary column")
        self._col_idx: Optional[int] = None

        layout = QVBoxLayout(self)
        layout.addWidget(
            QLabel(
                "Choose a column that contains numbers only. "
                "Rows where that column is empty, zero, or non-numeric will be removed (header row kept)."
            )
        )
        self._combo = QComboBox()
        for i, h in enumerate(header_row):
            label = str(h).strip() if h is not None else ""
            if not label:
                label = f"Column {i + 1}"
            self._combo.addItem(label, i)
        layout.addWidget(self._combo)

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(self._accept_ok)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def _accept_ok(self):
        self._col_idx = self._combo.currentData()
        self.accept()

    def selected_column_index(self) -> Optional[int]:
        return self._col_idx


class SheetSelectionDialog(QDialog):
    def __init__(self, parent, sheets_data: List[dict], *, title: str, prompt: str, default_idx: int = 0):
        super().__init__(parent)
        self.setWindowTitle(title)
        self._selected_idx: Optional[int] = None

        layout = QVBoxLayout(self)
        layout.addWidget(QLabel(prompt))
        self._combo = QComboBox()
        for i, s in enumerate(sheets_data):
            name = s.get("name", f"Sheet {i + 1}")
            self._combo.addItem(name, i)
        if 0 <= default_idx < self._combo.count():
            self._combo.setCurrentIndex(default_idx)
        layout.addWidget(self._combo)

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(self._accept_ok)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def _accept_ok(self):
        self._selected_idx = self._combo.currentData()
        self.accept()

    def selected_sheet_index(self) -> Optional[int]:
        return self._selected_idx


class ColumnSelectionDialog(QDialog):
    def __init__(
        self,
        parent,
        header_row: List[object],
        *,
        title: str,
        prompt: str,
        default_idx: int = 0,
    ):
        super().__init__(parent)
        self.setWindowTitle(title)
        self._col_idx: Optional[int] = None

        layout = QVBoxLayout(self)
        layout.addWidget(QLabel(prompt))
        self._combo = QComboBox()
        for i, h in enumerate(header_row):
            label = str(h).strip() if h is not None else ""
            if not label:
                label = f"Column {i + 1}"
            self._combo.addItem(label, i)
        if 0 <= default_idx < self._combo.count():
            self._combo.setCurrentIndex(default_idx)
        layout.addWidget(self._combo)

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(self._accept_ok)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def _accept_ok(self):
        self._col_idx = self._combo.currentData()
        self.accept()

    def selected_column_index(self) -> Optional[int]:
        return self._col_idx


class EventColumnModeDialog(QDialog):
    """Ask whether to insert a new Event column or use an existing sheet column."""

    def __init__(self, parent):
        super().__init__(parent)
        self.setWindowTitle("Event column setup")
        self._choice: Optional[Literal["create", "existing"]] = None

        layout = QVBoxLayout(self)
        layout.addWidget(
            QLabel(
                "How should the app set up the column where you choose an event for each row?"
            )
        )
        layout.addWidget(
            QLabel(
                "• Create new Event column: inserts a column titled “Event” to the right of a column you pick.\n"
                "• Use existing column: pick a column that is already in the sheet (header stays as-is)."
            )
        )
        row = QHBoxLayout()
        create_btn = QPushButton("Create new Event column")
        existing_btn = QPushButton("Use existing column")
        create_btn.clicked.connect(lambda: self._pick("create"))
        existing_btn.clicked.connect(lambda: self._pick("existing"))
        row.addWidget(create_btn)
        row.addWidget(existing_btn)
        layout.addLayout(row)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Cancel)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def _pick(self, choice: Literal["create", "existing"]):
        self._choice = choice
        self.accept()

    def choice(self) -> Optional[Literal["create", "existing"]]:
        return self._choice


class UpdateExistingExcelDialog(QDialog):
    """Configure update mode for writing into an existing workbook."""

    def __init__(self, parent, *, active_sheet_name: str, has_selection: bool):
        super().__init__(parent)
        self.setWindowTitle("Update existing Excel workbook")
        self._result = None

        layout = QVBoxLayout(self)
        layout.addWidget(QLabel("Destination workbook (.xlsx):"))
        path_row = QHBoxLayout()
        self._path_edit = QLineEdit()
        self._path_edit.setPlaceholderText("Choose existing workbook")
        browse_btn = QPushButton("Browse…")
        browse_btn.clicked.connect(self._browse_workbook)
        path_row.addWidget(self._path_edit, 1)
        path_row.addWidget(browse_btn)
        layout.addLayout(path_row)

        layout.addWidget(QLabel("Operation:"))
        self._mode_combo = QComboBox()
        self._mode_combo.addItem("Add selected app sheets as new sheet(s)", "add_sheets")
        self._mode_combo.addItem("Paste into existing sheet", "paste_range")
        self._mode_combo.currentIndexChanged.connect(self._sync_visibility)
        layout.addWidget(self._mode_combo)

        self._paste_block = QVBoxLayout()
        self._sheet_combo = QComboBox()
        self._start_cell_edit = QLineEdit("A1")
        self._start_cell_edit.setPlaceholderText("A1")
        self._source_combo = QComboBox()
        if has_selection:
            self._source_combo.addItem("Use highlighted cells on active sheet", "selection")
        self._source_combo.addItem(f"Use full active sheet ({active_sheet_name})", "full_active")
        self._paste_block.addWidget(QLabel("Destination sheet:"))
        self._paste_block.addWidget(self._sheet_combo)
        self._paste_block.addWidget(QLabel("Start cell (e.g. A2500):"))
        self._paste_block.addWidget(self._start_cell_edit)
        self._paste_block.addWidget(QLabel("Source to paste:"))
        self._paste_block.addWidget(self._source_combo)
        layout.addLayout(self._paste_block)

        self._clear_first = QRadioButton("Clear destination rectangle first, then paste values")
        self._clear_first.setChecked(False)
        layout.addWidget(self._clear_first)

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(self._accept_ok)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)
        self.resize(640, 320)
        self._sync_visibility()

    def _sync_visibility(self):
        paste_mode = self._mode_combo.currentData() == "paste_range"
        for i in range(self._paste_block.count()):
            item = self._paste_block.itemAt(i)
            w = item.widget()
            if w is not None:
                w.setVisible(paste_mode)
        self._clear_first.setVisible(paste_mode)

    def _browse_workbook(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Select destination workbook", "", "Excel Files (*.xlsx)"
        )
        if not path:
            return
        self._path_edit.setText(path)
        self._reload_sheet_names(path)

    def _reload_sheet_names(self, path: str):
        self._sheet_combo.clear()
        try:
            wb = load_workbook(path, read_only=True)
            for name in wb.sheetnames:
                self._sheet_combo.addItem(name, name)
            wb.close()
        except Exception:
            self._sheet_combo.clear()

    def _accept_ok(self):
        path = str(self._path_edit.text() or "").strip()
        if not path:
            QMessageBox.warning(self, "Update workbook", "Choose a destination workbook.")
            return
        mode = self._mode_combo.currentData()
        if mode == "paste_range":
            if self._sheet_combo.count() == 0:
                QMessageBox.warning(self, "Update workbook", "No destination sheets found in workbook.")
                return
            sheet_name = self._sheet_combo.currentData()
            start_cell = str(self._start_cell_edit.text() or "").strip().upper()
            if not start_cell or not any(ch.isdigit() for ch in start_cell):
                QMessageBox.warning(self, "Update workbook", "Enter a valid start cell like A1.")
                return
            source_mode = self._source_combo.currentData()
            self._result = {
                "path": path,
                "mode": mode,
                "sheet_name": sheet_name,
                "start_cell": start_cell,
                "source_mode": source_mode,
                "clear_first": bool(self._clear_first.isChecked()),
            }
        else:
            self._result = {
                "path": path,
                "mode": mode,
            }
        self.accept()

    def result(self) -> Optional[dict]:
        return self._result
