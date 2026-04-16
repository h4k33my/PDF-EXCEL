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
)
from PyQt6.QtGui import QColor


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
