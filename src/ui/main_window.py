"""
PyQt6 Main Window for Bank Statement Converter
"""
import os
from pathlib import Path

from PyQt6.QtWidgets import (
    QApplication,
    QAbstractItemView,
    QMainWindow,
    QVBoxLayout,
    QHBoxLayout,
    QPushButton,
    QLabel,
    QLineEdit,
    QTableWidget,
    QTableWidgetItem,
    QFileDialog,
    QWidget,
    QStatusBar,
    QTabWidget,
    QCheckBox,
    QScrollArea,
    QMessageBox,
    QDialog,
    QComboBox,
    QInputDialog,
    QPlainTextEdit,
    QSplitter,
    QMenu,
)
from PyQt6.QtCore import QThread, pyqtSignal, Qt, QTimer
from PyQt6.QtGui import QColor, QKeyEvent, QKeySequence

from ui.sheet_tools_dialogs import (
    ImportColumnsDialog,
    PrimaryColumnDialog,
    SheetSelectionDialog,
    ColumnSelectionDialog,
    EventColumnModeDialog,
    UpdateExistingExcelDialog,
)
from utils.sheet_ops import filter_rows_by_positive_primary, validate_numeric_primary_column
from utils.event_ops import (
    apply_event_amount_mapping,
    clone_grid,
    normalize_header,
    summarize_totals_for_headers,
)


class ConversionWorker(QThread):
    """Run PDF extraction in background to prevent UI freeze"""

    finished = pyqtSignal(list)
    error = pyqtSignal(str)

    def __init__(self, pdf_path):
        super().__init__()
        self.pdf_path = pdf_path

    def run(self):
        try:
            from converter import extract_all_tables_from_pdf

            sheets_data = extract_all_tables_from_pdf(self.pdf_path)
            self.finished.emit(sheets_data)
        except Exception as e:
            self.error.emit(f"Error extracting PDF: {str(e)}")


class EventCellWidget(QWidget):
    """Compact event editor: text field + small dropdown in the corner."""

    textCommitted = pyqtSignal(str)
    optionPicked = pyqtSignal(str)

    def __init__(self, options: list[str], text: str, selected_key: str, parent=None):
        super().__init__(parent)
        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(1)
        self.text_input = QLineEdit(self)
        self.text_input.setFrame(False)
        self.text_input.setStyleSheet("QLineEdit { padding: 0 2px; }")
        self.text_input.setText(str(text or ""))
        self.text_input.editingFinished.connect(self._emit_text_committed)
        self.combo = QComboBox(self)
        self.combo.setFixedWidth(22)
        self.combo.setMaximumHeight(20)
        self.combo.setMaxVisibleItems(18)
        self.combo.setStyleSheet("QComboBox { padding: 0px; }")
        self.combo.setSizeAdjustPolicy(QComboBox.SizeAdjustPolicy.AdjustToContents)
        self.combo.addItem("")
        for option in options:
            self.combo.addItem(option)
        self.combo.setCurrentText(selected_key if selected_key in options else "")
        self._ensure_popup_width()
        self.combo.currentTextChanged.connect(self._emit_option_picked)
        self.setToolTip("Type custom text, or pick mapping item from the corner dropdown.")
        self.combo.setToolTip("Mapping key")
        layout.addWidget(self.text_input, 1)
        layout.addWidget(self.combo, 0)

    def _emit_text_committed(self):
        self.textCommitted.emit(self.text_input.text())

    def _emit_option_picked(self, text: str):
        self.optionPicked.emit(text)

    def sync_state(self, options: list[str], text: str, selected_key: str, combo_enabled: bool):
        if self.text_input.text() != str(text or ""):
            self.text_input.setText(str(text or ""))
        self.combo.blockSignals(True)
        try:
            existing = [self.combo.itemText(i) for i in range(self.combo.count())]
            wanted = [""] + list(options)
            if existing != wanted:
                self.combo.clear()
                for option in wanted:
                    self.combo.addItem(option)
                self._ensure_popup_width()
            wanted_key = selected_key if selected_key in options else ""
            if self.combo.currentText() != wanted_key:
                self.combo.setCurrentText(wanted_key)
            self.combo.setEnabled(combo_enabled)
        finally:
            self.combo.blockSignals(False)

    def _ensure_popup_width(self):
        fm = self.combo.fontMetrics()
        longest = max((len(self.combo.itemText(i)) for i in range(self.combo.count())), default=0)
        popup_width = max(180, fm.horizontalAdvance("W" * min(longest + 4, 80)))
        self.combo.view().setMinimumWidth(popup_width)


def _deep_copy_sheets(sheets_data):
    return [
        {"name": s["name"], "data": [list(row) for row in s["data"]], "is_table": s.get("is_table", True)}
        for s in sheets_data
    ]


def _deep_copy_grid(data):
    return [list(row) for row in data] if data else []


# Snapshot before "Filter by primary column" so user can undo without full session reset.
_PRE_PRIMARY_FILTER_KEY = "_pre_primary_filter_data"


def _copy_sheet_dict(s: dict) -> dict:
    """Deep-copy one sheet entry including optional undo snapshot key."""
    out = {
        "name": s["name"],
        "data": _deep_copy_grid(s.get("data", [])),
        "is_table": s.get("is_table", True),
    }
    if _PRE_PRIMARY_FILTER_KEY in s:
        out[_PRE_PRIMARY_FILTER_KEY] = _deep_copy_grid(s[_PRE_PRIMARY_FILTER_KEY])
    return out


def _copy_all_sheets(sheets_data: list) -> list:
    return [_copy_sheet_dict(s) for s in sheets_data]


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.sheets_data = []
        self.original_sheets_data = []
        self.selected_sheets = {}
        self.preview_tables = []
        self.flow_mode = None  # reserved; classification happens at Done
        self.flow_session_active = False
        self.flow_source_sheet_idx = None
        self.flow_amount_col_idx = None
        self.flow_event_options = []
        self.flow_event_col_idx = None
        self.flow_working_sheet_name = None
        self.flow_working_sheet_idx = None
        self.flow_row_event_keys = {}
        self.flow_prefilled_event_rows = set()
        self.flow_unlocked_event_rows = set()
        self.flow_last_output_sheet_name = None
        self.flow_last_output_sheet_data = None
        self.flow_last_output_amount_col_idx = None
        self._is_rendering_preview = False
        self._undo_stack: list = []
        self._redo_stack: list = []
        self._history_suspended = False
        self._edit_history_pre_snapshot = None
        self._edit_history_timer = QTimer(self)
        self._edit_history_timer.setSingleShot(True)
        self._edit_history_timer.timeout.connect(self._finalize_edit_undo_batch)
        self._max_undo_steps = 50
        self._preview_min_height = 280
        self._bottom_min_height = 260
        self._preview_max_share = 0.70
        self._splitter_adjusting = False
        self._header_move_sync_in_progress = False
        self.initUI()

    def initUI(self):
        """Build main UI layout"""
        self.setWindowTitle("Bank Statement PDF-to-Excel Converter v1.1")
        self.setGeometry(100, 100, 1000, 750)

        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout()

        # ===== FILE INPUT SECTION =====
        input_layout = QHBoxLayout()
        input_layout.addWidget(QLabel("Select PDF File:"))
        self.pdf_path_input = QLineEdit()
        self.pdf_path_input.setReadOnly(True)
        self.pdf_path_input.setPlaceholderText("No file selected")
        input_layout.addWidget(self.pdf_path_input)
        browse_btn = QPushButton("Browse PDF…")
        browse_btn.clicked.connect(self.browse_pdf)
        input_layout.addWidget(browse_btn)

        input_layout.addWidget(QLabel("  Excel:"))
        self.excel_path_input = QLineEdit()
        self.excel_path_input.setReadOnly(True)
        self.excel_path_input.setPlaceholderText("No Excel loaded")
        input_layout.addWidget(self.excel_path_input)
        browse_xlsx_btn = QPushButton("Browse Excel…")
        browse_xlsx_btn.clicked.connect(self.browse_excel)
        input_layout.addWidget(browse_xlsx_btn)
        layout.addLayout(input_layout)

        # ===== PREVIEW / TOOLS SECTION (TOP SPLITTER PANE) =====
        top_panel = QWidget()
        top_layout = QVBoxLayout()
        top_layout.setContentsMargins(0, 0, 0, 0)
        top_layout.addWidget(QLabel("Preview of sheets:"))
        preview_row = QHBoxLayout()
        self.preview_tabs = QTabWidget()
        self.preview_tabs.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.preview_tabs.customContextMenuRequested.connect(self._show_sheet_tab_context_menu)
        self.preview_tabs.tabBarDoubleClicked.connect(self._rename_sheet_from_tab_double_click)
        preview_row.addWidget(self.preview_tabs, stretch=1)
        side_btn_layout = QVBoxLayout()
        side_btn_layout.setContentsMargins(0, 0, 0, 0)
        side_btn_layout.setSpacing(4)
        self._add_sheet_btn = QPushButton("+")
        self._add_sheet_btn.setToolTip("Add blank sheet")
        self._add_sheet_btn.setMinimumWidth(40)
        self._add_sheet_btn.setMaximumWidth(48)
        self._add_sheet_btn.clicked.connect(self.add_blank_sheet)
        side_btn_layout.addWidget(self._add_sheet_btn)
        self.mapping_refresh_button = QPushButton("↻")
        self.mapping_refresh_button.setToolTip("Refresh mapped values")
        self.mapping_refresh_button.setMinimumWidth(40)
        self.mapping_refresh_button.setMaximumWidth(48)
        self.mapping_refresh_button.clicked.connect(self.apply_inflow_outflow_mapping)
        side_btn_layout.addWidget(self.mapping_refresh_button)
        side_btn_layout.addStretch()
        preview_row.addLayout(side_btn_layout, stretch=0)
        top_layout.addLayout(preview_row)

        top_panel.setMinimumHeight(self._preview_min_height)
        top_panel.setLayout(top_layout)

        # ===== LOWER CONTROLS SECTION (BOTTOM SPLITTER PANE) =====
        bottom_panel = QWidget()
        bottom_layout = QVBoxLayout()
        bottom_layout.setContentsMargins(0, 0, 0, 0)
        # Split + sheet tools
        split_layout = QHBoxLayout()
        self.add_split_button = QPushButton("Mark split at selected row")
        self.add_split_button.clicked.connect(self.add_split_point)
        self.clear_split_button = QPushButton("Clear split markers")
        self.clear_split_button.clicked.connect(self.reset_preview_splits)
        self.import_columns_button = QPushButton("Import columns…")
        self.import_columns_button.clicked.connect(self.open_import_columns)
        self.primary_column_button = QPushButton("Filter by primary column…")
        self.primary_column_button.clicked.connect(self.open_primary_column_filter)
        self.clear_primary_filter_button = QPushButton("Clear primary filter")
        self.clear_primary_filter_button.setToolTip(
            "Undo the last primary-column filter on the active sheet (like clearing splits)"
        )
        self.clear_primary_filter_button.clicked.connect(self.clear_primary_filter)
        self.clear_primary_filter_button.setEnabled(False)
        self.split_status_label = QLabel("No split applied")
        split_layout.addWidget(self.add_split_button)
        split_layout.addWidget(self.clear_split_button)
        split_layout.addWidget(self.import_columns_button)
        split_layout.addWidget(self.primary_column_button)
        split_layout.addWidget(self.clear_primary_filter_button)
        split_layout.addWidget(self.split_status_label)
        bottom_layout.addLayout(split_layout)

        # Cash flow mapping (single workflow; classify as Inflow/Outflow when you click Done)
        flow_layout = QHBoxLayout()
        self.start_flow_button = QPushButton("Cash flow mapping…")
        self.start_flow_button.setToolTip("Pick a source sheet, then map amounts to event columns (same steps for credits or debits)")
        self.start_flow_button.clicked.connect(self.start_flow_workflow)
        self.amount_data_button = QPushButton("Amount data")
        self.amount_data_button.clicked.connect(self.choose_amount_data_column)
        self.add_column_button = QPushButton("Add column")
        self.add_column_button.clicked.connect(self.add_flow_header_column)
        self.list_items_button = QPushButton("List items")
        self.list_items_button.clicked.connect(self.capture_list_items_from_header_selection)
        self.events_button = QPushButton("Events")
        self.events_button.clicked.connect(self.setup_events_column)
        self.done_mapping_button = QPushButton("Done")
        self.done_mapping_button.clicked.connect(self.finish_flow_with_total_check)
        self.undo_mapping_button = QPushButton("X")
        self.undo_mapping_button.setToolTip("Remove mapped working sheet")
        self.undo_mapping_button.setMinimumWidth(34)
        self.undo_mapping_button.setMaximumWidth(40)
        self.undo_mapping_button.clicked.connect(self.undo_last_flow_output)
        self.flow_status_label = QLabel("Step 1: Start cash flow mapping")
        flow_layout.addWidget(self.start_flow_button)
        flow_layout.addWidget(self.amount_data_button)
        flow_layout.addWidget(self.add_column_button)
        flow_layout.addWidget(self.list_items_button)
        flow_layout.addWidget(self.events_button)
        flow_layout.addWidget(self.done_mapping_button)
        flow_layout.addWidget(self.undo_mapping_button)
        flow_layout.addWidget(self.flow_status_label)
        bottom_layout.addLayout(flow_layout)
        self.sheets_check_layout = QVBoxLayout()
        sheets_check_widget = QWidget()
        sheets_check_widget.setLayout(self.sheets_check_layout)
        sheets_check_scroll = QScrollArea()
        sheets_check_scroll.setWidget(sheets_check_widget)
        sheets_check_scroll.setWidgetResizable(True)

        bottom_layout.addWidget(QLabel("Select Sheets to Include in Excel:"))
        bottom_layout.addWidget(sheets_check_scroll)

        # ===== FILE OUTPUT SECTION =====
        output_layout = QHBoxLayout()
        output_layout.addWidget(QLabel("Save Excel As:"))
        self.output_path_input = QLineEdit()
        self.output_path_input.setPlaceholderText("bank_statements.xlsx")
        output_layout.addWidget(self.output_path_input)
        browse_out_btn = QPushButton("Browse…")
        browse_out_btn.clicked.connect(self.browse_output)
        output_layout.addWidget(browse_out_btn)
        bottom_layout.addLayout(output_layout)

        # ===== ACTION BUTTONS =====
        button_layout = QHBoxLayout()
        self.undo_btn = QPushButton("Undo")
        self.undo_btn.setToolTip("Undo (Ctrl+Z)")
        self.undo_btn.clicked.connect(self.universal_undo)
        self.redo_btn = QPushButton("Redo")
        self.redo_btn.setToolTip("Redo (Ctrl+Y or Ctrl+Shift+Z)")
        self.redo_btn.clicked.connect(self.universal_redo)
        button_layout.addWidget(self.undo_btn)
        button_layout.addWidget(self.redo_btn)
        convert_btn = QPushButton("Convert & Save")
        convert_btn.setStyleSheet(
            "background-color: #4472C4; color: white; font-weight: bold; padding: 8px;"
        )
        convert_btn.clicked.connect(self.convert_and_save)
        update_existing_btn = QPushButton("Update Existing Excel…")
        update_existing_btn.clicked.connect(self.update_existing_excel)
        exit_btn = QPushButton("Exit")
        exit_btn.clicked.connect(self.close)
        button_layout.addStretch()
        button_layout.addWidget(update_existing_btn)
        button_layout.addWidget(convert_btn)
        button_layout.addWidget(exit_btn)
        bottom_layout.addLayout(button_layout)
        bottom_panel.setMinimumHeight(self._bottom_min_height)
        bottom_panel.setLayout(bottom_layout)

        self._main_splitter = QSplitter(Qt.Orientation.Vertical)
        self._main_splitter.addWidget(top_panel)
        self._main_splitter.addWidget(bottom_panel)
        self._main_splitter.setChildrenCollapsible(False)
        self._main_splitter.setStretchFactor(0, 3)
        self._main_splitter.setStretchFactor(1, 2)
        self._main_splitter.setSizes([480, 270])
        self._main_splitter.splitterMoved.connect(self._on_main_splitter_moved)
        layout.addWidget(self._main_splitter)

        central_widget.setLayout(layout)

        self._update_flow_buttons_state()
        self._update_undo_redo_buttons()

        self.statusBar().showMessage("Ready — load a PDF or Excel file to begin")

    def _allowed_preview_height_bounds(self):
        total = self.height()
        min_preview = self._preview_min_height
        max_preview = int(total * self._preview_max_share)
        max_by_bottom = max(self._preview_min_height, total - self._bottom_min_height)
        max_preview = min(max_preview, max_by_bottom)
        if max_preview < min_preview:
            min_preview = max_preview
        return min_preview, max_preview

    def _enforce_preview_splitter_bounds(self):
        if not hasattr(self, "_main_splitter"):
            return
        sizes = self._main_splitter.sizes()
        if len(sizes) != 2:
            return
        top, _ = sizes
        min_preview, max_preview = self._allowed_preview_height_bounds()
        bounded = max(min_preview, min(top, max_preview))
        if bounded == top:
            return
        self._splitter_adjusting = True
        try:
            self._main_splitter.setSizes([bounded, max(0, self.height() - bounded)])
        finally:
            self._splitter_adjusting = False

    def _on_main_splitter_moved(self, pos, index):
        del pos, index
        if self._splitter_adjusting:
            return
        self._enforce_preview_splitter_bounds()

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self._enforce_preview_splitter_bounds()

    @staticmethod
    def _focus_is_in_table_subtree(widget: QWidget) -> bool:
        w = widget
        while w is not None:
            if isinstance(w, QTableWidget):
                return True
            w = w.parentWidget()
        return False

    def try_consume_undo_redo_key(self, receiver: QWidget, event: QKeyEvent) -> bool:
        """Session undo/redo from keyboard; True if the key event was handled."""
        if not self.isAncestorOf(receiver):
            return False

        aw = QApplication.activeWindow()
        if aw is not None and aw is not self:
            return False
        if isinstance(receiver, (QLineEdit, QPlainTextEdit)) and not self._focus_is_in_table_subtree(receiver):
            return False

        e = event
        if e.matches(QKeySequence.StandardKey.Undo):
            self.universal_undo()
            return True
        if e.matches(QKeySequence.StandardKey.Redo):
            self.universal_redo()
            return True
        mods = e.modifiers()
        ctrl = bool(mods & Qt.KeyboardModifier.ControlModifier)
        shift = bool(mods & Qt.KeyboardModifier.ShiftModifier)
        if ctrl and shift and e.key() == Qt.Key.Key_Z:
            self.universal_redo()
            return True
        if ctrl and not shift and e.key() == Qt.Key.Key_Y:
            self.universal_redo()
            return True
        return False

    def apply_session_sheets(self, sheets_data, status_msg: str):
        """Replace session from extracted PDF or loaded Excel."""
        if not sheets_data:
            self.statusBar().showMessage("Error: No sheets to display")
            return
        self._undo_stack.clear()
        self._redo_stack.clear()
        self._edit_history_pre_snapshot = None
        self._edit_history_timer.stop()
        self._reset_flow_session()
        self.original_sheets_data = _copy_all_sheets(sheets_data)
        self.sheets_data = _copy_all_sheets(sheets_data)
        self.render_preview_and_selection()
        self.split_status_label.setText("No split applied")
        self.statusBar().showMessage(status_msg)
        self._update_undo_redo_buttons()

    def _session_snapshot(self) -> dict:
        return {
            "sheets_data": _copy_all_sheets(self.sheets_data),
            "original_sheets_data": _copy_all_sheets(self.original_sheets_data),
            "selected_sheets": dict(self.selected_sheets),
            "split_status": self.split_status_label.text(),
            "active_tab_index": self.preview_tabs.currentIndex(),
            "flow": {
                "session_active": self.flow_session_active,
                "mode": self.flow_mode,
                "source_sheet_idx": self.flow_source_sheet_idx,
                "amount_col_idx": self.flow_amount_col_idx,
                "event_options": list(self.flow_event_options),
                "event_col_idx": self.flow_event_col_idx,
                "working_sheet_name": self.flow_working_sheet_name,
                "working_sheet_idx": self.flow_working_sheet_idx,
                "row_event_keys": dict(self.flow_row_event_keys),
                "prefilled_event_rows": sorted(self.flow_prefilled_event_rows),
                "unlocked_event_rows": sorted(self.flow_unlocked_event_rows),
                "last_output_sheet_name": self.flow_last_output_sheet_name,
                "last_output_sheet_data": clone_grid(self.flow_last_output_sheet_data)
                if self.flow_last_output_sheet_data
                else None,
                "last_output_amount_col_idx": self.flow_last_output_amount_col_idx,
            },
        }

    def _restore_session(self, snap: dict):
        self.sheets_data = _copy_all_sheets(snap["sheets_data"])
        self.original_sheets_data = _copy_all_sheets(snap["original_sheets_data"])
        self.split_status_label.setText(snap.get("split_status", "No split applied"))
        f = snap.get("flow", {})
        self.flow_session_active = bool(
            f.get("session_active", f.get("mode") is not None)
        )
        self.flow_mode = f.get("mode")
        self.flow_source_sheet_idx = f.get("source_sheet_idx")
        self.flow_amount_col_idx = f.get("amount_col_idx")
        self.flow_event_options = list(f.get("event_options") or [])
        self.flow_event_col_idx = f.get("event_col_idx")
        self.flow_working_sheet_name = f.get("working_sheet_name")
        self.flow_working_sheet_idx = f.get("working_sheet_idx")
        self.flow_row_event_keys = dict(f.get("row_event_keys") or {})
        self.flow_prefilled_event_rows = set(f.get("prefilled_event_rows") or [])
        self.flow_unlocked_event_rows = set(f.get("unlocked_event_rows") or [])
        self.flow_last_output_sheet_name = f.get("last_output_sheet_name")
        lod = f.get("last_output_sheet_data")
        self.flow_last_output_sheet_data = clone_grid(lod) if lod else None
        self.flow_last_output_amount_col_idx = f.get("last_output_amount_col_idx")

    def _trim_undo_stack(self):
        while len(self._undo_stack) > self._max_undo_steps:
            self._undo_stack.pop(0)

    def _finalize_edit_undo_batch(self):
        if self._history_suspended or self._edit_history_pre_snapshot is None:
            self._edit_history_pre_snapshot = None
            return
        self._undo_stack.append(self._edit_history_pre_snapshot)
        self._redo_stack.clear()
        self._edit_history_pre_snapshot = None
        self._trim_undo_stack()
        self._update_undo_redo_buttons()

    def _mark_cell_edit_for_undo_batch(self):
        if self._history_suspended:
            return
        if self._edit_history_pre_snapshot is None:
            self._edit_history_pre_snapshot = self._session_snapshot()
        self._edit_history_timer.start(450)

    def _push_history_before_change(self):
        if self._history_suspended:
            return
        self._edit_history_timer.stop()
        if self._edit_history_pre_snapshot is not None:
            self._undo_stack.append(self._edit_history_pre_snapshot)
            self._edit_history_pre_snapshot = None
        self._undo_stack.append(self._session_snapshot())
        self._redo_stack.clear()
        self._trim_undo_stack()
        self._update_undo_redo_buttons()

    def universal_undo(self):
        if self._history_suspended:
            return
        self._edit_history_timer.stop()
        if self._edit_history_pre_snapshot is not None:
            self._history_suspended = True
            try:
                current = self._session_snapshot()
                prev = self._edit_history_pre_snapshot
                self._edit_history_pre_snapshot = None
                self._redo_stack.append(current)
                self._restore_session(prev)
                preserved = prev.get("selected_sheets")
                tab_idx = prev.get("active_tab_index", 0)
                self.render_preview_and_selection(preserved_selection=preserved)
                if 0 <= tab_idx < len(self.sheets_data):
                    self.preview_tabs.setCurrentIndex(tab_idx)
                self._update_clear_primary_filter_button_state()
                self._update_flow_buttons_state()
                self.statusBar().showMessage("Undo.")
            finally:
                self._history_suspended = False
            self._update_undo_redo_buttons()
            return

        if not self._undo_stack:
            self.statusBar().showMessage("Nothing to undo.")
            self._update_undo_redo_buttons()
            return
        self._history_suspended = True
        try:
            current = self._session_snapshot()
            previous = self._undo_stack.pop()
            self._redo_stack.append(current)
            self._restore_session(previous)
            preserved = previous.get("selected_sheets")
            tab_idx = previous.get("active_tab_index", 0)
            self.render_preview_and_selection(preserved_selection=preserved)
            if 0 <= tab_idx < len(self.sheets_data):
                self.preview_tabs.setCurrentIndex(tab_idx)
            self._update_clear_primary_filter_button_state()
            self._update_flow_buttons_state()
            self.statusBar().showMessage("Undo.")
        finally:
            self._history_suspended = False
        self._update_undo_redo_buttons()

    def universal_redo(self):
        if self._history_suspended:
            return
        if not self._redo_stack:
            self.statusBar().showMessage("Nothing to redo.")
            self._update_undo_redo_buttons()
            return
        self._edit_history_timer.stop()
        self._edit_history_pre_snapshot = None
        self._history_suspended = True
        try:
            current = self._session_snapshot()
            nxt = self._redo_stack.pop()
            self._undo_stack.append(current)
            self._restore_session(nxt)
            preserved = nxt.get("selected_sheets")
            tab_idx = nxt.get("active_tab_index", 0)
            self.render_preview_and_selection(preserved_selection=preserved)
            if 0 <= tab_idx < len(self.sheets_data):
                self.preview_tabs.setCurrentIndex(tab_idx)
            self._update_clear_primary_filter_button_state()
            self._update_flow_buttons_state()
            self.statusBar().showMessage("Redo.")
        finally:
            self._history_suspended = False
        self._update_undo_redo_buttons()

    def _update_undo_redo_buttons(self):
        self.undo_btn.setEnabled(bool(self._undo_stack) or self._edit_history_pre_snapshot is not None)
        self.redo_btn.setEnabled(bool(self._redo_stack))

    def _reset_flow_session(self):
        self.flow_mode = None
        self.flow_session_active = False
        self.flow_source_sheet_idx = None
        self.flow_amount_col_idx = None
        self.flow_event_options = []
        self.flow_event_col_idx = None
        self.flow_working_sheet_name = None
        self.flow_working_sheet_idx = None
        self.flow_row_event_keys = {}
        self.flow_prefilled_event_rows = set()
        self.flow_unlocked_event_rows = set()
        self.flow_last_output_sheet_name = None
        self.flow_last_output_sheet_data = None
        self.flow_last_output_amount_col_idx = None
        self._update_flow_buttons_state()

    def _update_flow_buttons_state(self):
        mode_active = self.flow_session_active and self.flow_source_sheet_idx is not None
        self.amount_data_button.setEnabled(mode_active)
        self.add_column_button.setEnabled(mode_active)
        self.list_items_button.setEnabled(mode_active and self.flow_amount_col_idx is not None)
        self.events_button.setEnabled(
            mode_active and self.flow_amount_col_idx is not None and bool(self.flow_event_options)
        )
        self.mapping_refresh_button.setEnabled(
            mode_active
            and self._resolve_flow_working_sheet_idx() is not None
            and self.flow_amount_col_idx is not None
            and self.flow_event_col_idx is not None
            and bool(self.flow_event_options)
        )
        self.undo_mapping_button.setEnabled(bool(self.flow_last_output_sheet_name))
        self.done_mapping_button.setEnabled(bool(self.flow_last_output_sheet_name))
        if not mode_active:
            self.flow_status_label.setText("Step 1: Start cash flow mapping")

    def start_flow_workflow(self):
        if not self.sheets_data:
            self.statusBar().showMessage("Load a PDF or Excel file first.")
            return
        default_idx = self.preview_tabs.currentIndex() if self.preview_tabs.currentIndex() >= 0 else 0
        dlg = SheetSelectionDialog(
            self,
            self.sheets_data,
            title="Cash flow mapping",
            prompt="Select the working source sheet (same workflow for inflows or outflows):",
            default_idx=default_idx,
        )
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return
        src_idx = dlg.selected_sheet_index()
        if src_idx is None:
            return
        self._push_history_before_change()
        self.flow_session_active = True
        self.flow_mode = None
        self.flow_source_sheet_idx = src_idx
        self.flow_amount_col_idx = None
        self.flow_event_options = []
        self.flow_event_col_idx = None
        self.flow_working_sheet_name = None
        self.flow_working_sheet_idx = None
        self.flow_row_event_keys = {}
        self.flow_prefilled_event_rows = set()
        self.flow_unlocked_event_rows = set()
        self.flow_last_output_sheet_name = None
        self.flow_last_output_sheet_data = None
        self.flow_last_output_amount_col_idx = None
        self.flow_status_label.setText(
            f"Step 2: choose Amount data column on “{self.sheets_data[src_idx]['name']}”"
        )
        self.preview_tabs.setCurrentIndex(src_idx)
        self._update_flow_buttons_state()
        self.statusBar().showMessage(f"Cash flow mapping started for sheet “{self.sheets_data[src_idx]['name']}”.")

    def _current_source_data(self):
        if self.flow_source_sheet_idx is None:
            return None
        if self.flow_source_sheet_idx < 0 or self.flow_source_sheet_idx >= len(self.sheets_data):
            return None
        return self.sheets_data[self.flow_source_sheet_idx]["data"]

    def _resolve_flow_working_sheet_idx(self):
        if self.flow_working_sheet_name:
            for idx, sheet in enumerate(self.sheets_data):
                if sheet.get("name") == self.flow_working_sheet_name:
                    self.flow_working_sheet_idx = idx
                    return idx
        idx = self.flow_working_sheet_idx
        if idx is not None and 0 <= idx < len(self.sheets_data):
            self.flow_working_sheet_name = self.sheets_data[idx].get("name")
            return idx
        return None

    def _current_working_data(self):
        idx = self._resolve_flow_working_sheet_idx()
        if idx is None:
            return None
        return self.sheets_data[idx]["data"]

    def choose_amount_data_column(self):
        data = self._current_source_data()
        if data is None:
            self.statusBar().showMessage("Start cash flow mapping first.")
            return
        if not data:
            QMessageBox.warning(self, "Amount data", "Selected source sheet is empty.")
            return
        header = data[0]
        default_idx = self.flow_amount_col_idx if self.flow_amount_col_idx is not None else 0
        dlg = ColumnSelectionDialog(
            self,
            header,
            title="Amount data column",
            prompt="Select the column that contains transaction amounts:",
            default_idx=default_idx,
        )
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return
        col_idx = dlg.selected_column_index()
        if col_idx is None:
            return
        self._push_history_before_change()
        self.flow_amount_col_idx = col_idx
        label = header[col_idx] if col_idx < len(header) else f"Column {col_idx + 1}"
        self.flow_status_label.setText(f"Step 3: Select header cell(s), then click List items (amount: {label})")
        self._update_flow_buttons_state()
        self.statusBar().showMessage(f"Amount data set to column “{label}”.")

    def capture_list_items_from_header_selection(self):
        if self.flow_source_sheet_idx is None:
            self.statusBar().showMessage("Start cash flow mapping first.")
            return
        if self.preview_tabs.currentIndex() != self.flow_source_sheet_idx:
            self.preview_tabs.setCurrentIndex(self.flow_source_sheet_idx)
            self.statusBar().showMessage("Switched to source sheet. Select header cells and click again.")
            return
        if self.flow_source_sheet_idx >= len(self.preview_tables):
            return
        table = self.preview_tables[self.flow_source_sheet_idx]
        selected = table.selectionModel().selectedIndexes() if table.selectionModel() else []
        selected_cols = sorted({idx.column() for idx in selected if idx.row() == 0})
        if not selected_cols:
            QMessageBox.information(
                self,
                "List items",
                "Select one or more header cells (row 1) on the source sheet, then click List items.",
            )
            return
        data = self.sheets_data[self.flow_source_sheet_idx]["data"]
        header = data[0] if data else []
        options = []
        for c in selected_cols:
            value = str(header[c]).strip() if c < len(header) else ""
            if value and value not in options:
                options.append(value)
        if not options:
            QMessageBox.warning(self, "List items", "No valid non-empty header values were selected.")
            return
        previous = list(self.flow_event_options)
        merged = list(previous)
        added = 0
        for opt in options:
            if opt not in merged:
                merged.append(opt)
                added += 1
        if merged == previous:
            self.statusBar().showMessage(
                f"No new list items added (still {len(merged)} total option(s))."
            )
            return
        self._push_history_before_change()
        self.flow_event_options = merged
        if self.flow_event_col_idx is not None:
            self._refresh_event_widgets_on_working_sheet()
            self._recompute_flow_mapping_on_working_sheet(announce=True)
        self.flow_status_label.setText(
            f"Step 4: Click Events to add/select Event column ({len(merged)} total option(s))"
        )
        self._update_flow_buttons_state()
        self.statusBar().showMessage(
            f"List items updated: +{added} new, {len(merged)} total option(s)."
        )

    def add_flow_header_column(self):
        data = self._current_source_data()
        if data is None:
            self.statusBar().showMessage("Start cash flow mapping first.")
            return
        if not data:
            QMessageBox.warning(self, "Add column", "Source sheet is empty.")
            return
        header_name, ok = QInputDialog.getText(
            self,
            "Add column header",
            "Enter the new header name:",
        )
        if not ok:
            return
        header_name = str(header_name or "").strip()
        if not header_name:
            QMessageBox.warning(self, "Add column", "Header name cannot be empty.")
            return
        self._push_history_before_change()
        for row_idx, row in enumerate(data):
            row.append(header_name if row_idx == 0 else "")
        self.render_preview_and_selection()
        if self.flow_source_sheet_idx is not None:
            self.preview_tabs.setCurrentIndex(self.flow_source_sheet_idx)
        self.statusBar().showMessage(f"Added new column “{header_name}” to source sheet.")

    def setup_events_column(self):
        data = self._current_source_data()
        if data is None:
            self.statusBar().showMessage("Start cash flow mapping first.")
            return
        if not data:
            QMessageBox.warning(self, "Events", "Source sheet is empty.")
            return
        header = data[0]
        mode_dlg = EventColumnModeDialog(self)
        if mode_dlg.exec() != QDialog.DialogCode.Accepted:
            return
        mode = mode_dlg.choice()
        if mode is None:
            return
        event_col_idx = None
        if mode == "create":
            anchor_dlg = ColumnSelectionDialog(
                self,
                header,
                title="New Event column position",
                prompt="Select a column; a new “Event” column will be inserted immediately to its right:",
                default_idx=0,
            )
            if anchor_dlg.exec() != QDialog.DialogCode.Accepted:
                return
            anchor_idx = anchor_dlg.selected_column_index()
            if anchor_idx is None:
                return
            event_col_idx = anchor_idx + 1
        else:
            pick_dlg = ColumnSelectionDialog(
                self,
                header,
                title="Event column",
                prompt="Select the existing column to use for choosing an event on each row:",
                default_idx=0,
            )
            if pick_dlg.exec() != QDialog.DialogCode.Accepted:
                return
            event_col_idx = pick_dlg.selected_column_index()
            if event_col_idx is None:
                return
        self._push_history_before_change()
        if mode == "create":
            for r, row in enumerate(data):
                while len(row) < event_col_idx:
                    row.append("")
                row.insert(event_col_idx, "Event" if r == 0 else "")
        self.flow_event_col_idx = event_col_idx
        self.flow_row_event_keys = {}
        self.flow_prefilled_event_rows = set()
        self.flow_unlocked_event_rows = set()
        for row_idx in range(1, len(data)):
            row = data[row_idx]
            event_text = str(row[event_col_idx] if event_col_idx < len(row) else "").strip()
            if event_text:
                self.flow_prefilled_event_rows.add(row_idx)
                continue
            # Empty event cells are ready for dropdown-based mapping.
            self.flow_row_event_keys.pop(row_idx, None)

        old_work_idx = self._resolve_flow_working_sheet_idx()
        if old_work_idx is not None and 0 <= old_work_idx < len(self.sheets_data):
            self.sheets_data.pop(old_work_idx)
            self.flow_working_sheet_idx = None
            self.flow_working_sheet_name = None
            self.flow_last_output_sheet_name = None
            self.flow_last_output_sheet_data = None
        used_names = {s["name"] for s in self.sheets_data}
        output_name = self.make_unique_sheet_name("Mapped_cashflow_working", used_names)
        out_sheet = {"name": output_name, "data": clone_grid(data), "is_table": True}
        self.sheets_data.append(out_sheet)
        self.flow_working_sheet_idx = len(self.sheets_data) - 1
        self.flow_working_sheet_name = output_name
        self.flow_last_output_sheet_name = output_name
        self.flow_last_output_amount_col_idx = self.flow_amount_col_idx
        self._recompute_flow_mapping_on_working_sheet(render=True, announce=False)
        working_idx = self._resolve_flow_working_sheet_idx()
        if working_idx is not None:
            self.preview_tabs.setCurrentIndex(working_idx)
        self.flow_status_label.setText("Step 5: Edit Event cells or use dropdown; mapping updates in real time")
        self._update_flow_buttons_state()
        self.statusBar().showMessage("Events column is ready. Working mapped sheet created with real-time updates.")

    def _recompute_flow_mapping_on_working_sheet(self, render: bool = False, announce: bool = False):
        data = self._current_working_data()
        if data is None:
            return
        if self.flow_amount_col_idx is None or self.flow_event_col_idx is None:
            return
        mapped_data, stats, _created = apply_event_amount_mapping(
            clone_grid(data),
            amount_col_idx=self.flow_amount_col_idx,
            event_col_idx=self.flow_event_col_idx,
            options=self.flow_event_options,
            row_event_keys=self.flow_row_event_keys,
        )
        idx = self._resolve_flow_working_sheet_idx()
        if idx is None:
            return
        self.sheets_data[idx]["data"] = mapped_data
        self.flow_last_output_sheet_data = clone_grid(mapped_data)
        if render:
            self.render_preview_and_selection()
            idx = self._resolve_flow_working_sheet_idx()
            if idx is not None:
                self.preview_tabs.setCurrentIndex(idx)
        else:
            self._refresh_working_sheet_table_view()
        if announce:
            self.statusBar().showMessage(
                f"Mapped rows updated: {stats['rows_updated']} updated, {stats['rows_skipped']} skipped."
            )

    def _refresh_working_sheet_table_view(self):
        idx = self._resolve_flow_working_sheet_idx()
        if idx is None or idx >= len(self.preview_tables):
            return
        if idx >= len(self.sheets_data):
            return
        data = self.sheets_data[idx]["data"]
        table = self.preview_tables[idx]
        self._is_rendering_preview = True
        try:
            ncols = max((len(r) for r in data), default=0)
            table.setColumnCount(ncols)
            table.setRowCount(len(data))
            for row_idx, row in enumerate(data):
                for col_idx in range(ncols):
                    if row_idx > 0 and col_idx == self.flow_event_col_idx:
                        continue
                    cell_value = row[col_idx] if col_idx < len(row) else ""
                    item = table.item(row_idx, col_idx)
                    if item is None:
                        item = QTableWidgetItem(str(cell_value))
                        table.setItem(row_idx, col_idx, item)
                    else:
                        item.setText(str(cell_value))
                    if row_idx == 0:
                        item.setBackground(QColor(68, 114, 196))
                        item.setForeground(QColor(255, 255, 255))
            table.resizeColumnsToContents()
        finally:
            self._is_rendering_preview = False
        self._refresh_event_widgets_on_working_sheet()

    def apply_inflow_outflow_mapping(self):
        data = self._current_working_data()
        if data is None:
            self.statusBar().showMessage("Set up Events first to create the working mapped sheet.")
            return
        if self.flow_amount_col_idx is None or self.flow_event_col_idx is None:
            self.statusBar().showMessage("Set amount data and events column first.")
            return
        self._push_history_before_change()
        self._recompute_flow_mapping_on_working_sheet(render=True, announce=True)
        idx = self._resolve_flow_working_sheet_idx()
        if idx is not None:
            self.preview_tabs.setCurrentIndex(idx)
        self._update_flow_buttons_state()
        self.statusBar().showMessage("Mapping recomputed from current Event selections.")

    def undo_last_flow_output(self):
        if not self.flow_last_output_sheet_name:
            self.statusBar().showMessage("No recent mapped sheet to remove.")
            return
        target = self.flow_last_output_sheet_name
        remove_idx = None
        for i, sheet in enumerate(self.sheets_data):
            if sheet.get("name") == target:
                remove_idx = i
                break
        if remove_idx is None:
            self.flow_working_sheet_idx = None
            self.flow_working_sheet_name = None
            self.flow_row_event_keys = {}
            self.flow_prefilled_event_rows = set()
            self.flow_unlocked_event_rows = set()
            self.flow_last_output_sheet_name = None
            self.flow_last_output_sheet_data = None
            self.flow_last_output_amount_col_idx = None
            self._update_flow_buttons_state()
            self.statusBar().showMessage("Last mapped sheet was not found.")
            return
        self._push_history_before_change()
        self.sheets_data.pop(remove_idx)
        self.flow_working_sheet_idx = None
        self.flow_working_sheet_name = None
        self.flow_row_event_keys = {}
        self.flow_prefilled_event_rows = set()
        self.flow_unlocked_event_rows = set()
        self.flow_last_output_sheet_name = None
        self.flow_last_output_sheet_data = None
        self.flow_last_output_amount_col_idx = None
        self.render_preview_and_selection()
        if self.flow_source_sheet_idx is not None and self.flow_source_sheet_idx < len(self.sheets_data):
            self.preview_tabs.setCurrentIndex(self.flow_source_sheet_idx)
        self._update_flow_buttons_state()
        self.statusBar().showMessage("Last mapped output sheet removed.")

    def _rename_last_mapped_sheet(self, base_name: str) -> bool:
        """Rename the last mapped output sheet to Inflow/Outflow (unique). Returns False if not found."""
        old = self.flow_last_output_sheet_name
        if not old:
            return False
        target = None
        for s in self.sheets_data:
            if s.get("name") == old:
                target = s
                break
        if target is None:
            return False
        used = {s["name"] for s in self.sheets_data if s is not target}
        new_name = self.make_unique_sheet_name(base_name, used)
        target["name"] = new_name
        self.flow_last_output_sheet_name = new_name
        return True

    def finish_flow_with_total_check(self):
        if not self.flow_last_output_sheet_name:
            self.statusBar().showMessage("Set up Events first to create a mapped working sheet.")
            return
        sheet = None
        for s in self.sheets_data:
            if s.get("name") == self.flow_last_output_sheet_name:
                sheet = s
                break
        if sheet is None:
            QMessageBox.warning(self, "Done", "Mapped working sheet was not found. Recreate Events setup.")
            return

        classify = QMessageBox(self)
        classify.setWindowTitle("Inflow or outflow?")
        classify.setText(
            "Was this mapped sheet for inflows or outflows?\n\n"
            "The sheet tab will be renamed to match your choice, then totals will be checked."
        )
        in_btn = classify.addButton("Inflow", QMessageBox.ButtonRole.ActionRole)
        out_btn = classify.addButton("Outflow", QMessageBox.ButtonRole.ActionRole)
        classify.addButton(QMessageBox.StandardButton.Cancel)
        classify.exec()
        clicked = classify.clickedButton()
        if clicked == in_btn:
            flow_label = "Inflow"
        elif clicked == out_btn:
            flow_label = "Outflow"
        else:
            self.statusBar().showMessage("Done cancelled — sheet name unchanged.")
            return

        self._push_history_before_change()
        if not self._rename_last_mapped_sheet(flow_label):
            QMessageBox.warning(self, "Done", "Could not rename mapped sheet.")
            return
        sheet = None
        for s in self.sheets_data:
            if s.get("name") == self.flow_last_output_sheet_name:
                sheet = s
                break
        if sheet is None:
            QMessageBox.warning(self, "Done", "Mapped sheet was not found after rename.")
            return

        preserved = dict(self.selected_sheets)
        self.render_preview_and_selection(preserved_selection=preserved)
        for i, sh in enumerate(self.sheets_data):
            if sh.get("name") == self.flow_last_output_sheet_name:
                self.preview_tabs.setCurrentIndex(i)
                break
        self._update_flow_buttons_state()

        amount_col_idx = self.flow_last_output_amount_col_idx
        if amount_col_idx is None:
            QMessageBox.warning(self, "Done", "Amount column is not available for total check.")
            return
        amount_total, per_header, mapped_total = summarize_totals_for_headers(
            sheet.get("data", []),
            amount_col_idx=amount_col_idx,
            headers=self.flow_event_options,
        )
        diff = amount_total - mapped_total
        if abs(diff) > 0.005:
            totals_block = (
                f"Amount column total: {amount_total:,.2f}\n"
                f"Mapped headers total: {mapped_total:,.2f}\n"
                f"Difference: {diff:,.2f}"
            )
            msg = QMessageBox(self)
            msg.setIcon(QMessageBox.Icon.Warning)
            msg.setWindowTitle("Totals mismatch")
            msg.setText(
                "Amount total does not match mapped header totals.\n"
                f"{totals_block}\n\n"
                "Review event assignments and mapping before saving."
            )
            msg.addButton("OK", QMessageBox.ButtonRole.AcceptRole)
            msg.exec()
            self.statusBar().showMessage("Totals mismatch — review event assignments before saving.")
            return
        QMessageBox.information(
            self,
            "Totals verified",
            "Totals match.\n"
            f"Amount column total: {amount_total:,.2f}\n"
            f"Mapped headers total: {mapped_total:,.2f}",
        )
        self.statusBar().showMessage("Totals verified — you can continue editing or Convert & Save when ready.")
        return

    def browse_pdf(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Select Bank Statement PDF", "", "PDF Files (*.pdf)"
        )
        if file_path:
            self.pdf_path_input.setText(file_path)
            self.excel_path_input.clear()
            self.extract_preview(file_path)

    def browse_excel(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Select Excel workbook", "", "Excel Files (*.xlsx)"
        )
        if not file_path:
            return
        try:
            from utils.excel_loader import load_xlsx_to_sheets_data
        except ImportError:
            from excel_loader import load_xlsx_to_sheets_data

        try:
            sheets_data = load_xlsx_to_sheets_data(file_path)
        except Exception as e:
            QMessageBox.critical(self, "Excel load error", str(e))
            return
        if not sheets_data:
            QMessageBox.warning(self, "Excel load", "No worksheets found in the file.")
            return
        self.excel_path_input.setText(file_path)
        self.pdf_path_input.clear()
        self.apply_session_sheets(
            sheets_data,
            f"✓ Loaded {len(sheets_data)} sheet(s) from Excel — preview only, no conversion",
        )

    def extract_preview(self, pdf_path):
        self.statusBar().showMessage("Processing PDF…")
        self.worker = ConversionWorker(pdf_path)
        self.worker.finished.connect(self.on_extract_finished)
        self.worker.error.connect(self.on_extract_error)
        self.worker.start()

    def on_extract_finished(self, sheets_data):
        if not sheets_data:
            self.statusBar().showMessage("Error: No tables found in PDF")
            return
        self.excel_path_input.clear()
        self.apply_session_sheets(
            sheets_data,
            f"✓ Extracted {len(sheets_data)} sheet(s) — split rows, import columns, or save when ready",
        )

    def on_sheet_toggle(self, sheet_idx, state):
        new_val = state == Qt.CheckState.Checked.value
        if (
            not self._is_rendering_preview
            and not self._history_suspended
            and self.selected_sheets.get(sheet_idx, True) != new_val
        ):
            self._push_history_before_change()
        self.selected_sheets[sheet_idx] = new_val

    def on_extract_error(self, error_msg):
        self.statusBar().showMessage(f"Error: {error_msg}")

    def _next_unique_sheet_name(self) -> str:
        used = {s["name"] for s in self.sheets_data}
        base = "Sheet"
        name = base
        n = 2
        while name in used:
            name = f"{base}_{n}"
            n += 1
        return name

    def add_blank_sheet(self):
        if not self.sheets_data:
            self.statusBar().showMessage("Load a PDF or Excel file first.")
            return
        self._push_history_before_change()
        name = self._next_unique_sheet_name()
        self.sheets_data.append({"name": name, "data": [[""]], "is_table": True})
        self.render_preview_and_selection()
        self.preview_tabs.setCurrentIndex(len(self.sheets_data) - 1)
        self.statusBar().showMessage(f"Added blank sheet “{name}”.")

    def open_import_columns(self):
        if not self.sheets_data:
            self.statusBar().showMessage("No sheets loaded.")
            return
        dlg = ImportColumnsDialog(self, self.sheets_data)
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return
        src_idx, col_indices, dest_idx = dlg.result()
        if not col_indices:
            return
        src = self.sheets_data[src_idx]["data"]
        if not src:
            QMessageBox.warning(self, "Import columns", "Source sheet is empty.")
            return
        header = [src[0][i] if i < len(src[0]) else "" for i in col_indices]
        new_rows = [header]
        for r in range(1, len(src)):
            row = src[r]
            new_rows.append([row[i] if i < len(row) else "" for i in col_indices])
        self._push_history_before_change()
        if dest_idx < 0:
            dest_name = self._next_unique_sheet_name()
            self.sheets_data.append({"name": dest_name, "data": new_rows, "is_table": True})
        else:
            dest = self.sheets_data[dest_idx]
            dest.pop(_PRE_PRIMARY_FILTER_KEY, None)
            dest["data"] = new_rows
        self._reset_flow_session()
        self.render_preview_and_selection()
        self.statusBar().showMessage(
            f"Imported {len(col_indices)} column(s) "
            f"from “{self.sheets_data[src_idx]['name']}”."
        )

    def open_primary_column_filter(self):
        tab_idx = self.preview_tabs.currentIndex()
        if tab_idx < 0 or tab_idx >= len(self.sheets_data):
            self.statusBar().showMessage("Select a sheet tab first.")
            return
        data = self.sheets_data[tab_idx]["data"]
        if not data:
            QMessageBox.information(self, "Filter", "This sheet has no data.")
            return
        header = data[0]
        dlg = PrimaryColumnDialog(self, header)
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return
        col_idx = dlg.selected_column_index()
        if col_idx is None:
            return
        if not validate_numeric_primary_column(data, col_idx):
            QMessageBox.warning(
                self,
                "Filter by primary column",
                "This column must contain only numeric values in data rows "
                "(empty cells are OK). Choose a column with amounts, not text such as Reference.",
            )
            return
        self._push_history_before_change()
        sheet = self.sheets_data[tab_idx]
        sheet[_PRE_PRIMARY_FILTER_KEY] = _deep_copy_grid(sheet["data"])
        filtered = filter_rows_by_positive_primary(data, col_idx)
        sheet["data"] = filtered
        self.render_preview_and_selection()
        self.preview_tabs.setCurrentIndex(tab_idx)
        self.statusBar().showMessage(
            f"Filtered sheet to rows with value > 0 in column “{header[col_idx] if col_idx < len(header) else col_idx + 1}”. "
            "Use “Clear primary filter” to undo."
        )

    def clear_primary_filter(self):
        tab_idx = self.preview_tabs.currentIndex()
        if tab_idx < 0 or tab_idx >= len(self.sheets_data):
            self.statusBar().showMessage("Select a sheet tab first.")
            return
        sheet = self.sheets_data[tab_idx]
        snap = sheet.get(_PRE_PRIMARY_FILTER_KEY)
        if not snap:
            QMessageBox.information(
                self,
                "Clear primary filter",
                "This sheet has no primary-column filter to undo on the active tab.",
            )
            return
        self._push_history_before_change()
        sheet.pop(_PRE_PRIMARY_FILTER_KEY, None)
        sheet["data"] = _deep_copy_grid(snap)
        self.render_preview_and_selection()
        self.preview_tabs.setCurrentIndex(tab_idx)
        self.statusBar().showMessage("Restored sheet to before the last primary-column filter.")

    def _update_clear_primary_filter_button_state(self):
        idx = self.preview_tabs.currentIndex()
        if idx < 0 or idx >= len(self.sheets_data):
            self.clear_primary_filter_button.setEnabled(False)
            return
        self.clear_primary_filter_button.setEnabled(
            _PRE_PRIMARY_FILTER_KEY in self.sheets_data[idx]
        )

    def _on_preview_tab_changed(self, _index: int):
        self._update_clear_primary_filter_button_state()
        self._update_flow_buttons_state()

    def _insert_blank_sheet_at(self, index: int):
        name = self._next_unique_sheet_name()
        idx = max(0, min(index, len(self.sheets_data)))
        self.sheets_data.insert(idx, {"name": name, "data": [[""]], "is_table": True})
        return idx, name

    def _show_sheet_tab_context_menu(self, pos):
        tab_bar = self.preview_tabs.tabBar()
        tab_idx = tab_bar.tabAt(pos)
        if tab_idx < 0:
            return
        menu = QMenu(self)
        act_insert_before = menu.addAction("Insert blank sheet before")
        act_insert_after = menu.addAction("Insert blank sheet after")
        act_duplicate = menu.addAction("Duplicate sheet")
        act_rename = menu.addAction("Rename sheet")
        act_delete = menu.addAction("Delete sheet")
        chosen = menu.exec(tab_bar.mapToGlobal(pos))
        if chosen is None:
            return
        if chosen == act_insert_before:
            self._push_history_before_change()
            new_idx, name = self._insert_blank_sheet_at(tab_idx)
            self.render_preview_and_selection()
            self.preview_tabs.setCurrentIndex(new_idx)
            self.statusBar().showMessage(f"Inserted blank sheet “{name}” before current sheet.")
            return
        if chosen == act_insert_after:
            self._push_history_before_change()
            new_idx, name = self._insert_blank_sheet_at(tab_idx + 1)
            self.render_preview_and_selection()
            self.preview_tabs.setCurrentIndex(new_idx)
            self.statusBar().showMessage(f"Inserted blank sheet “{name}” after current sheet.")
            return
        if chosen == act_duplicate:
            if tab_idx >= len(self.sheets_data):
                return
            self._push_history_before_change()
            original = self.sheets_data[tab_idx]
            copy_name = self.make_unique_sheet_name(f"{original['name']}_copy", {s["name"] for s in self.sheets_data})
            self.sheets_data.insert(
                tab_idx + 1,
                {"name": copy_name, "data": _deep_copy_grid(original.get("data", [])), "is_table": True},
            )
            self.render_preview_and_selection()
            self.preview_tabs.setCurrentIndex(tab_idx + 1)
            self.statusBar().showMessage(f"Duplicated sheet as “{copy_name}”.")
            return
        if chosen == act_rename:
            if tab_idx >= len(self.sheets_data):
                return
            old_name = self.sheets_data[tab_idx]["name"]
            new_name, ok = QInputDialog.getText(self, "Rename sheet", "New sheet name:", text=old_name)
            if not ok:
                return
            new_name = str(new_name or "").strip()
            if not new_name or new_name == old_name:
                return
            used = {s["name"] for i, s in enumerate(self.sheets_data) if i != tab_idx}
            final_name = self.make_unique_sheet_name(new_name, used)
            self._push_history_before_change()
            self.sheets_data[tab_idx]["name"] = final_name
            self.render_preview_and_selection()
            self.preview_tabs.setCurrentIndex(tab_idx)
            self.statusBar().showMessage(f"Renamed sheet to “{final_name}”.")
            return
        if chosen == act_delete:
            if len(self.sheets_data) <= 1:
                QMessageBox.information(self, "Delete sheet", "At least one sheet must remain.")
                return
            if QMessageBox.question(
                self,
                "Delete sheet",
                f"Delete sheet “{self.sheets_data[tab_idx]['name']}”?",
            ) != QMessageBox.StandardButton.Yes:
                return
            self._push_history_before_change()
            self.sheets_data.pop(tab_idx)
            self.render_preview_and_selection()
            self.preview_tabs.setCurrentIndex(max(0, tab_idx - 1))
            self.statusBar().showMessage("Sheet deleted.")

    def _rename_sheet_from_tab_double_click(self, tab_idx: int):
        if tab_idx is None or tab_idx < 0:
            return
        if tab_idx >= len(self.sheets_data):
            return
        old_name = self.sheets_data[tab_idx]["name"]
        new_name, ok = QInputDialog.getText(self, "Rename sheet", "New sheet name:", text=old_name)
        if not ok:
            return
        new_name = str(new_name or "").strip()
        if not new_name or new_name == old_name:
            return
        used = {s["name"] for i, s in enumerate(self.sheets_data) if i != tab_idx}
        final_name = self.make_unique_sheet_name(new_name, used)
        self._push_history_before_change()
        self.sheets_data[tab_idx]["name"] = final_name
        self.render_preview_and_selection()
        self.preview_tabs.setCurrentIndex(tab_idx)
        self.statusBar().showMessage(f"Renamed sheet to “{final_name}”.")

    def _insert_row_at(self, sheet_idx: int, row_idx: int):
        data = self.sheets_data[sheet_idx]["data"]
        if not data:
            data.append([""])
        width = max((len(r) for r in data), default=1)
        data.insert(max(0, min(row_idx, len(data))), [""] * width)

    def _delete_row_at(self, sheet_idx: int, row_idx: int):
        data = self.sheets_data[sheet_idx]["data"]
        if 0 <= row_idx < len(data):
            data.pop(row_idx)
        if not data:
            data.append([""])

    def _insert_column_at(self, sheet_idx: int, col_idx: int):
        data = self.sheets_data[sheet_idx]["data"]
        if not data:
            data.append([""])
        insert_at = max(0, col_idx)
        for row in data:
            if len(row) < insert_at:
                row.extend([""] * (insert_at - len(row)))
            row.insert(insert_at, "")

    def _delete_column_at(self, sheet_idx: int, col_idx: int):
        data = self.sheets_data[sheet_idx]["data"]
        if not data:
            return
        max_cols = max(len(r) for r in data)
        if max_cols <= 1:
            return
        for row in data:
            if col_idx < len(row):
                row.pop(col_idx)
        if all(len(r) == 0 for r in data):
            for r in data:
                r.append("")

    def _sort_sheet_by_column(self, sheet_idx: int, col_idx: int, reverse: bool):
        data = self.sheets_data[sheet_idx]["data"]
        if len(data) <= 2:
            return
        header = data[0]
        rows = data[1:]
        non_empty_rows = []
        empty_rows = []
        for row in rows:
            row_text = " ".join(str(c).strip() for c in row if str(c).strip())
            if row_text:
                non_empty_rows.append(row)
            else:
                empty_rows.append(row)

        def key_fn(row):
            v = row[col_idx] if col_idx < len(row) else ""
            text = str(v).strip()
            try:
                num = float(text.replace(",", ""))
                return (0, num)
            except ValueError:
                return (1, text.lower())

        non_empty_rows.sort(key=key_fn, reverse=reverse)
        self.sheets_data[sheet_idx]["data"] = [header] + non_empty_rows + empty_rows

    def _show_row_header_context_menu(self, sheet_idx: int, row_idx: int, global_pos):
        if row_idx < 0:
            return
        table = self.preview_tables[sheet_idx] if sheet_idx < len(self.preview_tables) else None
        selected_rows = []
        if table is not None and table.selectionModel():
            selected_rows = sorted({idx.row() for idx in table.selectionModel().selectedRows()})
        targets = selected_rows if selected_rows and row_idx in selected_rows else [row_idx]
        menu = QMenu(self)
        act_insert_above = menu.addAction("Insert row above")
        act_insert_below = menu.addAction("Insert row below")
        act_delete_row = menu.addAction("Delete selected row(s)")
        chosen = menu.exec(global_pos)
        if chosen is None:
            return
        self._push_history_before_change()
        if chosen == act_insert_above:
            self._insert_row_at(sheet_idx, row_idx)
            focus_row = row_idx
        elif chosen == act_insert_below:
            self._insert_row_at(sheet_idx, row_idx + 1)
            focus_row = row_idx + 1
        else:
            for r in sorted(targets, reverse=True):
                self._delete_row_at(sheet_idx, r)
            focus_row = max(0, min(targets) - 1)
        self.render_preview_and_selection()
        self.preview_tabs.setCurrentIndex(sheet_idx)
        if sheet_idx < len(self.preview_tables):
            self.preview_tables[sheet_idx].selectRow(min(focus_row, self.preview_tables[sheet_idx].rowCount() - 1))

    def _show_column_header_context_menu(self, sheet_idx: int, col_idx: int, global_pos):
        if col_idx < 0:
            return
        table = self.preview_tables[sheet_idx] if sheet_idx < len(self.preview_tables) else None
        selected_cols = []
        if table is not None and table.selectionModel():
            selected_cols = sorted({idx.column() for idx in table.selectionModel().selectedColumns()})
        targets = selected_cols if selected_cols and col_idx in selected_cols else [col_idx]
        menu = QMenu(self)
        act_insert_left = menu.addAction("Insert column left")
        act_insert_right = menu.addAction("Insert column right")
        act_delete_col = menu.addAction("Delete selected column(s)")
        menu.addSeparator()
        act_sort_asc = menu.addAction("Sort by this column (A→Z / low→high)")
        act_sort_desc = menu.addAction("Sort by this column (Z→A / high→low)")
        chosen = menu.exec(global_pos)
        if chosen is None:
            return
        self._push_history_before_change()
        if chosen == act_insert_left:
            self._insert_column_at(sheet_idx, col_idx)
        elif chosen == act_insert_right:
            self._insert_column_at(sheet_idx, col_idx + 1)
        elif chosen == act_delete_col:
            for c in sorted(targets, reverse=True):
                self._delete_column_at(sheet_idx, c)
        elif chosen == act_sort_asc:
            self._sort_sheet_by_column(sheet_idx, col_idx, reverse=False)
        elif chosen == act_sort_desc:
            self._sort_sheet_by_column(sheet_idx, col_idx, reverse=True)
        self.render_preview_and_selection()
        self.preview_tabs.setCurrentIndex(sheet_idx)

    def _show_cell_context_menu(self, sheet_idx: int, table: QTableWidget, pos):
        idx = table.indexAt(pos)
        if not idx.isValid():
            return
        row_idx = idx.row()
        col_idx = idx.column()
        menu = QMenu(self)
        act_insert_row_above = menu.addAction("Insert row above")
        act_insert_row_below = menu.addAction("Insert row below")
        act_delete_row = menu.addAction("Delete row")
        menu.addSeparator()
        act_insert_col_left = menu.addAction("Insert column left")
        act_insert_col_right = menu.addAction("Insert column right")
        act_delete_col = menu.addAction("Delete column")
        chosen = menu.exec(table.viewport().mapToGlobal(pos))
        if chosen is None:
            return
        self._push_history_before_change()
        if chosen == act_insert_row_above:
            self._insert_row_at(sheet_idx, row_idx)
        elif chosen == act_insert_row_below:
            self._insert_row_at(sheet_idx, row_idx + 1)
        elif chosen == act_delete_row:
            self._delete_row_at(sheet_idx, row_idx)
        elif chosen == act_insert_col_left:
            self._insert_column_at(sheet_idx, col_idx)
        elif chosen == act_insert_col_right:
            self._insert_column_at(sheet_idx, col_idx + 1)
        elif chosen == act_delete_col:
            self._delete_column_at(sheet_idx, col_idx)
        self.render_preview_and_selection()
        self.preview_tabs.setCurrentIndex(sheet_idx)

    def _persist_row_visual_order(self, sheet_idx: int, table: QTableWidget):
        if sheet_idx >= len(self.sheets_data):
            return
        header = table.verticalHeader()
        count = header.count()
        visual_order = [header.logicalIndex(i) for i in range(count)]
        data = self.sheets_data[sheet_idx]["data"]
        if len(data) != count:
            return
        if visual_order == list(range(count)):
            return
        self._push_history_before_change()
        self.sheets_data[sheet_idx]["data"] = [data[i] for i in visual_order]
        work_idx = self._resolve_flow_working_sheet_idx()
        if work_idx is not None and sheet_idx == work_idx:
            remapped = {}
            remapped_prefilled = set()
            remapped_unlocked = set()
            for new_row, old_row in enumerate(visual_order):
                if new_row == 0:
                    continue
                key = self.flow_row_event_keys.get(old_row)
                if key:
                    remapped[new_row] = key
                if old_row in self.flow_prefilled_event_rows:
                    remapped_prefilled.add(new_row)
                if old_row in self.flow_unlocked_event_rows:
                    remapped_unlocked.add(new_row)
            self.flow_row_event_keys = remapped
            self.flow_prefilled_event_rows = remapped_prefilled
            self.flow_unlocked_event_rows = remapped_unlocked
        current = self.preview_tabs.currentIndex()
        self.render_preview_and_selection()
        self.preview_tabs.setCurrentIndex(current if current >= 0 else sheet_idx)
        self.statusBar().showMessage("Rows reordered.")

    def _persist_column_visual_order(self, sheet_idx: int, table: QTableWidget):
        if sheet_idx >= len(self.sheets_data):
            return
        header = table.horizontalHeader()
        count = header.count()
        visual_order = [header.logicalIndex(i) for i in range(count)]
        if visual_order == list(range(count)):
            return
        self._push_history_before_change()
        data = self.sheets_data[sheet_idx]["data"]
        reordered = []
        for row in data:
            padded = list(row) + [""] * max(0, count - len(row))
            reordered.append([padded[i] for i in visual_order])
        self.sheets_data[sheet_idx]["data"] = reordered
        if self.flow_source_sheet_idx == sheet_idx or self._resolve_flow_working_sheet_idx() == sheet_idx:
            def remap(idx):
                return visual_order.index(idx) if idx is not None and idx in visual_order else idx
            self.flow_amount_col_idx = remap(self.flow_amount_col_idx)
            self.flow_event_col_idx = remap(self.flow_event_col_idx)
        current = self.preview_tabs.currentIndex()
        self.render_preview_and_selection()
        self.preview_tabs.setCurrentIndex(current if current >= 0 else sheet_idx)
        self.statusBar().showMessage("Columns reordered.")

    def _on_row_header_section_moved(
        self,
        sheet_idx: int,
        table: QTableWidget,
        logical_index: int,
        old_visual: int,
        new_visual: int,
    ):
        del logical_index, old_visual, new_visual
        if self._header_move_sync_in_progress:
            return
        self._header_move_sync_in_progress = True
        try:
            self._persist_row_visual_order(sheet_idx, table)
        finally:
            self._header_move_sync_in_progress = False

    def _on_col_header_section_moved(
        self,
        sheet_idx: int,
        table: QTableWidget,
        logical_index: int,
        old_visual: int,
        new_visual: int,
    ):
        del logical_index, old_visual, new_visual
        if self._header_move_sync_in_progress:
            return
        self._header_move_sync_in_progress = True
        try:
            self._persist_column_visual_order(sheet_idx, table)
        finally:
            self._header_move_sync_in_progress = False

    def add_split_point(self):
        current_index = self.preview_tabs.currentIndex()
        if current_index < 0 or current_index >= len(self.preview_tables):
            self.statusBar().showMessage("No preview table available for split marking.")
            return

        table = self.preview_tables[current_index]
        selected_rows = table.selectionModel().selectedRows()
        if not selected_rows:
            self.statusBar().showMessage("Select at least one row to mark a split.")
            return

        split_rows = sorted({idx.row() for idx in selected_rows if idx.row() > 0})
        if not split_rows:
            self.statusBar().showMessage("Cannot mark header row as a split point.")
            return
        rows = self.sheets_data[current_index]["data"]
        valid_splits = [s for s in split_rows if 0 < s < len(rows)]
        if not valid_splits:
            self.statusBar().showMessage("No valid split rows for this sheet.")
            return

        self._push_history_before_change()
        self.split_sheet_at_rows(current_index, split_rows)
        self._reset_flow_session()
        self.render_preview_and_selection()
        self.split_status_label.setText(
            "Split applied at row(s): " + ", ".join(str(r + 1) for r in split_rows)
        )
        self.statusBar().showMessage("Split applied. Preview updated with new sheet tabs.")

    def reset_preview_splits(self):
        if not self.original_sheets_data:
            self.statusBar().showMessage("No loaded data to reset.")
            return
        self._push_history_before_change()
        self._reset_flow_session()
        self.sheets_data = _copy_all_sheets(self.original_sheets_data)
        self.render_preview_and_selection()
        self.split_status_label.setText("No split applied")
        self.statusBar().showMessage("Preview reset to the originally loaded sheet(s).")

    def split_sheet_at_rows(self, sheet_idx, split_rows):
        sheet = self.sheets_data[sheet_idx]
        rows = sheet["data"]
        valid_splits = [s for s in split_rows if 0 < s < len(rows)]
        if not valid_splits:
            return False

        segments = []
        start = 0
        segment_counter = 1
        for split in valid_splits:
            segment_rows = rows[start:split]
            if segment_rows:
                segment_name = sheet["name"] if start == 0 else f"{sheet['name']}_{segment_counter}"
                segments.append({"name": segment_name, "data": segment_rows, "is_table": True})
                segment_counter += 1
            start = split
        tail_rows = rows[start:]
        if tail_rows:
            segment_name = sheet["name"] if start == 0 else f"{sheet['name']}_{segment_counter}"
            segments.append({"name": segment_name, "data": tail_rows, "is_table": True})

        self.sheets_data = self.sheets_data[:sheet_idx] + segments + self.sheets_data[sheet_idx + 1 :]
        return True

    def render_preview_and_selection(self, preserved_selection=None):
        self._is_rendering_preview = True
        try:
            self.preview_tabs.currentChanged.disconnect(self._on_preview_tab_changed)
        except TypeError:
            pass

        self.selected_sheets = {}
        self.preview_tables = []
        self.preview_tabs.clear()

        for i in reversed(range(self.sheets_check_layout.count())):
            widget = self.sheets_check_layout.itemAt(i).widget()
            if widget is not None:
                widget.setParent(None)

        for sheet_idx, sheet_info in enumerate(self.sheets_data):
            sheet_name = sheet_info["name"]
            data = sheet_info["data"]
            table = QTableWidget()
            table.setVerticalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)
            table.setHorizontalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)
            table.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
            table.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
            table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
            table.customContextMenuRequested.connect(
                lambda pos, idx=sheet_idx, t=table: self._show_cell_context_menu(idx, t, pos)
            )
            v_header = table.verticalHeader()
            h_header = table.horizontalHeader()
            v_header.setSectionsMovable(True)
            h_header.setSectionsMovable(True)
            v_header.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
            h_header.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
            v_header.customContextMenuRequested.connect(
                lambda pos, idx=sheet_idx, hdr=v_header: self._show_row_header_context_menu(
                    idx, hdr.logicalIndexAt(pos), hdr.mapToGlobal(pos)
                )
            )
            h_header.customContextMenuRequested.connect(
                lambda pos, idx=sheet_idx, hdr=h_header: self._show_column_header_context_menu(
                    idx, hdr.logicalIndexAt(pos), hdr.mapToGlobal(pos)
                )
            )
            v_header.sectionMoved.connect(
                lambda logical, old_v, new_v, idx=sheet_idx, t=table: self._on_row_header_section_moved(
                    idx, t, logical, old_v, new_v
                )
            )
            h_header.sectionMoved.connect(
                lambda logical, old_v, new_v, idx=sheet_idx, t=table: self._on_col_header_section_moved(
                    idx, t, logical, old_v, new_v
                )
            )
            if not data:
                table.setRowCount(0)
                table.setColumnCount(0)
            else:
                ncols = max(len(row) for row in data) if data else 1
                nrows = len(data)
                table.setColumnCount(ncols)
                table.setRowCount(nrows)
                for row_idx, row in enumerate(data):
                    for col_idx in range(ncols):
                        cell_value = row[col_idx] if col_idx < len(row) else ""
                        item = QTableWidgetItem(str(cell_value))
                        if row_idx == 0:
                            item.setBackground(QColor(68, 114, 196))
                            item.setForeground(QColor(255, 255, 255))
                        table.setItem(row_idx, col_idx, item)
                table.resizeColumnsToContents()
            table.itemChanged.connect(lambda item, idx=sheet_idx: self._on_preview_item_changed(idx, item))
            self.preview_tabs.addTab(table, sheet_name)
            self.preview_tables.append(table)

            ncols_display = max((len(r) for r in data), default=0) if data else 0
            checkbox = QCheckBox(f"{sheet_name} ({len(data)} rows, {ncols_display} cols)")
            checked = True
            if preserved_selection is not None and sheet_idx in preserved_selection:
                checked = bool(preserved_selection[sheet_idx])
            checkbox.blockSignals(True)
            checkbox.setChecked(checked)
            checkbox.blockSignals(False)
            checkbox.stateChanged.connect(lambda state, idx=sheet_idx: self.on_sheet_toggle(idx, state))
            self.sheets_check_layout.addWidget(checkbox)
            self.selected_sheets[sheet_idx] = checked

        self.preview_tabs.currentChanged.connect(self._on_preview_tab_changed)
        self._update_clear_primary_filter_button_state()
        self._refresh_event_widgets_on_working_sheet()
        self._is_rendering_preview = False

    def _on_preview_item_changed(self, sheet_idx: int, item: QTableWidgetItem):
        if self._is_rendering_preview:
            return
        if sheet_idx < 0 or sheet_idx >= len(self.sheets_data):
            return
        if not self._history_suspended:
            self._mark_cell_edit_for_undo_batch()
        row_idx = item.row()
        col_idx = item.column()
        data = self.sheets_data[sheet_idx]["data"]
        while len(data) <= row_idx:
            data.append([])
        row = data[row_idx]
        while len(row) <= col_idx:
            row.append("")
        row[col_idx] = item.text()
        work_idx = self._resolve_flow_working_sheet_idx()
        if (
            work_idx is not None
            and sheet_idx == work_idx
            and row_idx > 0
            and (col_idx == self.flow_amount_col_idx or col_idx == self.flow_event_col_idx)
        ):
            if col_idx == self.flow_event_col_idx and not str(item.text() or "").strip():
                self.flow_row_event_keys.pop(row_idx, None)
            self._recompute_flow_mapping_on_working_sheet(render=False, announce=False)

    def _refresh_event_widgets_on_working_sheet(self):
        work_idx = self._resolve_flow_working_sheet_idx()
        if work_idx is None or self.flow_event_col_idx is None:
            return
        if work_idx >= len(self.preview_tables) or work_idx >= len(self.sheets_data):
            return
        table = self.preview_tables[work_idx]
        data = self.sheets_data[work_idx]["data"]
        col_idx = self.flow_event_col_idx
        if not data:
            return
        focus_widget = QApplication.focusWidget()
        focused_row = None
        focused_cursor_pos = None
        if isinstance(focus_widget, QLineEdit):
            parent_widget = focus_widget.parentWidget()
            if isinstance(parent_widget, EventCellWidget):
                for row_idx in range(1, len(data)):
                    if table.cellWidget(row_idx, col_idx) is parent_widget:
                        focused_row = row_idx
                        focused_cursor_pos = focus_widget.cursorPosition()
                        break
        self._is_rendering_preview = True
        try:
            for row_idx in range(1, len(data)):
                if col_idx >= len(data[row_idx]):
                    data[row_idx].extend([""] * (col_idx - len(data[row_idx]) + 1))
                current_text = str(data[row_idx][col_idx] or "")
                selected_key = str(self.flow_row_event_keys.get(row_idx, "") or "")
                widget = table.cellWidget(row_idx, col_idx)
                if isinstance(widget, EventCellWidget):
                    widget.sync_state(
                        self.flow_event_options,
                        text=current_text,
                        selected_key=selected_key,
                        combo_enabled=True,
                    )
                else:
                    table.removeCellWidget(row_idx, col_idx)
                    widget = EventCellWidget(
                        self.flow_event_options,
                        text=current_text,
                        selected_key=selected_key,
                        parent=table,
                    )
                    widget.combo.setEnabled(True)
                    widget.textCommitted.connect(
                        lambda text, r=row_idx: self._on_event_text_committed(r, text)
                    )
                    widget.optionPicked.connect(
                        lambda text, r=row_idx: self._on_event_option_changed(r, text)
                    )
                    table.setCellWidget(row_idx, col_idx, widget)
        finally:
            self._is_rendering_preview = False
        if focused_row is not None:
            focused_widget = table.cellWidget(focused_row, col_idx)
            if isinstance(focused_widget, EventCellWidget):
                focused_widget.text_input.setFocus()
                if focused_cursor_pos is not None:
                    focused_widget.text_input.setCursorPosition(
                        min(focused_cursor_pos, len(focused_widget.text_input.text()))
                    )

    def _on_event_text_committed(self, row_idx: int, text: str):
        work_idx = self._resolve_flow_working_sheet_idx()
        if work_idx is None or self.flow_event_col_idx is None:
            return
        if not self._history_suspended:
            self._mark_cell_edit_for_undo_batch()
        if work_idx >= len(self.sheets_data):
            return
        data = self.sheets_data[work_idx]["data"]
        if row_idx >= len(data):
            return
        col = self.flow_event_col_idx
        row = data[row_idx]
        if col >= len(row):
            row.extend([""] * (col - len(row) + 1))
        row[col] = text
        text_clean = str(text or "").strip()
        if not text_clean:
            if row_idx in self.flow_prefilled_event_rows:
                self.flow_unlocked_event_rows.add(row_idx)
            self.flow_row_event_keys.pop(row_idx, None)
        elif self._is_known_header_text(text_clean, data):
            # Manual overwrite to a known header immediately changes mapping destination.
            self.flow_row_event_keys[row_idx] = text_clean
        self._recompute_flow_mapping_on_working_sheet(render=False, announce=False)

    def _on_event_option_changed(self, row_idx: int, text: str):
        work_idx = self._resolve_flow_working_sheet_idx()
        if work_idx is None or self.flow_event_col_idx is None:
            return
        if not self._history_suspended:
            self._mark_cell_edit_for_undo_batch()
        if work_idx >= len(self.sheets_data):
            return
        data = self.sheets_data[work_idx]["data"]
        if row_idx >= len(data):
            return
        col = self.flow_event_col_idx
        row = data[row_idx]
        if col >= len(row):
            row.extend([""] * (col - len(row) + 1))
        current_text = str(row[col] or "").strip()
        if text:
            self.flow_row_event_keys[row_idx] = text
            if not current_text:
                row[col] = text
            elif self._is_known_header_text(current_text, data):
                # Existing header text should be replaced by selected option text.
                row[col] = text
        else:
            self.flow_row_event_keys.pop(row_idx, None)
        self._recompute_flow_mapping_on_working_sheet(render=False, announce=False)

    def _is_known_header_text(self, text: str, data: list[list[object]]) -> bool:
        if not text or not data:
            return False
        header = data[0] if data else []
        normalized = normalize_header(text)
        if not normalized:
            return False
        for value in header:
            if normalize_header(value) == normalized:
                return True
        return False

    def make_unique_sheet_name(self, base_name, used_names):
        name = base_name
        index = 2
        while name in used_names:
            name = f"{base_name}_{index}"
            index += 1
        used_names.add(name)
        return name

    def browse_output(self):
        home_dir = str(Path.home() / "Desktop")
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "Save Excel File",
            os.path.join(home_dir, "bank_statements.xlsx"),
            "Excel Files (*.xlsx)",
        )
        if file_path:
            self.output_path_input.setText(file_path)

    def convert_and_save(self):
        if not self.sheets_data:
            self.statusBar().showMessage("Error: No data. Load a PDF or Excel file first.")
            return

        if not self.output_path_input.text():
            self.statusBar().showMessage("Error: Specify output file path.")
            return

        if not any(self.selected_sheets.values()):
            self.statusBar().showMessage("Error: Select at least one sheet to save.")
            return

        final_sheets = []
        used_names = set()
        for sheet_idx, sheet in enumerate(self.sheets_data):
            if not self.selected_sheets.get(sheet_idx, False):
                continue
            name = self.make_unique_sheet_name(sheet["name"], used_names)
            final_sheets.append({"name": name, "data": sheet["data"], "is_table": True})

        try:
            from converter import export_to_excel

            output_path = self.output_path_input.text()

            if not output_path.lower().endswith(".xlsx"):
                output_path += ".xlsx"

            export_to_excel(final_sheets, output_path)
            self.statusBar().showMessage(f"✓ Excel file saved successfully: {output_path}")

            if os.name == "nt":
                os.startfile(os.path.dirname(output_path))

        except Exception as e:
            self.statusBar().showMessage(f"Error saving: {str(e)}")

    def _full_grid_from_sheet_idx(self, sheet_idx: int):
        if sheet_idx < 0 or sheet_idx >= len(self.sheets_data):
            return []
        return [list(r) for r in self.sheets_data[sheet_idx].get("data", [])]

    def _selection_grid_from_active_table(self):
        sheet_idx = self.preview_tabs.currentIndex()
        if sheet_idx < 0 or sheet_idx >= len(self.preview_tables):
            return []
        table = self.preview_tables[sheet_idx]
        sel_model = table.selectionModel()
        if sel_model is None:
            return []
        selected = sel_model.selectedIndexes()
        if not selected:
            return []
        min_row = min(i.row() for i in selected)
        max_row = max(i.row() for i in selected)
        min_col = min(i.column() for i in selected)
        max_col = max(i.column() for i in selected)
        selected_lookup = {(i.row(), i.column()): i for i in selected}
        grid = []
        for r in range(min_row, max_row + 1):
            out_row = []
            for c in range(min_col, max_col + 1):
                if (r, c) in selected_lookup:
                    item = table.item(r, c)
                    out_row.append(item.text() if item else "")
                else:
                    out_row.append("")
            grid.append(out_row)
        return grid

    def update_existing_excel(self):
        if not self.sheets_data:
            self.statusBar().showMessage("Load a PDF or Excel file first.")
            return
        active_idx = self.preview_tabs.currentIndex()
        if active_idx < 0 or active_idx >= len(self.sheets_data):
            active_idx = 0
        active_name = self.sheets_data[active_idx]["name"]
        has_selection = bool(self._selection_grid_from_active_table())
        dlg = UpdateExistingExcelDialog(
            self,
            active_sheet_name=active_name,
            has_selection=has_selection,
        )
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return
        config = dlg.result()
        if not config:
            return
        dest_path = config["path"]
        try:
            if config["mode"] == "add_sheets":
                sheets_to_add = []
                used = set()
                for sheet_idx, sheet in enumerate(self.sheets_data):
                    if not self.selected_sheets.get(sheet_idx, False):
                        continue
                    name = self.make_unique_sheet_name(sheet["name"], used)
                    sheets_to_add.append({"name": name, "data": self._full_grid_from_sheet_idx(sheet_idx)})
                if not sheets_to_add:
                    QMessageBox.information(self, "Update existing workbook", "Select at least one app sheet to add.")
                    return
                from converter import append_sheets_to_existing_workbook
                append_sheets_to_existing_workbook(dest_path, sheets_to_add)
                self.statusBar().showMessage(f"Added {len(sheets_to_add)} sheet(s) to existing workbook.")
                return

            source_mode = config.get("source_mode")
            if source_mode == "selection":
                grid = self._selection_grid_from_active_table()
                if not grid:
                    QMessageBox.information(self, "Update existing workbook", "No highlighted cells found on active sheet.")
                    return
            else:
                grid = self._full_grid_from_sheet_idx(active_idx)
                if not grid:
                    QMessageBox.information(self, "Update existing workbook", "Active sheet has no data.")
                    return
            rows = len(grid)
            cols = max((len(r) for r in grid), default=0)
            from converter import has_nonempty_cells_in_target_range, paste_values_into_existing_sheet
            if has_nonempty_cells_in_target_range(
                dest_path,
                config["sheet_name"],
                config["start_cell"],
                rows,
                cols,
            ):
                if QMessageBox.question(
                    self,
                    "Overwrite destination cells?",
                    "Destination range already has values. Continue and overwrite those cells?",
                ) != QMessageBox.StandardButton.Yes:
                    return
            paste_values_into_existing_sheet(
                dest_path,
                config["sheet_name"],
                config["start_cell"],
                grid,
                clear_grid=bool(config.get("clear_first")),
            )
            self.statusBar().showMessage("Pasted values into existing workbook successfully.")
        except PermissionError:
            QMessageBox.warning(
                self,
                "Workbook locked",
                "Could not write to destination workbook. Close it in Excel and try again.",
            )
        except Exception as e:
            QMessageBox.critical(self, "Update existing workbook", str(e))
