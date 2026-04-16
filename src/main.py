"""
Bank Statement PDF-to-Excel Converter
Main entry point for the application
"""
import sys
import os

def _app_base_dir():
    """Directory containing converter, ui, etc. (src/ in dev; PyInstaller _MEIPASS when frozen)."""
    if getattr(sys, "frozen", False):
        return getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.dirname(os.path.abspath(__file__))


sys.path.insert(0, _app_base_dir())

from PyQt6.QtCore import QEvent
from PyQt6.QtGui import QKeyEvent
from PyQt6.QtWidgets import QApplication, QWidget
from ui.main_window import MainWindow


class BankConverterApp(QApplication):
    """Intercept undo/redo keys before QTableWidget/cell editors consume them."""

    def notify(self, receiver, event):
        if event.type() == QEvent.Type.KeyPress and isinstance(event, QKeyEvent):
            if isinstance(receiver, QWidget):
                mw = self.activeWindow()
                if isinstance(mw, MainWindow) and mw.try_consume_undo_redo_key(receiver, event):
                    return True
        return super().notify(receiver, event)


def main():
    """Main application entry point"""
    app = BankConverterApp(sys.argv)
    
    # Set application properties
    app.setApplicationName('Bank Statement Converter')
    app.setApplicationVersion('1.1')
    
    # Create and show main window
    window = MainWindow()
    window.show()
    
    sys.exit(app.exec())


if __name__ == '__main__':
    main()
