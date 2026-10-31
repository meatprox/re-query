from __future__ import annotations

import sys

from PySide6.QtGui import QColor, QFont, QPalette
from PySide6.QtWidgets import QApplication

from .window import MainWindow

APP_STYLE = """
QMainWindow { background: #ece6da; }
QStatusBar {
    background: #ece6da;
    color: #4d5a4d;
    border-top: 1px solid #d7d0c4;
}
QSplitter::handle { background: #ddd6c8; width: 1px; }
QMessageBox { font-size: 13px; }
QScrollBar:vertical {
    background: #c9bfb2;
    width: 14px;
    margin: 0;
    border: none;
}
QScrollBar::handle:vertical {
    background: #3d5c47;
    min-height: 32px;
    border-radius: 5px;
    margin: 2px;
}
QScrollBar::handle:vertical:hover { background: #2d5a3d; }
QScrollBar::add-line:vertical,
QScrollBar::sub-line:vertical { height: 0; }
QScrollBar::add-page:vertical,
QScrollBar::sub-page:vertical { background: #c9bfb2; }
QScrollBar:horizontal {
    background: #c9bfb2;
    height: 14px;
    margin: 0;
    border: none;
}
QScrollBar::handle:horizontal {
    background: #3d5c47;
    min-width: 32px;
    border-radius: 5px;
    margin: 2px;
}
QScrollBar::handle:horizontal:hover { background: #2d5a3d; }
QScrollBar::add-line:horizontal,
QScrollBar::sub-line:horizontal { width: 0; }
QScrollBar::add-page:horizontal,
QScrollBar::sub-page:horizontal { background: #c9bfb2; }
"""


def _palette() -> QPalette:
    palette = QPalette()
    palette.setColor(QPalette.Window, QColor("#ece6da"))
    palette.setColor(QPalette.WindowText, QColor("#1c241c"))
    palette.setColor(QPalette.Base, QColor("#fffdf8"))
    palette.setColor(QPalette.Text, QColor("#1c241c"))
    palette.setColor(QPalette.Button, QColor("#1e3a2b"))
    palette.setColor(QPalette.ButtonText, QColor("#f7fff8"))
    palette.setColor(QPalette.Highlight, QColor("#2f7d4a"))
    palette.setColor(QPalette.HighlightedText, QColor("#ffffff"))
    return palette


def main() -> int:
    app = QApplication(sys.argv)
    app.setApplicationName("Excechenger")
    app.setOrganizationName("Excechenger")
    app.setStyle("Fusion")
    app.setPalette(_palette())
    app.setFont(QFont("Yu Gothic UI", 10))
    app.setStyleSheet(APP_STYLE)
    window = MainWindow()
    window.showMaximized()
    return app.exec()
