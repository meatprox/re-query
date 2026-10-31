from __future__ import annotations

from datetime import date
from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QButtonGroup,
    QDialog,
    QFileDialog,
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QPushButton,
    QRadioButton,
    QVBoxLayout,
    QWidget,
)

from .auto_map import MODE_ALL, MODE_NUMBERS, build_auto_mappings
from .excel_io import dump_mappings, load_workbook_file
from .models import WorkbookData

EXCEL_FILTER = "Excel / CSV (*.xlsx *.xlsm *.csv);;すべてのファイル (*.*)"


class _FileSlot(QWidget):
    def __init__(self, number: str, title: str, optional: bool = False, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.book: WorkbookData | None = None
        self.path: str | None = None
        self.optional = optional
        self._title = title

        caption = QLabel(f"{number} {title}" + ("（任意）" if optional else ""))
        caption.setObjectName("slotTitle")
        self.path_label = QLabel("未選択")
        self.path_label.setObjectName("slotPath")
        self.path_label.setWordWrap(True)
        self.btn = QPushButton("読み込む")
        self.btn.setObjectName("paneAction")
        self.btn.clicked.connect(self._choose)

        row = QHBoxLayout()
        row.setContentsMargins(0, 0, 0, 0)
        row.setSpacing(8)
        row.addWidget(self.path_label, 1)
        row.addWidget(self.btn, 0)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(4)
        layout.addWidget(caption)
        layout.addLayout(row)

    def _choose(self) -> None:
        path, _ = QFileDialog.getOpenFileName(self, self._title, "", EXCEL_FILTER)
        if not path:
            return
        try:
            book = load_workbook_file(path)
        except Exception as exc:
            QMessageBox.critical(self, "読み込みエラー", str(exc))
            return
        self.book = book
        self.path = str(Path(path).resolve())
        self.path_label.setText(Path(path).name)
        self.path_label.setToolTip(self.path)


class AutoMapDialog(QDialog):
    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("配置関係自動作成")
        self.resize(560, 460)
        self.setModal(True)

        intro = QLabel(
            "データファイルと記入済みの出力ファイルを比較し、同じ値が書かれたセルから配置関係を作ります。"
            "未記入の出力フォーマットを指定すると、もともと書いてあるセルは対象外にします。"
        )
        intro.setWordWrap(True)
        intro.setObjectName("intro")

        self.slot_source = _FileSlot("1)", "データファイル")
        self.slot_filled = _FileSlot("2)", "記入済みの出力ファイル")
        self.slot_blank = _FileSlot("3)", "未記入の出力フォーマット", optional=True)

        mode_label = QLabel("比較対象")
        mode_label.setObjectName("slotTitle")
        self.radio_numbers = QRadioButton("数値のみ")
        self.radio_all = QRadioButton("すべて")
        self.radio_numbers.setChecked(True)
        group = QButtonGroup(self)
        group.addButton(self.radio_numbers)
        group.addButton(self.radio_all)
        mode_row = QHBoxLayout()
        mode_row.setContentsMargins(0, 0, 0, 0)
        mode_row.setSpacing(16)
        mode_row.addWidget(self.radio_numbers)
        mode_row.addWidget(self.radio_all)
        mode_row.addStretch(1)

        self.btn_create = QPushButton("配置関係ファイルを作成")
        self.btn_create.setObjectName("paneAction")
        self.btn_create.setDefault(True)
        self.btn_create.clicked.connect(self._create)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(18, 16, 18, 16)
        layout.setSpacing(14)
        layout.addWidget(intro)
        layout.addWidget(self.slot_source)
        layout.addWidget(self.slot_filled)
        layout.addWidget(self.slot_blank)
        layout.addWidget(mode_label)
        layout.addLayout(mode_row)
        layout.addStretch(1)
        layout.addWidget(self.btn_create, 0, Qt.AlignRight)
        self.setStyleSheet(
            """
            QDialog { background: #ece6da; }
            QLabel#intro { color: #4d5a4d; font-size: 12px; }
            QLabel#slotTitle { font-weight: 700; font-size: 12px; }
            QLabel#slotPath { color: #5c6b5c; font-size: 12px; }
            QRadioButton { font-size: 13px; }
            QPushButton#paneAction {
                background: #2d5a3d;
                color: #f7fff8;
                border: none;
                border-radius: 5px;
                padding: 7px 14px;
                font-size: 12px;
            }
            QPushButton#paneAction:hover { background: #3c9a5c; }
            """
        )

    def _mode(self) -> str:
        return MODE_NUMBERS if self.radio_numbers.isChecked() else MODE_ALL

    def _suggested_name(self) -> str:
        day = date.today().strftime("%Y%m%d")
        source = self.slot_blank.path or self.slot_filled.path or "配置関係"
        stem = Path(source).stem
        return f"{stem}_{day}.json"

    def _create(self) -> None:
        if self.slot_source.book is None or self.slot_source.path is None:
            QMessageBox.information(self, "配置関係自動作成", "データファイルを読み込んでください。")
            return
        if self.slot_filled.book is None or self.slot_filled.path is None:
            QMessageBox.information(self, "配置関係自動作成", "記入済みの出力ファイルを読み込んでください。")
            return
        path, _ = QFileDialog.getSaveFileName(self, "配置関係ファイルを作成", self._suggested_name(), "JSON (*.json)")
        if not path:
            return
        dest = Path(path)
        if dest.suffix.lower() != ".json":
            dest = dest.with_suffix(".json")
        mappings = build_auto_mappings(
            self.slot_source.book,
            self.slot_filled.book,
            self.slot_blank.book,
            self._mode(),
        )
        format_path = self.slot_blank.path
        try:
            dest.write_text(dump_mappings(mappings, format_path), encoding="utf-8")
        except Exception as exc:
            QMessageBox.critical(self, "保存エラー", str(exc))
            return
        QMessageBox.information(
            self,
            "作成完了",
            f"配置関係ファイルを作成しました。\n\n"
            f"ファイル名: {dest.name}\n"
            f"件数: {len(mappings)} 件\n\n"
            "メイン画面で配置関係を読み込み、中身を確認してください。",
        )
