from __future__ import annotations

import json

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QFrame,
    QHBoxLayout,
    QLabel,
    QMenu,
    QPlainTextEdit,
    QPushButton,
    QScrollArea,
    QSizePolicy,
    QToolButton,
    QVBoxLayout,
    QWidget,
)

from .models import Mapping
from .spreadsheet import MIME_RANGE


class MappingItem(QFrame):
    description_changed = Signal(str, str)
    removed = Signal(str)
    activated = Signal(str)
    chosen = Signal(str)

    def __init__(self, mapping: Mapping, index: int, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.mapping_id = mapping.id
        self.setObjectName("mappingItem")
        self.setProperty("color", mapping.color)
        self.setCursor(Qt.PointingHandCursor)

        swatch = QFrame()
        swatch.setFixedWidth(8)
        swatch.setStyleSheet(f"background: {mapping.color}; border-radius: 4px;")

        self.title = QLabel()
        self.title.setObjectName("mapTitle")
        self.sub = QLabel()
        self.sub.setObjectName("mapSub")
        self.warn = QLabel()
        self.warn.setObjectName("mapWarn")
        self.warn.setWordWrap(True)

        meta = QVBoxLayout()
        meta.setContentsMargins(0, 0, 0, 0)
        meta.setSpacing(2)
        meta.addWidget(self.title)
        meta.addWidget(self.sub)
        meta.addWidget(self.warn)

        self.desc = QPlainTextEdit()
        self.desc.setPlaceholderText("説明")
        self.desc.setPlainText(mapping.description)
        self.desc.setFixedWidth(124)
        self.desc.setFixedHeight(54)
        self.desc.setObjectName("mapDesc")
        self.desc.setAcceptDrops(False)
        self.desc.textChanged.connect(self._on_desc)

        self.btn_delete = QPushButton("×")
        self.btn_delete.setObjectName("mapDel")
        self.btn_delete.setFixedSize(24, 24)
        self.btn_delete.clicked.connect(lambda: self.removed.emit(self.mapping_id))
        self.btn_choose = QPushButton("○")
        self.btn_choose.setObjectName("mapChoose")
        self.btn_choose.setFixedSize(24, 24)
        self.btn_choose.setToolTip("この候補に決定")
        self.btn_choose.clicked.connect(lambda: self.chosen.emit(self.mapping_id))

        side = QVBoxLayout()
        side.setContentsMargins(0, 0, 0, 0)
        side.setSpacing(4)
        side.addWidget(self.btn_delete, 0, Qt.AlignHCenter | Qt.AlignTop)
        side.addStretch(1)
        side.addWidget(self.btn_choose, 0, Qt.AlignHCenter | Qt.AlignBottom)

        layout = QHBoxLayout(self)
        layout.setContentsMargins(8, 8, 8, 8)
        layout.setSpacing(8)
        layout.addWidget(swatch)
        layout.addLayout(meta, 1)
        layout.addWidget(self.desc)
        layout.addLayout(side)
        self.set_active(False)
        self._set_labels(mapping, index)

    def _set_labels(self, mapping: Mapping, index: int) -> None:
        self.title.setText(f"{index}. {mapping.source.label()} → {mapping.dest.label()}")
        self.sub.setText(f"データ {mapping.source.sheet}  /  出 {mapping.dest.sheet}")
        if mapping.can_cycle():
            current = mapping.choice_index() + 1
            total = len(mapping.candidates)
            self.warn.setText(f"複数一致 {total}件  {current}/{total}　クリックで切替")
            self.warn.show()
            self.btn_choose.show()
        else:
            self.warn.clear()
            self.warn.hide()
            self.btn_choose.hide()

    def update_mapping(self, mapping: Mapping, index: int) -> None:
        self._set_labels(mapping, index)
        if self.desc.toPlainText() != mapping.description:
            self.desc.blockSignals(True)
            self.desc.setPlainText(mapping.description)
            self.desc.blockSignals(False)

    def _on_desc(self) -> None:
        self.description_changed.emit(self.mapping_id, self.desc.toPlainText())

    def mousePressEvent(self, event) -> None:  # type: ignore[override]
        if event.button() == Qt.LeftButton:
            self.activated.emit(self.mapping_id)
        super().mousePressEvent(event)

    def set_active(self, active: bool) -> None:
        color = self.property("color") or "#2f7d4a"
        border = f"2px solid {color}" if active else "1px solid #d7d0c4"
        self.setStyleSheet(
            f"""
            QFrame#mappingItem {{
                background: #fff;
                border: {border};
                border-radius: 8px;
            }}
            QLabel#mapTitle {{ font-weight: 700; font-size: 12px; font-family: Consolas, monospace; }}
            QLabel#mapSub {{ color: #5c6b5c; font-size: 10px; }}
            QLabel#mapWarn {{ color: #b54708; font-size: 10px; font-weight: 700; }}
            QPlainTextEdit#mapDesc {{
                background: #fffcf6;
                border: 1px solid #d7d0c4;
                border-radius: 6px;
                font-size: 11px;
            }}
            QPushButton#mapDel {{
                border: none;
                color: #7a7a7a;
                font-size: 16px;
            }}
            QPushButton#mapDel:hover {{ color: #b42318; }}
            QPushButton#mapChoose {{
                border: 2px solid #2d5a3d;
                color: #2d5a3d;
                background: #f7fff8;
                border-radius: 12px;
                font-size: 13px;
                font-weight: 700;
            }}
            QPushButton#mapChoose:hover {{
                background: #2d5a3d;
                color: #f7fff8;
            }}
            """
        )


class MappingPanel(QWidget):
    range_dropped = Signal(dict)
    description_changed = Signal(str, str)
    removed = Signal(str)
    activated = Signal(str)
    chosen = Signal(str)
    load_clicked = Signal()
    save_clicked = Signal()
    clear_clicked = Signal()
    auto_map_clicked = Signal()
    select_move_toggled = Signal(bool)

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setAcceptDrops(True)
        self._ids: list[str] = []
        self._items: dict[str, MappingItem] = {}

        title = QLabel("配置リスト")
        title.setObjectName("paneTitle")
        self.btn_menu = QToolButton()
        self.btn_menu.setText("メニュー")
        self.btn_menu.setObjectName("paneMenu")
        self.btn_menu.setPopupMode(QToolButton.InstantPopup)
        self.btn_menu.setToolButtonStyle(Qt.ToolButtonTextOnly)
        menu = QMenu(self.btn_menu)
        auto_map = menu.addAction("配置関係自動作成")
        auto_map.triggered.connect(lambda: self.auto_map_clicked.emit())
        self.act_select_move = menu.addAction("選択して移動")
        self.act_select_move.setCheckable(True)
        self.act_select_move.toggled.connect(self.select_move_toggled.emit)
        self.btn_menu.setMenu(menu)
        title_row = QHBoxLayout()
        title_row.setContentsMargins(0, 0, 0, 0)
        title_row.setSpacing(8)
        title_row.addWidget(title)
        title_row.addStretch(1)
        title_row.addWidget(self.btn_menu)
        self.file_label = QLabel("未読込")
        self.file_label.setObjectName("paneFile")
        self.file_label.setWordWrap(True)

        self.btn_load = QPushButton("配置関係を読込")
        self.btn_save = QPushButton("配置関係を保存")
        self.btn_clear = QPushButton("配置をクリア")
        for button in (self.btn_load, self.btn_save, self.btn_clear):
            button.setObjectName("paneAction")
        self.btn_load.clicked.connect(lambda: self.load_clicked.emit())
        self.btn_save.clicked.connect(lambda: self.save_clicked.emit())
        self.btn_clear.clicked.connect(lambda: self.clear_clicked.emit())

        actions = QHBoxLayout()
        actions.setContentsMargins(0, 4, 0, 0)
        actions.setSpacing(6)
        actions.addWidget(self.btn_load)
        actions.addWidget(self.btn_save)
        actions.addWidget(self.btn_clear)
        actions.addStretch(1)

        head = QWidget()
        head.setObjectName("paneHead")
        head_layout = QVBoxLayout(head)
        head_layout.setContentsMargins(10, 8, 10, 8)
        head_layout.setSpacing(2)
        head_layout.addLayout(title_row)
        head_layout.addWidget(self.file_label)
        head_layout.addLayout(actions)

        self.list_host = QWidget()
        self.list_layout = QVBoxLayout(self.list_host)
        self.list_layout.setContentsMargins(10, 6, 10, 12)
        self.list_layout.setSpacing(8)
        self.list_layout.addStretch(1)

        self.empty = QLabel(
            "データ元の選択範囲をこのリストへドラッグして登録します。\n"
            "左右の色枠が対応関係です。出力側の色枠をドラッグすると貼り付け位置を移動できます。"
        )
        self.empty.setWordWrap(True)
        self.empty.setObjectName("emptyHint")
        self.list_layout.insertWidget(0, self.empty)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        scroll.setWidget(self.list_host)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)
        layout.addWidget(head)
        layout.addWidget(scroll, 1)
        self.setMinimumWidth(320)
        self.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Expanding)
        self.setStyleSheet(
            """
            QWidget#paneHead { background: #f7f0e4; border-bottom: 1px solid #d7d0c4; }
            QLabel#paneTitle { font-weight: 700; font-size: 12px; letter-spacing: 1px; }
            QLabel#paneFile { color: #5c6b5c; font-size: 11px; }
            QLabel#emptyHint { color: #5c6b5c; font-size: 12px; }
            QPushButton#paneAction {
                background: #2d5a3d;
                color: #f7fff8;
                border: none;
                border-radius: 5px;
                padding: 5px 10px;
                font-size: 12px;
            }
            QPushButton#paneAction:hover { background: #3c9a5c; }
            QPushButton#paneAction:disabled { background: #9aa89d; color: #eef3ee; }
            QToolButton#paneMenu {
                background: #2d5a3d;
                color: #f7fff8;
                border: none;
                border-radius: 5px;
                padding: 4px 10px;
                font-size: 12px;
            }
            QToolButton#paneMenu:hover { background: #3c9a5c; }
            QToolButton#paneMenu::menu-indicator { image: none; }
            """
        )

    def set_file_name(self, name: str | None) -> None:
        text = name.strip() if name and name.strip() else "未読込"
        self.file_label.setText(text)
        self.file_label.setToolTip(text)

    def set_mappings(self, mappings: list[Mapping], active_id: str | None) -> None:
        ids = [item.id for item in mappings]
        self.empty.setVisible(not mappings)
        missing = any(mapping_id not in self._items for mapping_id in ids)
        if ids != self._ids or missing:
            for item in list(self._items.values()):
                item.setParent(None)
                item.deleteLater()
            self._items.clear()
            self._ids = ids
            for index, mapping in enumerate(mappings, start=1):
                widget = MappingItem(mapping, index)
                widget.description_changed.connect(self.description_changed)
                widget.removed.connect(self.removed)
                widget.activated.connect(self.activated)
                widget.chosen.connect(self.chosen)
                self._items[mapping.id] = widget
                self.list_layout.insertWidget(self.list_layout.count() - 1, widget)
        else:
            for index, mapping in enumerate(mappings, start=1):
                self._items[mapping.id].update_mapping(mapping, index)
        for mapping_id, widget in self._items.items():
            widget.set_active(mapping_id == active_id)

    def dragEnterEvent(self, event) -> None:  # type: ignore[override]
        if event.mimeData().hasFormat(MIME_RANGE):
            event.acceptProposedAction()
        else:
            event.ignore()

    def dragMoveEvent(self, event) -> None:  # type: ignore[override]
        if event.mimeData().hasFormat(MIME_RANGE):
            event.acceptProposedAction()
        else:
            event.ignore()

    def dropEvent(self, event) -> None:  # type: ignore[override]
        if not event.mimeData().hasFormat(MIME_RANGE):
            event.ignore()
            return
        payload = json.loads(bytes(event.mimeData().data(MIME_RANGE)).decode("utf-8"))
        self.range_dropped.emit(payload)
        event.acceptProposedAction()
