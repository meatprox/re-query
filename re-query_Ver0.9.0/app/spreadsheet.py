from __future__ import annotations

import json
from typing import Any

from PySide6.QtCore import QByteArray, QMimeData, QRect, Qt, Signal
from PySide6.QtGui import QColor, QDrag, QFont, QPainter, QPen, QPixmap
from PySide6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QPushButton,
    QSizePolicy,
    QStyledItemDelegate,
    QStyle,
    QStyleOptionViewItem,
    QTabBar,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from .models import CellStyle, Mapping, SheetData, WorkbookData, col_letter, range_a1

MIME_RANGE = "application/x-excechenger-range"


class RangeDelegate(QStyledItemDelegate):
    def paint(self, painter: QPainter, option: QStyleOptionViewItem, index) -> None:  # type: ignore[override]
        view: SpreadsheetView = self.parent()  # type: ignore[assignment]
        painter.save()
        painter.setRenderHint(QPainter.Antialiasing, False)
        rect = option.rect
        row = index.row()
        col = index.column()
        style = view.styles.get((row, col))
        if style is not None and style.fill:
            painter.fillRect(rect, QColor(style.fill))
        else:
            painter.fillRect(rect, QColor("#ffffff"))

        for mapping in view.mappings:
            rng = view.range_of(mapping)
            if rng is None or rng.sheet != view.sheet_name or not rng.contains(row, col):
                continue
            fill = QColor(mapping.color)
            grouped = mapping.id in view.selected_ids
            if mapping.id == view.active_id:
                fill.setAlpha(58)
            elif grouped:
                fill.setAlpha(72)
            else:
                fill.setAlpha(34)
            painter.fillRect(rect, fill)

        selected = bool(option.state & QStyle.State_Selected) and (
            view.role == "source" or view.select_move
        )
        if selected:
            painter.fillRect(rect, QColor(33, 115, 70, 55))

        value = index.data(Qt.DisplayRole)
        if value not in (None, ""):
            text = str(value)
            painter.setPen(QColor(style.font_color if style and style.font_color else "#1c241c"))
            font = QFont(option.font)
            font.setPointSize(9)
            font.setBold(bool(style and style.bold))
            painter.setFont(font)
            align = Qt.AlignRight | Qt.AlignVCenter if _is_number(value) else Qt.AlignLeft | Qt.AlignVCenter
            painter.drawText(rect.adjusted(5, 0, -5, 0), int(align), text)

        painter.setPen(QPen(QColor("#e1dacd"), 1))
        painter.drawRect(rect.adjusted(0, 0, -1, -1))
        if style is not None:
            _paint_excel_borders(painter, rect, style)

        for mapping in view.mappings:
            rng = view.range_of(mapping)
            if rng is None or rng.sheet != view.sheet_name or not rng.contains(row, col):
                continue
            color = QColor(mapping.color)
            grouped = mapping.id in view.selected_ids
            if mapping.id == view.active_id or grouped:
                width = 3
            else:
                width = 2
            painter.setPen(QPen(color, width))
            box = rect.adjusted(1, 1, -1, -1)
            if row == rng.r1:
                painter.drawLine(box.topLeft(), box.topRight())
            if row == rng.r2:
                painter.drawLine(box.bottomLeft(), box.bottomRight())
            if col == rng.c1:
                painter.drawLine(box.topLeft(), box.bottomLeft())
            if col == rng.c2:
                painter.drawLine(box.topRight(), box.bottomRight())
            if row == rng.r1 and col == rng.c1:
                badge = QRect(box.left() + 1, box.top() + 1, 16, 12)
                painter.fillRect(badge, color)
                painter.setPen(QColor("#ffffff"))
                font = QFont(option.font)
                font.setPointSize(7)
                font.setBold(True)
                painter.setFont(font)
                painter.drawText(badge, Qt.AlignCenter, str(view.mapping_index(mapping.id)))

        if selected:
            painter.setPen(QPen(QColor("#217346"), 2))
            painter.drawRect(rect.adjusted(1, 1, -2, -2))

        painter.restore()


def _paint_excel_borders(painter: QPainter, rect: QRect, style: CellStyle) -> None:
    edges = (
        ("top", style.border_top),
        ("right", style.border_right),
        ("bottom", style.border_bottom),
        ("left", style.border_left),
    )
    for side, edge in edges:
        if edge is None:
            continue
        width, color = edge
        pen = QPen(QColor(color), width)
        pen.setCapStyle(Qt.FlatCap)
        painter.setPen(pen)
        inset = width // 2
        if side == "top":
            y = rect.top() + inset
            painter.drawLine(rect.left(), y, rect.right(), y)
        elif side == "bottom":
            y = rect.bottom() - inset
            painter.drawLine(rect.left(), y, rect.right(), y)
        elif side == "left":
            x = rect.left() + inset
            painter.drawLine(x, rect.top(), x, rect.bottom())
        else:
            x = rect.right() - inset
            painter.drawLine(x, rect.top(), x, rect.bottom())


def _is_number(value: Any) -> bool:
    return isinstance(value, (int, float)) and not isinstance(value, bool)


class SpreadsheetView(QTableWidget):
    dest_moved = Signal(str, int, int)
    dests_moved = Signal(list)
    group_selected = Signal(list)
    range_dropped = Signal(dict, int, int)

    def __init__(self, role: str, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.role = role
        self.sheet_name = "Sheet1"
        self.styles: dict[tuple[int, int], CellStyle] = {}
        self.mappings: list[Mapping] = []
        self.active_id: str | None = None
        self._press_inside_selection = False
        self._press_pos = None
        self._moving_id: str | None = None
        self._grab = (0, 0)
        self._ignore_drag_residue = False
        self.select_move = False
        self.selected_ids: list[str] = []
        self._moving_group = False
        self._group_press: tuple[int, int] | None = None
        self._group_origins: dict[str, tuple[int, int]] = {}

        self.setItemDelegate(RangeDelegate(self))
        self.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.setDragEnabled(False)
        self.setSelectionMode(
            QAbstractItemView.ContiguousSelection if role == "source" else QAbstractItemView.NoSelection
        )
        self.setSelectionBehavior(QAbstractItemView.SelectItems)
        if role == "output":
            self.setDragDropMode(QAbstractItemView.DropOnly)
        else:
            self.setDragDropMode(QAbstractItemView.NoDragDrop)
        self.setShowGrid(False)
        self.setFocusPolicy(Qt.StrongFocus)
        self.verticalHeader().setDefaultSectionSize(26)
        self.horizontalHeader().setDefaultSectionSize(86)
        self.verticalHeader().setSectionResizeMode(QHeaderView.Fixed)
        self.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.verticalHeader().setHighlightSections(False)
        self.horizontalHeader().setHighlightSections(False)
        self.setStyleSheet(
            """
            QTableWidget { background: #fbfaf6; gridline-color: #e1dacd; }
            QTableWidget::item:selected { background: transparent; }
            QHeaderView::section {
                background: #eef1ea;
                color: #445044;
                border: 1px solid #d7d0c4;
                font-weight: 600;
                padding: 0 4px;
            }
            """
        )
        if role == "output":
            self.setAcceptDrops(True)
            self.setMouseTracking(True)

    def range_of(self, mapping: Mapping):
        return mapping.source if self.role == "source" else mapping.dest

    def mapping_index(self, mapping_id: str) -> int:
        for i, mapping in enumerate(self.mappings, start=1):
            if mapping.id == mapping_id:
                return i
        return 0

    def set_mappings(self, mappings: list[Mapping], active_id: str | None) -> None:
        self.mappings = mappings
        self.active_id = active_id
        self.viewport().update()

    def _cell_flags(self):
        flags = Qt.ItemIsEnabled
        if self.role == "source" or (self.role == "output" and self.select_move):
            flags |= Qt.ItemIsSelectable
        return flags

    def _refresh_item_flags(self) -> None:
        flags = self._cell_flags()
        for row in range(self.rowCount()):
            for col in range(self.columnCount()):
                item = self.item(row, col)
                if item is not None:
                    item.setFlags(flags)

    def set_select_move(self, enabled: bool) -> None:
        self.select_move = enabled
        if not enabled:
            self.selected_ids = []
            self._moving_group = False
            self._group_origins = {}
            self._group_press = None
            self.clearSelection()
            self.setSelectionMode(QAbstractItemView.NoSelection)
            self.viewport().setCursor(Qt.ArrowCursor)
        else:
            self.setSelectionMode(QAbstractItemView.ContiguousSelection)
        self._refresh_item_flags()
        self.viewport().update()

    def mappings_in_selection(self) -> list[str]:
        ranges = self.selectedRanges()
        if not ranges:
            return []
        ids: list[str] = []
        for mapping in self.mappings:
            dest = mapping.dest.normalized()
            if dest.sheet != self.sheet_name:
                continue
            for rng in ranges:
                if dest.overlaps_rect(rng.topRow(), rng.leftColumn(), rng.bottomRow(), rng.rightColumn()):
                    ids.append(mapping.id)
                    break
        return ids

    def _put_value(self, row: int, col: int, value: Any) -> None:
        display: Any = "" if value is None else value
        item = self.item(row, col)
        if item is None:
            item = QTableWidgetItem()
            item.setFlags(self._cell_flags())
            self.setItem(row, col, item)
        if item.data(Qt.DisplayRole) != display:
            item.setData(Qt.DisplayRole, display)

    def set_sheet(self, sheet: SheetData) -> None:
        self.sheet_name = sheet.name
        self.styles = sheet.styles
        same_shape = self.rowCount() == sheet.rows and self.columnCount() == sheet.cols and self.item(0, 0) is not None
        if same_shape:
            for r, row in enumerate(sheet.values):
                for c, value in enumerate(row):
                    self._put_value(r, c, value)
            self.viewport().update()
            return
        current_r, current_c = self.currentRow(), self.currentColumn()
        scroll_h = self.horizontalScrollBar().value()
        scroll_v = self.verticalScrollBar().value()
        self.clearContents()
        self.setRowCount(sheet.rows)
        self.setColumnCount(sheet.cols)
        self.setHorizontalHeaderLabels([col_letter(i) for i in range(sheet.cols)])
        self.setVerticalHeaderLabels([str(i + 1) for i in range(sheet.rows)])
        for r, row in enumerate(sheet.values):
            for c, value in enumerate(row):
                self._put_value(r, c, value)
        if 0 <= current_r < sheet.rows and 0 <= current_c < sheet.cols:
            self.setCurrentCell(current_r, current_c)
        self.horizontalScrollBar().setValue(scroll_h)
        self.verticalScrollBar().setValue(scroll_v)

    def current_range_dict(self) -> dict[str, Any] | None:
        ranges = self.selectedRanges()
        if not ranges:
            return None
        rng = ranges[0]
        return {
            "sheet": self.sheet_name,
            "r1": rng.topRow(),
            "c1": rng.leftColumn(),
            "r2": rng.bottomRow(),
            "c2": rng.rightColumn(),
        }

    def select_range(self, r1: int, c1: int, r2: int, c2: int) -> None:
        self.clearSelection()
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                item = self.item(r, c)
                if item is not None:
                    item.setSelected(True)
        self.scrollToItem(self.item(r1, c1) or self.item(0, 0))

    def _mapping_at(self, row: int, col: int) -> Mapping | None:
        found = None
        for mapping in self.mappings:
            if mapping.dest.sheet == self.sheet_name and mapping.dest.contains(row, col):
                found = mapping
        return found

    def _end_pointer_gesture(self) -> None:
        self._press_inside_selection = False
        self._press_pos = None
        self._moving_id = None
        self._moving_group = False
        self._group_press = None
        self._group_origins = {}
        grabber = QWidget.mouseGrabber()
        if grabber is self or grabber is self.viewport():
            grabber.releaseMouse()
        if hasattr(self, "setState"):
            self.setState(QAbstractItemView.NoState)

    def mousePressEvent(self, event) -> None:  # type: ignore[override]
        self._ignore_drag_residue = False
        index = self.indexAt(event.position().toPoint())
        if event.button() == Qt.LeftButton and self.role == "output" and index.isValid():
            if self.select_move:
                mapping = self._mapping_at(index.row(), index.column())
                if mapping is not None and mapping.id in self.selected_ids:
                    self._moving_group = True
                    self._group_press = (index.row(), index.column())
                    self._group_origins = {}
                    for item in self.mappings:
                        if item.id in self.selected_ids:
                            dest = item.dest.normalized()
                            self._group_origins[item.id] = (dest.r1, dest.c1)
                    self.viewport().setCursor(Qt.ClosedHandCursor)
                    event.accept()
                    return
                super().mousePressEvent(event)
                return
            mapping = self._mapping_at(index.row(), index.column())
            if mapping is not None:
                self._moving_id = mapping.id
                self._grab = (index.row() - mapping.dest.r1, index.column() - mapping.dest.c1)
                self.viewport().setCursor(Qt.ClosedHandCursor)
                event.accept()
                return
        if event.button() == Qt.LeftButton and self.role == "source" and index.isValid():
            self._press_pos = event.position().toPoint()
            item = self.item(index.row(), index.column())
            self._press_inside_selection = item is not None and item.isSelected()
            if self._press_inside_selection:
                event.accept()
                return
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event) -> None:  # type: ignore[override]
        if self._ignore_drag_residue:
            event.accept()
            return
        pos = event.position().toPoint()
        index = self.indexAt(pos)
        if (
            self.role == "output"
            and self.select_move
            and self._moving_group
            and event.buttons() & Qt.LeftButton
            and index.isValid()
            and self._group_press is not None
        ):
            delta_r = index.row() - self._group_press[0]
            delta_c = index.column() - self._group_press[1]
            if self._group_origins:
                min_r = min(origin[0] for origin in self._group_origins.values())
                min_c = min(origin[1] for origin in self._group_origins.values())
                delta_r = max(-min_r, delta_r)
                delta_c = max(-min_c, delta_c)
            payload = [
                (mapping_id, origin[0] + delta_r, origin[1] + delta_c)
                for mapping_id, origin in self._group_origins.items()
            ]
            self.dests_moved.emit(payload)
            event.accept()
            return
        if self.role == "output" and self.select_move and not self._moving_group:
            if not (event.buttons() & Qt.LeftButton):
                mapping = self._mapping_at(index.row(), index.column()) if index.isValid() else None
                grouped = mapping is not None and mapping.id in self.selected_ids
                self.viewport().setCursor(Qt.OpenHandCursor if grouped else Qt.ArrowCursor)
            super().mouseMoveEvent(event)
            return
        if self.role == "output" and self._moving_id and event.buttons() & Qt.LeftButton and index.isValid():
            r1 = max(0, index.row() - self._grab[0])
            c1 = max(0, index.column() - self._grab[1])
            self.dest_moved.emit(self._moving_id, r1, c1)
            event.accept()
            return
        if self.role == "output" and not (event.buttons() & Qt.LeftButton):
            mapping = self._mapping_at(index.row(), index.column()) if index.isValid() else None
            self.viewport().setCursor(Qt.OpenHandCursor if mapping else Qt.ArrowCursor)
        if self.role == "source" and self._press_inside_selection:
            if (
                event.buttons() & Qt.LeftButton
                and self._press_pos is not None
                and (pos - self._press_pos).manhattanLength() >= QApplication.startDragDistance()
            ):
                self._press_inside_selection = False
                self._start_range_drag()
            event.accept()
            return
        super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event) -> None:  # type: ignore[override]
        if self._ignore_drag_residue:
            self._ignore_drag_residue = False
            self._end_pointer_gesture()
            event.accept()
            return
        if self._moving_group:
            self._end_pointer_gesture()
            self.viewport().setCursor(Qt.OpenHandCursor)
            event.accept()
            return
        if self.role == "output" and self.select_move:
            super().mouseReleaseEvent(event)
            self.selected_ids = self.mappings_in_selection()
            self.group_selected.emit(self.selected_ids)
            self.viewport().update()
            event.accept()
            return
        if self._moving_id:
            self._end_pointer_gesture()
            self.viewport().setCursor(Qt.ArrowCursor)
            event.accept()
            return
        if self.role == "source" and self._press_inside_selection:
            self._press_inside_selection = False
            super().mousePressEvent(event)
            super().mouseReleaseEvent(event)
            self._end_pointer_gesture()
            return
        self._end_pointer_gesture()
        super().mouseReleaseEvent(event)

    def _start_range_drag(self) -> None:
        payload = self.current_range_dict()
        if payload is None:
            self._end_pointer_gesture()
            return
        mime = QMimeData()
        mime.setData(MIME_RANGE, QByteArray(json.dumps(payload).encode("utf-8")))
        mime.setText(range_a1(payload["r1"], payload["c1"], payload["r2"], payload["c2"]))
        drag = QDrag(self)
        drag.setMimeData(mime)
        pix = QPixmap(168, 30)
        pix.fill(QColor("#1e3a2b"))
        painter = QPainter(pix)
        painter.setPen(QColor("#ffffff"))
        painter.drawText(pix.rect(), Qt.AlignCenter, f"{payload['sheet']} {mime.text()}")
        painter.end()
        drag.setPixmap(pix)
        drag.setHotSpot(pix.rect().center())
        try:
            drag.exec(Qt.CopyAction)
        finally:
            self._end_pointer_gesture()
            self._ignore_drag_residue = True

    def dragEnterEvent(self, event) -> None:  # type: ignore[override]
        if self.role == "output" and event.mimeData().hasFormat(MIME_RANGE):
            event.acceptProposedAction()
        else:
            event.ignore()

    def dragMoveEvent(self, event) -> None:  # type: ignore[override]
        if self.role == "output" and event.mimeData().hasFormat(MIME_RANGE):
            event.acceptProposedAction()
        else:
            event.ignore()

    def dropEvent(self, event) -> None:  # type: ignore[override]
        if self.role != "output" or not event.mimeData().hasFormat(MIME_RANGE):
            event.ignore()
            return
        payload = json.loads(bytes(event.mimeData().data(MIME_RANGE)).decode("utf-8"))
        index = self.indexAt(event.position().toPoint())
        if not index.isValid():
            event.ignore()
            return
        self.range_dropped.emit(payload, index.row(), index.column())
        event.acceptProposedAction()


class SheetPane(QWidget):
    sheet_changed = Signal(str)
    open_clicked = Signal()
    export_clicked = Signal()

    def __init__(self, title: str, role: str, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.role = role
        self.title_label = QLabel(title)
        self.title_label.setObjectName("paneTitle")
        self.file_label = QLabel("未読込")
        self.file_label.setObjectName("paneFile")
        self.table = SpreadsheetView(role)
        self.tabs = QTabBar()
        self.tabs.setObjectName("sheetTabs")
        self.tabs.setDrawBase(False)
        self.tabs.setExpanding(False)
        self.tabs.setDocumentMode(True)
        self.tabs.setUsesScrollButtons(True)
        self.tabs.setFocusPolicy(Qt.NoFocus)
        self.tabs.setElideMode(Qt.ElideRight)
        self.tabs.currentChanged.connect(self._on_tab)

        buttons = QHBoxLayout()
        buttons.setContentsMargins(0, 4, 0, 0)
        buttons.setSpacing(6)
        if role == "source":
            self.btn_open = QPushButton("データ元を開く")
            self.btn_open.setObjectName("paneAction")
            self.btn_open.clicked.connect(lambda: self.open_clicked.emit())
            buttons.addWidget(self.btn_open)
            buttons.addStretch(1)
        else:
            self.btn_open = QPushButton("出力フォーマットを設定")
            self.btn_open.setObjectName("paneAction")
            self.btn_open.clicked.connect(lambda: self.open_clicked.emit())
            self.btn_export = QPushButton("出力ファイルを保存")
            self.btn_export.setObjectName("paneAction")
            self.btn_export.clicked.connect(lambda: self.export_clicked.emit())
            buttons.addWidget(self.btn_export)
            buttons.addStretch(1)
            buttons.addWidget(self.btn_open)

        head = QWidget()
        head.setObjectName("paneHead")
        head_layout = QVBoxLayout(head)
        head_layout.setContentsMargins(10, 8, 10, 8)
        head_layout.setSpacing(2)
        head_layout.addWidget(self.title_label)
        head_layout.addWidget(self.file_label)
        head_layout.addLayout(buttons)

        self.placeholder = QLabel("フィールドにデータを読み込んでください")
        self.placeholder.setObjectName("panePlaceholder")
        self.placeholder.setAlignment(Qt.AlignCenter)
        self.placeholder.setWordWrap(True)
        self.has_sheet = False

        self.tab_strip = QWidget()
        self.tab_strip.setObjectName("sheetTabStrip")
        tab_row = QHBoxLayout(self.tab_strip)
        tab_row.setContentsMargins(8, 0, 8, 0)
        tab_row.setSpacing(0)
        tab_row.addWidget(self.tabs, 0, Qt.AlignLeft | Qt.AlignBottom)
        tab_row.addStretch(1)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)
        layout.addWidget(head)
        layout.addWidget(self.tab_strip)
        layout.addWidget(self.placeholder, 1)
        layout.addWidget(self.table, 1)
        self.table.hide()
        self.tab_strip.hide()
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        tint = "#eef4fb" if role == "output" else "#eef6ef"
        self.setStyleSheet(
            f"""
            QWidget#paneHead {{ background: {tint}; }}
            QWidget#sheetTabStrip {{
                background: {tint};
                border-bottom: 1px solid #d7d0c4;
            }}
            QLabel#paneTitle {{ font-weight: 700; font-size: 12px; letter-spacing: 1px; }}
            QLabel#paneFile {{ color: #5c6b5c; font-size: 11px; }}
            QLabel#panePlaceholder {{
                background: #f4f1ea;
                color: #8a8074;
                font-size: 14px;
                padding: 24px;
            }}
            QPushButton#paneAction {{
                background: #2d5a3d;
                color: #f7fff8;
                border: none;
                border-radius: 5px;
                padding: 5px 10px;
                font-size: 12px;
            }}
            QPushButton#paneAction:hover {{ background: #3c9a5c; }}
            QPushButton#paneAction:disabled {{ background: #9aa89d; color: #eef3ee; }}
            QTabBar#sheetTabs {{
                background: transparent;
            }}
            QTabBar#sheetTabs::tab {{
                background: #dceee1;
                color: #3d5c47;
                border: none;
                border-bottom: 2px solid transparent;
                padding: 8px 14px 7px 14px;
                margin: 0 1px;
                font-size: 12px;
                min-width: 36px;
            }}
            QTabBar#sheetTabs::tab:hover {{
                color: #1e3a2b;
                background: #cfe8d6;
            }}
            QTabBar#sheetTabs::tab:selected {{
                color: #1e3a2b;
                font-weight: 700;
                background: #c5e6d0;
                border-bottom: 2px solid #7eb892;
            }}
            QTabBar#sheetTabs QToolButton {{
                background: transparent;
                border: none;
                color: #2d5a3d;
            }}
            """
        )

    def set_filename(self, name: str) -> None:
        self.file_label.setText(name)

    def set_empty(self, filename: str = "ファイル未読込") -> None:
        self.has_sheet = False
        self.set_filename(filename)
        self.placeholder.show()
        self.table.hide()
        self.tab_strip.hide()
        self.tabs.blockSignals(True)
        while self.tabs.count():
            self.tabs.removeTab(0)
        self.tabs.blockSignals(False)
        self.table.set_mappings([], None)

    def set_workbook(self, book: WorkbookData, mappings: list[Mapping], active_id: str | None) -> None:
        self.has_sheet = True
        self.placeholder.hide()
        self.table.show()
        self.tab_strip.show()
        self.tabs.blockSignals(True)
        try:
            while self.tabs.count():
                self.tabs.removeTab(0)
            for name in book.sheet_names:
                self.tabs.addTab(name)
            index = book.sheet_names.index(book.active) if book.active in book.sheet_names else 0
            self.tabs.setCurrentIndex(index)
        finally:
            self.tabs.blockSignals(False)
        self.table.set_sheet(book.sheet(book.active))
        self.table.set_mappings(mappings, active_id)

    def _on_tab(self, index: int) -> None:
        if 0 <= index < self.tabs.count():
            self.sheet_changed.emit(self.tabs.tabText(index))
