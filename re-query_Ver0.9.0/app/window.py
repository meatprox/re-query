from __future__ import annotations

from datetime import date
from pathlib import Path

from PySide6.QtCore import QSettings, Qt
from PySide6.QtWidgets import (
    QApplication,
    QFileDialog,
    QLabel,
    QMainWindow,
    QMessageBox,
    QPlainTextEdit,
    QSplitter,
    QStatusBar,
)

from .auto_map_dialog import AutoMapDialog
from .excel_io import dump_mappings, load_mappings_json, load_workbook_file, resolve_format_path, save_output
from .mapping_panel import MappingPanel
from .models import (
    CellRange,
    Mapping,
    WorkbookData,
    apply_mappings,
    empty_like,
    empty_workbook,
    next_color,
)
from .spreadsheet import SheetPane

EXCEL_FILTER = "Excel / CSV (*.xlsx *.xlsm *.csv);;すべてのファイル (*.*)"
SAVE_FILTER = "Excel (*.xlsx);;CSV (*.csv)"


class MainWindow(QMainWindow):
    def __init__(self, restore_last: bool = True, settings: QSettings | None = None) -> None:
        super().__init__()
        self.setWindowTitle("Excechenger — セル配置エディタ")
        self.resize(1440, 860)
        self._settings = settings or QSettings("Excechenger", "Excechenger")

        self.source: WorkbookData | None = None
        self.base_output = empty_workbook()
        self.format_path: str | None = None
        self.mappings: list[Mapping] = []
        self.active_id: str | None = None

        self.output_pane = SheetPane("出力ファイル", "output")
        self.source_pane = SheetPane("データ元", "source")
        self.mapping_panel = MappingPanel()

        self.output_pane.set_filename("ファイル未読込")
        self.source_pane.set_filename("ファイル未読込")

        splitter = QSplitter(Qt.Horizontal)
        splitter.addWidget(self.output_pane)
        splitter.addWidget(self.source_pane)
        splitter.addWidget(self.mapping_panel)
        splitter.setStretchFactor(0, 3)
        splitter.setStretchFactor(1, 3)
        splitter.setStretchFactor(2, 2)
        splitter.setSizes([560, 560, 340])
        self.setCentralWidget(splitter)

        self.source_pane.open_clicked.connect(self.open_source)
        self.output_pane.open_clicked.connect(self.open_format)
        self.output_pane.export_clicked.connect(self.export_output)
        self.mapping_panel.load_clicked.connect(self.load_mappings)
        self.mapping_panel.save_clicked.connect(self.save_mappings)
        self.mapping_panel.clear_clicked.connect(self.clear_mappings)
        self.mapping_panel.auto_map_clicked.connect(self.open_auto_map)
        self.mapping_panel.select_move_toggled.connect(self.set_select_move)
        self.output_pane.btn_export.setEnabled(False)
        self.mapping_panel.btn_save.setEnabled(False)
        self.mapping_panel.btn_clear.setEnabled(False)

        self.status = QStatusBar()
        self.setStatusBar(self.status)
        self.status_sel = QLabel()
        self.status.addPermanentWidget(self.status_sel)

        self.output_pane.sheet_changed.connect(self._on_output_sheet)
        self.source_pane.sheet_changed.connect(self._on_source_sheet)
        self.source_pane.table.itemSelectionChanged.connect(self._on_source_selection)
        self.output_pane.table.dest_moved.connect(self._on_dest_moved)
        self.output_pane.table.dests_moved.connect(self._on_dests_moved)
        self.output_pane.table.group_selected.connect(self._on_group_selected)
        self.output_pane.table.range_dropped.connect(self._on_drop_to_output)
        self.mapping_panel.range_dropped.connect(self._on_drop_to_list)
        self.mapping_panel.description_changed.connect(self._on_description)
        self.mapping_panel.removed.connect(self._remove_mapping)
        self.mapping_panel.activated.connect(self._activate_mapping)
        self.mapping_panel.chosen.connect(self._confirm_mapping)

        self._reload_panes(rebuild_tabs=True)
        if restore_last and self._restore_last_mappings():
            return
        self._set_status("データ元を開くか、配置関係を読み込んでください")

    def _remember_mapping_file(self, path: str) -> None:
        resolved = str(Path(path).resolve())
        self._settings.setValue("lastMappingFile", resolved)
        self._settings.sync()
        self.mapping_panel.set_file_name(Path(resolved).name)

    def _restore_last_mappings(self) -> bool:
        stored = self._settings.value("lastMappingFile")
        if not stored:
            return False
        path = Path(str(stored))
        if not path.is_file():
            self._set_status(f"前回の配置関係ファイルが見つかりません: {path.name}")
            return False
        try:
            mappings, format_path = load_mappings_json(path.read_text(encoding="utf-8"))
        except Exception as exc:
            self._set_status(f"前回の配置関係を開けませんでした: {exc}")
            return False
        self._apply_mapping_list(
            mappings,
            format_path=format_path,
            json_path=str(path),
            warn_format=False,
        )
        extra = "（出力フォーマットも開きました）" if self.base_output.from_format else ""
        self._set_status(f"前回の配置関係 {path.name} を開きました{extra}。")
        return True

    def displayed_output(self) -> WorkbookData:
        book = apply_mappings(self.base_output, self.source, self.mappings)
        book.active = self.base_output.active
        book.filename = self.base_output.filename
        return book

    def _output_ready(self) -> bool:
        return self.source is not None or self.base_output.from_format

    def _reload_panes(self, rebuild_tabs: bool = False, refresh_source: bool = False) -> None:
        displayed = self.displayed_output()
        if rebuild_tabs:
            if self._output_ready():
                self.output_pane.set_workbook(displayed, self.mappings, self.active_id)
            else:
                self.output_pane.set_empty()
            if self.source is not None:
                self.source_pane.set_workbook(self.source, self.mappings, self.active_id)
            else:
                self.source_pane.set_empty()
            self.mapping_panel.set_mappings(self.mappings, self.active_id)
            return
        if self._output_ready() and self.output_pane.has_sheet:
            self.output_pane.table.set_sheet(displayed.sheet())
            self.output_pane.table.set_mappings(self.mappings, self.active_id)
        if self.source is not None and self.source_pane.has_sheet:
            if refresh_source:
                self.source_pane.table.set_sheet(self.source.sheet())
            self.source_pane.table.set_mappings(self.mappings, self.active_id)
        self.mapping_panel.set_mappings(self.mappings, self.active_id)

    def _update_actions(self) -> None:
        has_map = bool(self.mappings)
        self.output_pane.btn_export.setEnabled(self.source is not None or self.base_output.from_format)
        self.mapping_panel.btn_save.setEnabled(has_map)
        self.mapping_panel.btn_clear.setEnabled(has_map)

    def _set_status(self, text: str) -> None:
        self.status.showMessage(text)
        self._update_actions()

    def _on_source_selection(self) -> None:
        payload = self.source_pane.table.current_range_dict()
        if payload is None:
            self.status_sel.setText("")
            return
        rng = CellRange.from_dict(payload)
        self.status_sel.setText(f"選択 {rng.sheet} {rng.label()}")

    def _on_output_sheet(self, name: str) -> None:
        if name in self.base_output.sheets:
            self.base_output.active = name
            self._reload_panes(rebuild_tabs=False)

    def _on_source_sheet(self, name: str) -> None:
        if self.source is not None and name in self.source.sheets:
            self.source.active = name
            self._reload_panes(rebuild_tabs=False, refresh_source=True)

    def _add_mapping(self, payload: dict, dest_r: int | None = None, dest_c: int | None = None) -> None:
        if self.source is None:
            QMessageBox.information(self, "Excechenger", "先にデータ元を開いてください。")
            return
        source = CellRange.from_dict(payload)
        if dest_r is None:
            dest_r, dest_c = source.r1, source.c1
        dest = CellRange(
            sheet=self.base_output.active,
            r1=dest_r,
            c1=dest_c,
            r2=dest_r + source.row_count - 1,
            c2=dest_c + source.col_count - 1,
        )
        mapping = Mapping.create(next_color(len(self.mappings)), source, dest)
        self.mappings.append(mapping)
        self.active_id = mapping.id
        self._reload_panes(rebuild_tabs=False)
        self._set_status(
            f"配置を登録しました: 元 {source.label()} → 出 {dest.label()}。"
            "出力側の色枠をドラッグすると貼り付け位置を動かせます。"
        )

    def _on_drop_to_list(self, payload: dict) -> None:
        self._add_mapping(payload)

    def _on_drop_to_output(self, payload: dict, row: int, col: int) -> None:
        self._add_mapping(payload, row, col)

    def _on_dest_moved(self, mapping_id: str, row: int, col: int) -> None:
        for mapping in self.mappings:
            if mapping.id == mapping_id:
                mapping.dest = mapping.dest.moved_to(row, col)
                self.active_id = mapping_id
                break
        self._reload_panes(rebuild_tabs=False)

    def _on_dests_moved(self, moves: list) -> None:
        by_id = {mapping_id: (row, col) for mapping_id, row, col in moves}
        for mapping in self.mappings:
            if mapping.id not in by_id:
                continue
            row, col = by_id[mapping.id]
            mapping.dest = mapping.dest.moved_to(row, col)
        if moves:
            self.active_id = moves[0][0]
        self._ensure_mapping_bounds()
        self._reload_panes(rebuild_tabs=False)

    def _on_group_selected(self, ids: list) -> None:
        count = len(ids)
        if count == 0:
            self._set_status("配置は選ばれていません。ドラッグして出力側の配置を囲んでください。")
            return
        self._set_status(f"{count} 件の配置を選択しました。選択した枠をドラッグするとまとめて移動できます。")

    def set_select_move(self, enabled: bool) -> None:
        self.output_pane.table.set_select_move(enabled)
        if enabled:
            self._set_status("選択して移動: 出力側をドラッグして配置を囲み、その枠をまとめて移動できます。")
        else:
            self._set_status("選択して移動を終了しました。")

    def _on_description(self, mapping_id: str, text: str) -> None:
        for mapping in self.mappings:
            if mapping.id == mapping_id:
                mapping.description = text
                break

    def _remove_mapping(self, mapping_id: str) -> None:
        self.mappings = [item for item in self.mappings if item.id != mapping_id]
        if self.active_id == mapping_id:
            self.active_id = self.mappings[-1].id if self.mappings else None
        self._reload_panes(rebuild_tabs=False)
        self._set_status("配置を削除しました")

    def _activate_mapping(self, mapping_id: str) -> None:
        mapping = next((item for item in self.mappings if item.id == mapping_id), None)
        if mapping is None:
            return
        if self.active_id == mapping_id and mapping.can_cycle():
            mapping.cycle_source()
        self.active_id = mapping_id
        self._focus_mapping(mapping)

    def _confirm_mapping(self, mapping_id: str) -> None:
        mapping = next((item for item in self.mappings if item.id == mapping_id), None)
        if mapping is None:
            return
        mapping.confirm_source()
        self.active_id = mapping_id
        self._focus_mapping(mapping)
        self._set_status(f"候補を {mapping.source.label()} に決定しました")

    def _focus_mapping(self, mapping: Mapping) -> None:
        if self.source is not None and mapping.source.sheet in self.source.sheets:
            self.source.active = mapping.source.sheet
        if mapping.dest.sheet in self.base_output.sheets:
            self.base_output.active = mapping.dest.sheet
        self._reload_panes(rebuild_tabs=True)
        src = mapping.source
        if self.source is not None and self.source_pane.has_sheet:
            self.source_pane.table.select_range(src.r1, src.c1, src.r2, src.c2)
        dest_item = self.output_pane.table.item(mapping.dest.r1, mapping.dest.c1)
        if dest_item is not None:
            self.output_pane.table.scrollToItem(dest_item)

    def _align_dest_sheets_to_output(self) -> None:
        names = self.base_output.sheet_names
        if not names:
            return
        default = self.base_output.active if self.base_output.active in self.base_output.sheets else names[0]
        for mapping in self.mappings:
            dest = mapping.dest.normalized()
            if dest.sheet not in self.base_output.sheets:
                mapping.dest = CellRange(default, dest.r1, dest.c1, dest.r2, dest.c2)

    def _set_format_book(self, book: WorkbookData, path: str) -> None:
        book.from_format = True
        self.base_output = book
        self.format_path = str(Path(path).resolve())
        self.output_pane.set_filename(f"出力フォーマット: {book.filename}")
        self._align_dest_sheets_to_output()

    def _refresh_after_file_load(self) -> None:
        self._ensure_mapping_bounds()
        self._focus_mapping_sheets()
        self._reload_panes(rebuild_tabs=True)

    def _ensure_mapping_bounds(self) -> None:
        for mapping in self.mappings:
            source_range = mapping.source.normalized()
            dest_range = mapping.dest.normalized()
            if self.source is not None and source_range.sheet in self.source.sheets:
                self.source.sheets[source_range.sheet].ensure(source_range.r2 + 1, source_range.c2 + 1)
            if dest_range.sheet in self.base_output.sheets:
                self.base_output.sheets[dest_range.sheet].ensure(dest_range.r2 + 1, dest_range.c2 + 1)
            else:
                self.base_output.sheet().ensure(dest_range.r2 + 1, dest_range.c2 + 1)

    def _focus_mapping_sheets(self) -> None:
        if not self.mappings:
            return
        mapping = next((item for item in self.mappings if item.id == self.active_id), self.mappings[0])
        if self.source is not None and mapping.source.sheet in self.source.sheets:
            self.source.active = mapping.source.sheet
        if mapping.dest.sheet in self.base_output.sheets:
            self.base_output.active = mapping.dest.sheet

    def _status_after_file_load(self, filename: str, kind: str) -> str:
        if self.mappings:
            return f"{filename} を読み込み、配置関係 {len(self.mappings)} 件を適用しました。"
        if kind == "source":
            return f"{filename} を読み込みました。セルを選択して右のリストへドラッグしてください。"
        return f"出力フォーマット {filename} を出力側に表示しています。"

    def _apply_source_book(self, book: WorkbookData) -> None:
        self.source = book
        self.source_pane.set_filename(book.filename)
        if not self.base_output.from_format:
            active = book.active
            self.base_output = empty_like(book)
            self.base_output.active = active
            self.output_pane.set_filename("未作成（空のセル）")
            self._align_dest_sheets_to_output()
        self._refresh_after_file_load()
        self._set_status(self._status_after_file_load(book.filename, "source"))

    def _apply_format_book(self, book: WorkbookData, path: str) -> None:
        kept = list(self.mappings)
        kept_id = self.active_id
        self._set_format_book(book, path)
        self.mappings = kept
        self.active_id = kept_id if any(item.id == kept_id for item in kept) else (kept[0].id if kept else None)
        self._refresh_after_file_load()
        self._set_status(self._status_after_file_load(book.filename, "format"))

    def _try_open_format_path(self, stored: str, json_dir: Path | None, warn: bool = True) -> bool:
        path = resolve_format_path(stored, json_dir)
        if path is None:
            if warn:
                QMessageBox.warning(
                    self,
                    "出力フォーマットを開けません",
                    "配置関係に記録された出力フォーマットのパスが通っていません。\n\n"
                    f"{stored}\n\n"
                    "出力フォーマットは読み込まず、配置関係のみを開きます。",
                )
            return False
        try:
            book = load_workbook_file(path)
        except Exception as exc:
            if warn:
                QMessageBox.warning(
                    self,
                    "出力フォーマットを開けません",
                    "配置関係に記録された出力フォーマットを開けませんでした。\n\n"
                    f"{path}\n{exc}\n\n"
                    "出力フォーマットは読み込まず、配置関係のみを開きます。",
                )
            return False
        self._set_format_book(book, str(path))
        return True

    def _apply_mapping_list(
        self,
        mappings: list[Mapping],
        format_path: str | None = None,
        json_path: str | None = None,
        warn_format: bool = True,
    ) -> None:
        self.mappings = mappings
        self.active_id = mappings[0].id if mappings else None
        if json_path:
            self.mapping_panel.set_file_name(Path(json_path).name)
        format_loaded = False
        if format_path:
            json_dir = Path(json_path).parent if json_path else None
            format_loaded = self._try_open_format_path(format_path, json_dir, warn=warn_format)
        self._refresh_after_file_load()
        if format_loaded:
            self._set_status(
                f"配置関係 {len(mappings)} 件を読み込み、出力フォーマットを開きました。"
            )
        elif self.source is None:
            self._set_status("配置関係を読み込みました。続けてデータ元を開くと配置が適用されます。")
        else:
            self._set_status(f"配置関係 {len(mappings)} 件を読み込み、現在のファイルに適用しました。")

    def open_source(self) -> None:
        path, _ = QFileDialog.getOpenFileName(self, "データ元を開く", "", EXCEL_FILTER)
        if not path:
            return
        try:
            book = load_workbook_file(path)
        except Exception as exc:
            QMessageBox.critical(self, "読み込みエラー", str(exc))
            return
        self._apply_source_book(book)

    def open_format(self) -> None:
        path, _ = QFileDialog.getOpenFileName(self, "出力フォーマットを設定", "", EXCEL_FILTER)
        if not path:
            return
        try:
            book = load_workbook_file(path)
        except Exception as exc:
            QMessageBox.critical(self, "読み込みエラー", str(exc))
            return
        self._apply_format_book(book, path)

    def open_auto_map(self) -> None:
        dialog = AutoMapDialog(self)
        dialog.exec()

    def load_mappings(self) -> None:
        path, _ = QFileDialog.getOpenFileName(self, "配置関係を読込", "", "JSON (*.json)")
        if not path:
            return
        try:
            mappings, format_path = load_mappings_json(Path(path).read_text(encoding="utf-8"))
        except Exception as exc:
            QMessageBox.critical(self, "読み込みエラー", str(exc))
            return
        self._apply_mapping_list(mappings, format_path=format_path, json_path=path)
        self._remember_mapping_file(path)

    def _suggested_save_name(self, fallback: str, suffix: str) -> str:
        day = date.today().strftime("%Y%m%d")
        if self.base_output.from_format:
            source = self.format_path or self.base_output.filename
            stem = Path(source).stem if source else fallback
        else:
            stem = fallback
        return f"{stem}_{day}{suffix}"

    def export_output(self) -> None:
        suggested = self._suggested_save_name("出力", ".xlsx")
        path, selected = QFileDialog.getSaveFileName(self, "出力ファイルを保存", suggested, SAVE_FILTER)
        if not path:
            return
        dest = Path(path)
        if dest.suffix.lower() not in {".xlsx", ".csv"}:
            dest = dest.with_suffix(".csv" if "csv" in selected.lower() else ".xlsx")
        try:
            save_output(
                dest,
                self.displayed_output(),
                self.mappings,
                self.source,
                self.format_path if self.base_output.from_format else None,
            )
        except Exception as exc:
            QMessageBox.critical(self, "保存エラー", str(exc))
            return
        self._set_status(f"{dest.name} に保存しました")

    def save_mappings(self) -> None:
        path, _ = QFileDialog.getSaveFileName(
            self, "配置関係を保存", self._suggested_save_name("配置関係", ".json"), "JSON (*.json)"
        )
        if not path:
            return
        format_path = None
        if self.base_output.from_format and self.format_path:
            format_path = str(Path(self.format_path).resolve())
        Path(path).write_text(
            dump_mappings(self.mappings, format_path),
            encoding="utf-8",
        )
        self._remember_mapping_file(path)
        self._set_status("配置関係を保存しました")

    def clear_mappings(self) -> None:
        if not self.mappings:
            return
        if QMessageBox.question(self, "配置をクリア", "登録した配置をすべて削除しますか？") != QMessageBox.Yes:
            return
        self.mappings = []
        self.active_id = None
        self._reload_panes(rebuild_tabs=False)
        self._set_status("配置をクリアしました")

    def keyPressEvent(self, event) -> None:  # type: ignore[override]
        if event.key() == Qt.Key_Delete:
            focused = QApplication.focusWidget()
            if isinstance(focused, QPlainTextEdit):
                super().keyPressEvent(event)
                return
            if self.active_id:
                self._remove_mapping(self.active_id)
                return
        super().keyPressEvent(event)
