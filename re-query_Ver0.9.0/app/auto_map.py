from __future__ import annotations

from typing import Any, Hashable

from .models import CellRange, Mapping, SheetData, WorkbookData, next_color

MODE_NUMBERS = "numbers"
MODE_ALL = "all"


def strip_spaces(value: Any) -> str:
    return "".join(str(value).split())


def as_number(value: Any) -> int | float | None:
    if value is None or value == "":
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, int):
        return value
    if isinstance(value, float):
        return int(value) if value.is_integer() else value
    text = strip_spaces(value).replace(",", "").replace("，", "")
    if not text:
        return None
    try:
        if "." in text:
            number = float(text)
            return int(number) if number.is_integer() else number
        return int(text)
    except ValueError:
        return None


def match_key(value: Any, mode: str) -> Hashable | None:
    if value is None or value == "":
        return None
    if mode == MODE_NUMBERS:
        return as_number(value)
    text = strip_spaces(value)
    return text if text else None


def _index_source(book: WorkbookData, mode: str) -> dict[Hashable, list[tuple[str, int, int]]]:
    index: dict[Hashable, list[tuple[str, int, int]]] = {}
    for sheet_name in book.sheet_names:
        sheet = book.sheets[sheet_name]
        for row, values in enumerate(sheet.values):
            for col, value in enumerate(values):
                key = match_key(value, mode)
                if key is None:
                    continue
                index.setdefault(key, []).append((sheet_name, row, col))
    return index


def _same_as_blank(blank: SheetData | None, row: int, col: int, key: Hashable, mode: str) -> bool:
    if blank is None:
        return False
    blank_key = match_key(blank.get(row, col), mode)
    return blank_key is not None and blank_key == key


def build_auto_mappings(
    source: WorkbookData,
    filled: WorkbookData,
    blank: WorkbookData | None,
    mode: str,
) -> list[Mapping]:
    index = _index_source(source, mode)
    mappings: list[Mapping] = []
    for sheet_name in filled.sheet_names:
        filled_sheet = filled.sheets[sheet_name]
        blank_sheet = blank.sheets.get(sheet_name) if blank is not None else None
        for row, values in enumerate(filled_sheet.values):
            for col, value in enumerate(values):
                key = match_key(value, mode)
                if key is None or _same_as_blank(blank_sheet, row, col, key, mode):
                    continue
                hits = index.get(key) or []
                if not hits:
                    continue
                ranges = [
                    CellRange(src_sheet, src_row, src_col, src_row, src_col)
                    for src_sheet, src_row, src_col in hits
                ]
                mapping = Mapping.create(
                    next_color(len(mappings)),
                    ranges[0],
                    CellRange(sheet_name, row, col, row, col),
                )
                mapping.candidates = ranges
                mapping.match_count = len(ranges)
                mappings.append(mapping)
    return mappings
