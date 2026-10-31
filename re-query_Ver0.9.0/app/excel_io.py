from __future__ import annotations

import colorsys
import csv
import json
from datetime import datetime
from pathlib import Path
from typing import Any

from openpyxl import Workbook, load_workbook
from openpyxl.styles.colors import COLOR_INDEX
from openpyxl.xml.functions import fromstring

from .models import (
    MIN_COLS,
    MIN_ROWS,
    CellStyle,
    Mapping,
    SheetData,
    WorkbookData,
    empty_sheet,
)

_THEME_NAMES = (
    "lt1",
    "dk1",
    "lt2",
    "dk2",
    "accent1",
    "accent2",
    "accent3",
    "accent4",
    "accent5",
    "accent6",
    "hlink",
    "folHlink",
)
_DEFAULT_THEME = (
    "FFFFFF",
    "000000",
    "EEECE1",
    "1F497D",
    "4F81BD",
    "C0504D",
    "9BBB59",
    "8064A2",
    "4BACC6",
    "F79646",
    "0000FF",
    "800080",
)
_BORDER_WIDTH = {
    "hair": 1,
    "thin": 1,
    "dotted": 1,
    "dashed": 1,
    "dashDot": 1,
    "dashDotDot": 1,
    "slantDashDot": 1,
    "medium": 2,
    "mediumDashed": 2,
    "mediumDashDot": 2,
    "mediumDashDotDot": 2,
    "thick": 3,
    "double": 3,
}


def _normalize_value(value: Any) -> Any:
    if value is None:
        return None
    if isinstance(value, datetime):
        if value.hour == 0 and value.minute == 0 and value.second == 0:
            return value.strftime("%Y/%m/%d")
        return value.strftime("%Y/%m/%d %H:%M")
    if isinstance(value, bool):
        return value
    if isinstance(value, float) and value.is_integer():
        return int(value)
    return value


def _local_tag(tag: str) -> str:
    return tag.split("}", 1)[-1]


def _hex6(raw: Any, *, allow_transparent_black: bool) -> str | None:
    if raw is None:
        return None
    text = str(raw).strip().upper().removeprefix("#")
    if text.startswith("VALUES MUST BE") or text in {"", "NONE"}:
        return None
    if len(text) == 8:
        alpha, rgb = text[:2], text[2:]
        if rgb == "000000" and alpha == "00":
            return "000000" if allow_transparent_black else None
        return rgb
    if len(text) == 6 and all(ch in "0123456789ABCDEF" for ch in text):
        return text
    return None


def apply_tint(rgb_hex: str, tint: float) -> str:
    rgb_hex = rgb_hex[-6:]
    red = int(rgb_hex[0:2], 16) / 255.0
    green = int(rgb_hex[2:4], 16) / 255.0
    blue = int(rgb_hex[4:6], 16) / 255.0
    hue, light, sat = colorsys.rgb_to_hls(red, green, blue)
    if tint < 0:
        light *= 1.0 + tint
    else:
        light = light * (1.0 - tint) + tint
    light = min(1.0, max(0.0, light))
    red, green, blue = colorsys.hls_to_rgb(hue, light, sat)
    return f"#{int(red * 255 + 0.5):02X}{int(green * 255 + 0.5):02X}{int(blue * 255 + 0.5):02X}"


def _scheme_hex(node: Any) -> str | None:
    children = list(node)
    if not children:
        return None
    inner = children[0]
    raw = inner.get("lastClr") or inner.get("val")
    parsed = _hex6(raw, allow_transparent_black=True)
    return parsed


def theme_colors_from_workbook(workbook: Any) -> list[str]:
    colors = list(_DEFAULT_THEME)
    xml = getattr(workbook, "loaded_theme", None)
    if not xml:
        return colors
    try:
        root = fromstring(xml)
    except Exception:
        return colors
    scheme = None
    for element in root.iter():
        if _local_tag(element.tag) == "clrScheme":
            scheme = element
            break
    if scheme is None:
        return colors
    by_name: dict[str, str] = {}
    for child in list(scheme):
        hex_value = _scheme_hex(child)
        if hex_value:
            by_name[_local_tag(child.tag)] = hex_value
    return [by_name.get(name, colors[index]) for index, name in enumerate(_THEME_NAMES)]


def _color_type(color: Any) -> str | None:
    try:
        value = color.type
    except Exception:
        return None
    if value in (None, "auto"):
        return value
    return str(value)


def _safe_int(value: Any) -> int | None:
    try:
        return int(value)
    except Exception:
        return None


def resolve_color(color: Any, theme: list[str], *, for_border: bool = False) -> str | None:
    if color is None:
        return "#000000" if for_border else None
    ctype = _color_type(color)
    try:
        tint = float(color.tint or 0.0)
    except Exception:
        tint = 0.0
    hex6 = None
    if ctype == "theme":
        index = _safe_int(getattr(color, "theme", None))
        if index is not None and 0 <= index < len(theme):
            hex6 = theme[index]
    elif ctype == "indexed":
        index = _safe_int(getattr(color, "indexed", None))
        if index == 64:
            hex6 = "000000"
        elif index == 65:
            hex6 = "FFFFFF"
        elif index is not None and 0 <= index < len(COLOR_INDEX):
            hex6 = _hex6(COLOR_INDEX[index], allow_transparent_black=True)
    else:
        try:
            hex6 = _hex6(color.rgb, allow_transparent_black=for_border)
        except Exception:
            hex6 = None
    if not hex6:
        return "#000000" if for_border else None
    if tint:
        return apply_tint(hex6, tint)
    return f"#{hex6.upper()}"


def _border_edge(side: Any, theme: list[str]) -> tuple[int, str] | None:
    if side is None or not getattr(side, "style", None):
        return None
    width = _BORDER_WIDTH.get(str(side.style), 1)
    color = resolve_color(side.color, theme, for_border=True) or "#000000"
    return (width, color)


def read_cell_style(cell: Any, theme: list[str]) -> CellStyle | None:
    fill_hex = None
    fill = getattr(cell, "fill", None)
    if fill is not None and fill.patternType == "solid":
        fill_hex = resolve_color(fill.fgColor, theme)
        if fill_hex in {"#FFFFFF", "#FFF"}:
            fill_hex = None

    font = getattr(cell, "font", None)
    bold = bool(font.bold) if font is not None else False
    font_color = None
    if font is not None and font.color is not None:
        color = font.color
        is_default = (
            _color_type(color) == "theme"
            and _safe_int(getattr(color, "theme", None)) == 1
            and not (getattr(color, "tint", None) or 0)
        )
        if not is_default:
            font_color = resolve_color(color, theme)
            if font_color == "#000000":
                font_color = None

    border = getattr(cell, "border", None)
    top = right = bottom = left = None
    if border is not None:
        left = _border_edge(border.left, theme)
        right = _border_edge(border.right, theme)
        top = _border_edge(border.top, theme)
        bottom = _border_edge(border.bottom, theme)

    style = CellStyle(
        fill=fill_hex,
        font_color=font_color,
        bold=bold,
        border_top=top,
        border_right=right,
        border_bottom=bottom,
        border_left=left,
    )
    return None if style.is_empty() else style


def _sheet_from_rows(
    name: str,
    rows: list[list[Any]],
    styles: dict[tuple[int, int], CellStyle] | None = None,
) -> SheetData:
    max_cols = max((len(row) for row in rows), default=0)
    cols = max(MIN_COLS, max_cols + 4)
    height = max(MIN_ROWS, len(rows) + 10)
    data = empty_sheet(name, height, cols)
    for r, row in enumerate(rows):
        for c, value in enumerate(row):
            data.values[r][c] = _normalize_value(value)
    if styles:
        data.styles = styles
    return data


def load_csv(path: str | Path) -> WorkbookData:
    path = Path(path)
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.reader(handle)
        rows = [list(row) for row in reader]
    sheet = _sheet_from_rows("Sheet1", rows)
    return WorkbookData(
        filename=path.name,
        path=str(path),
        sheet_names=["Sheet1"],
        sheets={"Sheet1": sheet},
        active="Sheet1",
        from_format=False,
    )


def load_xlsx(path: str | Path) -> WorkbookData:
    path = Path(path)
    cached = load_workbook(path, data_only=True)
    formulas = load_workbook(path, data_only=False)
    try:
        names = list(cached.sheetnames)
        current = cached.active
        active_name = current.title if current is not None and current.title in names else names[0]
        theme = theme_colors_from_workbook(formulas)
        sheets: dict[str, SheetData] = {}
        for name in names:
            cached_ws = cached[name]
            formula_ws = formulas[name]
            max_r = max(cached_ws.max_row or 1, formula_ws.max_row or 1)
            max_c = max(cached_ws.max_column or 1, formula_ws.max_column or 1)
            merged: list[list[Any]] = []
            styles: dict[tuple[int, int], CellStyle] = {}
            for row in range(1, max_r + 1):
                line: list[Any] = []
                for col in range(1, max_c + 1):
                    formula_cell = formula_ws.cell(row, col)
                    value = cached_ws.cell(row, col).value
                    if value is None:
                        value = formula_cell.value
                    line.append(value)
                    style = read_cell_style(formula_cell, theme)
                    if style is not None:
                        styles[(row - 1, col - 1)] = style
                merged.append(line)
            sheets[name] = _sheet_from_rows(name, merged, styles)
    finally:
        cached.close()
        formulas.close()
    return WorkbookData(
        filename=path.name,
        path=str(path),
        sheet_names=names,
        sheets=sheets,
        active=active_name,
        from_format=False,
    )


def load_workbook_file(path: str | Path) -> WorkbookData:
    path = Path(path)
    if path.suffix.lower() == ".csv":
        return load_csv(path)
    return load_xlsx(path)


def _used_aoa(sheet: SheetData) -> list[list[Any]]:
    last_r = -1
    last_c = -1
    for r, row in enumerate(sheet.values):
        for c, value in enumerate(row):
            if value is not None and value != "":
                last_r = max(last_r, r)
                last_c = max(last_c, c)
    if last_r < 0:
        return []
    rows: list[list[Any]] = []
    for row in sheet.values[: last_r + 1]:
        rows.append(["" if value is None else value for value in row[: last_c + 1]])
    return rows


def save_output_csv(path: str | Path, displayed: WorkbookData) -> None:
    path = Path(path)
    sheet = displayed.sheet()
    rows = _used_aoa(sheet)
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.writer(handle)
        writer.writerows(rows)


def save_output(
    path: str | Path,
    displayed: WorkbookData,
    mappings: list[Mapping],
    source: WorkbookData | None,
    format_path: str | None,
) -> None:
    path = Path(path)
    if path.suffix.lower() == ".csv":
        save_output_csv(path, displayed)
        return
    save_output_xlsx(path, displayed, mappings, source, format_path)


def save_output_xlsx(
    path: str | Path,
    displayed: WorkbookData,
    mappings: list[Mapping],
    source: WorkbookData | None,
    format_path: str | None,
) -> None:
    path = Path(path)
    if format_path and Path(format_path).exists() and Path(format_path).suffix.lower() != ".csv":
        wb = load_workbook(format_path)
        if source is not None:
            for mapping in mappings:
                src_sheet = source.sheets.get(mapping.source.sheet)
                if src_sheet is None:
                    continue
                dest_name = mapping.dest.sheet
                if dest_name not in wb.sheetnames:
                    wb.create_sheet(dest_name)
                ws = wb[dest_name]
                src = mapping.source.normalized()
                dest = mapping.dest.normalized()
                for row_i in range(src.row_count):
                    for col_i in range(src.col_count):
                        value = src_sheet.get(src.r1 + row_i, src.c1 + col_i)
                        ws.cell(dest.r1 + row_i + 1, dest.c1 + col_i + 1, value)
        wb.save(path)
        wb.close()
        return

    wb = Workbook()
    default = wb.active
    first = True
    for name in displayed.sheet_names:
        ws = default if first else wb.create_sheet(name)
        if first:
            ws.title = name[:31]
            first = False
        sheet = displayed.sheets[name]
        for r, row in enumerate(sheet.values, start=1):
            for c, value in enumerate(row, start=1):
                if value is not None and value != "":
                    ws.cell(r, c, value)
    wb.save(path)
    wb.close()


def dump_mappings(mappings: list[Mapping], format_path: str | None) -> str:
    payload = {
        "format_file": format_path,
        "mappings": [item.to_dict() for item in mappings],
    }
    return json.dumps(payload, ensure_ascii=False, indent=2)


def load_mappings_json(text: str) -> tuple[list[Mapping], str | None]:
    payload = json.loads(text)
    mappings = [Mapping.from_dict(item) for item in payload.get("mappings", [])]
    format_path = payload.get("format_file") or None
    if format_path == "":
        format_path = None
    return mappings, format_path


def resolve_format_path(stored: str | None, json_dir: Path | None = None) -> Path | None:
    if not stored:
        return None
    path = Path(stored)
    if path.is_file():
        return path
    if json_dir is not None:
        relative = json_dir / stored
        if relative.is_file():
            return relative
        by_name = json_dir / path.name
        if by_name.is_file():
            return by_name
    return None
