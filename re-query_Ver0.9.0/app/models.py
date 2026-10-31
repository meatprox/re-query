from __future__ import annotations

from dataclasses import asdict, dataclass, field
from typing import Any
from uuid import uuid4

COLORS = [
    "#c92a2a",
    "#1971c2",
    "#2f9e44",
    "#e8590c",
    "#9c36b5",
    "#0c8599",
    "#f08c00",
    "#3b5bdb",
    "#c2255c",
    "#087f5b",
]

MIN_ROWS = 40
MIN_COLS = 16


def col_letter(index: int) -> str:
    n = index + 1
    out = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        out = chr(65 + rem) + out
    return out


def a1(row: int, col: int) -> str:
    return f"{col_letter(col)}{row + 1}"


def range_a1(r1: int, c1: int, r2: int, c2: int) -> str:
    start = a1(r1, c1)
    end = a1(r2, c2)
    return start if start == end else f"{start}:{end}"


def next_color(count: int) -> str:
    return COLORS[count % len(COLORS)]


@dataclass
class CellRange:
    sheet: str
    r1: int
    c1: int
    r2: int
    c2: int

    def normalized(self) -> CellRange:
        return CellRange(
            self.sheet,
            min(self.r1, self.r2),
            min(self.c1, self.c2),
            max(self.r1, self.r2),
            max(self.c1, self.c2),
        )

    @property
    def row_count(self) -> int:
        n = self.normalized()
        return n.r2 - n.r1 + 1

    @property
    def col_count(self) -> int:
        n = self.normalized()
        return n.c2 - n.c1 + 1

    def contains(self, row: int, col: int) -> bool:
        n = self.normalized()
        return n.r1 <= row <= n.r2 and n.c1 <= col <= n.c2

    def overlaps_rect(self, r1: int, c1: int, r2: int, c2: int) -> bool:
        n = self.normalized()
        top, bottom = min(r1, r2), max(r1, r2)
        left, right = min(c1, c2), max(c1, c2)
        return n.r1 <= bottom and n.r2 >= top and n.c1 <= right and n.c2 >= left

    def moved_to(self, r1: int, c1: int) -> CellRange:
        n = self.normalized()
        r1 = max(0, r1)
        c1 = max(0, c1)
        return CellRange(n.sheet, r1, c1, r1 + n.row_count - 1, c1 + n.col_count - 1)

    def label(self) -> str:
        n = self.normalized()
        return range_a1(n.r1, n.c1, n.r2, n.c2)

    def to_dict(self) -> dict[str, Any]:
        n = self.normalized()
        return asdict(n)

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> CellRange:
        return cls(
            sheet=str(data["sheet"]),
            r1=int(data["r1"]),
            c1=int(data["c1"]),
            r2=int(data["r2"]),
            c2=int(data["c2"]),
        ).normalized()


@dataclass
class Mapping:
    id: str
    color: str
    source: CellRange
    dest: CellRange
    description: str = ""
    match_count: int = 1
    candidates: list[CellRange] = field(default_factory=list)

    @staticmethod
    def create(color: str, source: CellRange, dest: CellRange, description: str = "") -> Mapping:
        return Mapping(
            id=uuid4().hex[:10],
            color=color,
            source=source.normalized(),
            dest=dest.normalized(),
            description=description,
        )

    def can_cycle(self) -> bool:
        return len(self.candidates) > 1

    def choice_index(self) -> int:
        current = self.source.normalized()
        for index, candidate in enumerate(self.candidates):
            if candidate.normalized() == current:
                return index
        return 0

    def cycle_source(self) -> None:
        if not self.can_cycle():
            return
        nxt = (self.choice_index() + 1) % len(self.candidates)
        self.source = self.candidates[nxt].normalized()

    def confirm_source(self) -> None:
        self.source = self.source.normalized()
        self.candidates = []
        self.match_count = 1

    def to_dict(self) -> dict[str, Any]:
        return {
            "id": self.id,
            "color": self.color,
            "source": self.source.to_dict(),
            "dest": self.dest.to_dict(),
            "description": self.description,
            "match_count": self.match_count,
            "candidates": [item.to_dict() for item in self.candidates],
        }

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> Mapping:
        try:
            match_count = int(data.get("match_count") or 1)
        except (TypeError, ValueError):
            match_count = 1
        candidates: list[CellRange] = []
        for item in data.get("candidates") or []:
            try:
                candidates.append(CellRange.from_dict(item))
            except Exception:
                continue
        if candidates:
            match_count = len(candidates)
        return cls(
            id=str(data.get("id") or uuid4().hex[:10]),
            color=str(data.get("color") or COLORS[0]),
            source=CellRange.from_dict(data["source"]),
            dest=CellRange.from_dict(data["dest"]),
            description=str(data.get("description") or ""),
            match_count=max(1, match_count),
            candidates=candidates,
        )


@dataclass(frozen=True)
class CellStyle:
    fill: str | None = None
    font_color: str | None = None
    bold: bool = False
    border_top: tuple[int, str] | None = None
    border_right: tuple[int, str] | None = None
    border_bottom: tuple[int, str] | None = None
    border_left: tuple[int, str] | None = None

    def is_empty(self) -> bool:
        return (
            self.fill is None
            and self.font_color is None
            and not self.bold
            and self.border_top is None
            and self.border_right is None
            and self.border_bottom is None
            and self.border_left is None
        )


@dataclass
class SheetData:
    name: str
    values: list[list[Any]] = field(default_factory=list)
    styles: dict[tuple[int, int], CellStyle] = field(default_factory=dict)

    @property
    def rows(self) -> int:
        return len(self.values)

    @property
    def cols(self) -> int:
        return len(self.values[0]) if self.values else 0

    def get(self, row: int, col: int) -> Any:
        if 0 <= row < self.rows and 0 <= col < self.cols:
            return self.values[row][col]
        return None

    def style_at(self, row: int, col: int) -> CellStyle | None:
        return self.styles.get((row, col))

    def ensure(self, rows: int, cols: int) -> None:
        cols = max(cols, self.cols, 1)
        for row in self.values:
            if len(row) < cols:
                row.extend([None] * (cols - len(row)))
        while len(self.values) < rows:
            self.values.append([None] * cols)

    def clone(self) -> SheetData:
        return SheetData(self.name, [row[:] for row in self.values], dict(self.styles))


@dataclass
class WorkbookData:
    filename: str
    path: str | None
    sheet_names: list[str]
    sheets: dict[str, SheetData]
    active: str
    from_format: bool = False

    def sheet(self, name: str | None = None) -> SheetData:
        return self.sheets[name or self.active]

    def clone(self, filename: str | None = None) -> WorkbookData:
        return WorkbookData(
            filename=filename or self.filename,
            path=self.path,
            sheet_names=self.sheet_names[:],
            sheets={key: value.clone() for key, value in self.sheets.items()},
            active=self.active,
            from_format=self.from_format,
        )


def empty_sheet(name: str, rows: int = MIN_ROWS, cols: int = MIN_COLS) -> SheetData:
    return SheetData(name, [[None] * cols for _ in range(rows)])


def empty_workbook(
    filename: str = "出力.xlsx",
    rows: int = MIN_ROWS,
    cols: int = MIN_COLS,
    sheet_names: list[str] | None = None,
) -> WorkbookData:
    names = sheet_names or ["Sheet1"]
    sheets = {name: empty_sheet(name, rows, cols) for name in names}
    return WorkbookData(
        filename=filename,
        path=None,
        sheet_names=names,
        sheets=sheets,
        active=names[0],
        from_format=False,
    )


def empty_like(source: WorkbookData, filename: str = "出力.xlsx") -> WorkbookData:
    sheets = {
        name: empty_sheet(name, source.sheets[name].rows, source.sheets[name].cols)
        for name in source.sheet_names
    }
    return WorkbookData(
        filename=filename,
        path=None,
        sheet_names=source.sheet_names[:],
        sheets=sheets,
        active=source.active,
        from_format=False,
    )


def apply_mappings(
    base: WorkbookData,
    source: WorkbookData | None,
    mappings: list[Mapping],
) -> WorkbookData:
    result = base.clone()
    if source is None:
        return result
    for mapping in mappings:
        src_sheet = source.sheets.get(mapping.source.sheet)
        if src_sheet is None:
            continue
        dest_name = mapping.dest.sheet if mapping.dest.sheet in result.sheets else result.active
        dest_sheet = result.sheets[dest_name]
        src = mapping.source.normalized()
        dest = mapping.dest.normalized()
        dest_sheet.ensure(dest.r2 + 1, dest.c2 + 1)
        for row_i in range(src.row_count):
            for col_i in range(src.col_count):
                dest_sheet.values[dest.r1 + row_i][dest.c1 + col_i] = src_sheet.get(
                    src.r1 + row_i, src.c1 + col_i
                )
    return result
