"""Adapter for tablib library (read+write, value-only).

tablib is a dataset-oriented library with xlsx support (via openpyxl
internally).  Like pandas and pyexcel, it exposes cell values only — no
formatting.  This adapter measures the abstraction cost of going through
tablib's Dataset/Databook model.
"""

import re
from datetime import date, datetime, time
from pathlib import Path
from typing import Any

import tablib

from excelbench.harness.adapters.base import ExcelAdapter
from excelbench.models import (
    BorderInfo,
    CellFormat,
    CellType,
    CellValue,
    LibraryInfo,
)

JSONDict = dict[str, Any]
WorkbookData = dict[str, Any]


def _get_version() -> str:
    try:
        from importlib.metadata import version

        return version("tablib")
    except Exception:
        return "unknown"


def _parse_cell_ref(cell: str) -> tuple[int, int]:
    """Parse a cell reference like 'A1' to (row_0based, col_0based)."""
    match = re.match(r"([A-Z]+)(\d+)", cell.upper())
    if not match:
        raise ValueError(f"Invalid cell reference: {cell}")
    col_str, row_str = match.groups()
    row = int(row_str) - 1
    col = 0
    for char in col_str:
        col = col * 26 + (ord(char) - ord("A") + 1)
    col -= 1
    return row, col


def _parse_cell_range(cell_range: str) -> tuple[int, int, int, int]:
    """Parse A1:B2 into (r0, c0, r1, c1) inclusive, 0-based."""
    clean = cell_range.replace("$", "").upper()
    if ":" in clean:
        a, b = clean.split(":", 1)
    else:
        a, b = clean, clean
    r0, c0 = _parse_cell_ref(a)
    r1, c1 = _parse_cell_ref(b)
    if r1 < r0:
        r0, r1 = r1, r0
    if c1 < c0:
        c0, c1 = c1, c0
    return r0, c0, r1, c1


class TablibAdapter(ExcelAdapter):
    """Adapter for tablib library (read+write, value-only).

    tablib wraps openpyxl behind its Dataset/Databook model.
    It exposes cell values only — no formatting, conditional formatting,
    comments, or images.
    """

    @property
    def info(self) -> LibraryInfo:
        return LibraryInfo(
            name="tablib",
            version=_get_version(),
            language="python",
            capabilities={"read", "write"},
        )

    # =========================================================================
    # Read Operations
    # =========================================================================

    def open_workbook(self, path: Path) -> Any:
        with open(path, "rb") as f:
            raw = f.read()
        book = tablib.Databook()
        book.load(raw, "xlsx")
        # tablib treats the first row as headers by default.
        # Re-insert headers as a data row so we have 0-indexed access.
        for ds in book.sheets():
            if ds.headers:
                row = tuple(ds.headers)
                ds.headers = None
                ds.insert(0, row)
        return book

    def close_workbook(self, workbook: Any) -> None:
        pass

    def get_sheet_names(self, workbook: Any) -> list[str]:
        return [ds.title for ds in workbook.sheets()]

    def read_sheet_values(
        self,
        workbook: Any,
        sheet: str,
        cell_range: str | None = None,
    ) -> list[list[Any]]:
        """Bulk read a rectangular range of values.

        Optional helper used by performance workloads.
        """
        ds = None
        for dataset in workbook.sheets():
            if dataset.title == sheet:
                ds = dataset
                break
        if ds is None:
            return []

        if not cell_range:
            return [list(ds[r]) for r in range(ds.height)]

        r0, c0, r1, c1 = _parse_cell_range(cell_range)
        out: list[list[Any]] = []
        for r in range(r0, min(r1, ds.height - 1) + 1):
            row = ds[r]
            out.append(list(row[c0 : c1 + 1]))
        return out

    def read_cell_value(
        self,
        workbook: Any,
        sheet: str,
        cell: str,
    ) -> CellValue:
        row_idx, col_idx = _parse_cell_ref(cell)

        # Find the dataset by name
        ds = None
        for dataset in workbook.sheets():
            if dataset.title == sheet:
                ds = dataset
                break
        if ds is None:
            return CellValue(type=CellType.BLANK)

        if row_idx >= ds.height:
            return CellValue(type=CellType.BLANK)
        row_data = ds[row_idx]
        if col_idx >= len(row_data):
            return CellValue(type=CellType.BLANK)

        value = row_data[col_idx]

        if value is None or value == "":
            return CellValue(type=CellType.BLANK)

        if isinstance(value, bool):
            return CellValue(type=CellType.BOOLEAN, value=value)

        if isinstance(value, (int, float)):
            return CellValue(type=CellType.NUMBER, value=value)

        if isinstance(value, datetime):
            if (
                value.hour == 0
                and value.minute == 0
                and value.second == 0
                and value.microsecond == 0
            ):
                return CellValue(type=CellType.DATE, value=value.date())
            return CellValue(type=CellType.DATETIME, value=value)

        if isinstance(value, date) and not isinstance(value, datetime):
            return CellValue(type=CellType.DATE, value=value)

        if isinstance(value, time):
            return CellValue(
                type=CellType.DATETIME,
                value=datetime.combine(date.today(), value),
            )

        if isinstance(value, str):
            if value in ("#N/A", "#NULL!", "#NAME?", "#REF!"):
                return CellValue(type=CellType.ERROR, value=value)
            if value.startswith("#") and value.endswith("!"):
                return CellValue(type=CellType.ERROR, value=value)
            if value.startswith("="):
                return CellValue(type=CellType.FORMULA, value=value, formula=value)
            return CellValue(type=CellType.STRING, value=value)

        return CellValue(type=CellType.STRING, value=str(value))

    def read_cell_format(self, workbook: Any, sheet: str, cell: str) -> CellFormat:
        return CellFormat()

    def read_cell_border(self, workbook: Any, sheet: str, cell: str) -> BorderInfo:
        return BorderInfo()

    def read_row_height(self, workbook: Any, sheet: str, row: int) -> float | None:
        return None

    def read_column_width(self, workbook: Any, sheet: str, column: str) -> float | None:
        return None

    # =========================================================================
    # Tier 2 Read Operations
    # =========================================================================

    def read_merged_ranges(self, workbook: Any, sheet: str) -> list[str]:
        return []

    def read_conditional_formats(self, workbook: Any, sheet: str) -> list[JSONDict]:
        return []

    def read_data_validations(self, workbook: Any, sheet: str) -> list[JSONDict]:
        return []

    def read_hyperlinks(self, workbook: Any, sheet: str) -> list[JSONDict]:
        return []

    def read_images(self, workbook: Any, sheet: str) -> list[JSONDict]:
        return []

    def read_pivot_tables(self, workbook: Any, sheet: str) -> list[JSONDict]:
        return []

    def read_comments(self, workbook: Any, sheet: str) -> list[JSONDict]:
        return []

    def read_freeze_panes(self, workbook: Any, sheet: str) -> JSONDict:
        return {}

    # =========================================================================
    # Write Operations
    # =========================================================================

    def create_workbook(self) -> WorkbookData:
        return {"sheets": {}, "_order": [], "_bulk": {}}

    def add_sheet(self, workbook: WorkbookData, name: str) -> None:
        if name not in workbook["sheets"]:
            workbook["sheets"][name] = {}
            workbook["_order"].append(name)

    def write_sheet_values(
        self,
        workbook: WorkbookData,
        sheet: str,
        start_cell: str,
        values: list[list[Any]],
    ) -> None:
        """Bulk write a grid of raw Python values.

        Optional helper used by performance workloads.
        """
        if sheet not in workbook["sheets"]:
            workbook["sheets"][sheet] = {}
            workbook["_order"].append(sheet)

        r0, c0 = _parse_cell_ref(start_cell)
        if r0 == 0 and c0 == 0:
            workbook.setdefault("_bulk", {})[sheet] = values
            return

        for r, row_vals in enumerate(values):
            for c, v in enumerate(row_vals):
                workbook["sheets"][sheet][(r0 + r, c0 + c)] = v

    def write_cell_value(
        self,
        workbook: WorkbookData,
        sheet: str,
        cell: str,
        value: CellValue,
    ) -> None:
        if sheet not in workbook["sheets"]:
            workbook["sheets"][sheet] = {}
            workbook["_order"].append(sheet)

        row_idx, col_idx = _parse_cell_ref(cell)

        raw_value: Any = ""
        if value.type == CellType.BLANK:
            raw_value = ""
        elif value.type == CellType.FORMULA:
            raw_value = value.formula or value.value or ""
        elif value.type == CellType.BOOLEAN:
            raw_value = bool(value.value)
        elif value.type == CellType.NUMBER:
            raw_value = value.value
        elif value.type in (CellType.DATE, CellType.DATETIME):
            raw_value = value.value
        elif value.type == CellType.ERROR:
            raw_value = str(value.value)
        else:
            raw_value = str(value.value) if value.value is not None else ""

        workbook["sheets"][sheet][(row_idx, col_idx)] = raw_value

    def write_cell_format(self, workbook: Any, sheet: str, cell: str, format: CellFormat) -> None:
        pass

    def write_cell_border(self, workbook: Any, sheet: str, cell: str, border: BorderInfo) -> None:
        pass

    def set_row_height(self, workbook: Any, sheet: str, row: int, height: float) -> None:
        pass

    def set_column_width(self, workbook: Any, sheet: str, column: str, width: float) -> None:
        pass

    # =========================================================================
    # Tier 2 Write Operations
    # =========================================================================

    def merge_cells(self, workbook: Any, sheet: str, cell_range: str) -> None:
        pass

    def add_conditional_format(self, workbook: Any, sheet: str, rule: JSONDict) -> None:
        pass

    def add_data_validation(self, workbook: Any, sheet: str, validation: JSONDict) -> None:
        pass

    def add_hyperlink(self, workbook: Any, sheet: str, link: JSONDict) -> None:
        pass

    def add_image(self, workbook: Any, sheet: str, image: JSONDict) -> None:
        pass

    def add_pivot_table(self, workbook: Any, sheet: str, pivot: JSONDict) -> None:
        pass

    def add_comment(self, workbook: Any, sheet: str, comment: JSONDict) -> None:
        pass

    def set_freeze_panes(self, workbook: Any, sheet: str, settings: JSONDict) -> None:
        pass

    def save_workbook(self, workbook: WorkbookData, path: Path) -> None:
        book = tablib.Databook()
        for name in workbook["_order"]:
            bulk = workbook.get("_bulk", {}).get(name)
            cells = workbook["sheets"].get(name, {})
            ds = tablib.Dataset(title=name)
            if isinstance(bulk, list):
                for row_vals in bulk:
                    if isinstance(row_vals, list):
                        ds.append(row_vals)
                    else:
                        ds.append([row_vals])
            elif not cells:
                # Empty dataset — add a single empty row so tablib creates the sheet
                ds.append([""])
            else:
                max_row = max(r for r, _ in cells.keys()) + 1
                max_col = max(c for _, c in cells.keys()) + 1
                for r in range(max_row):
                    row = [cells.get((r, c), "") for c in range(max_col)]
                    ds.append(row)
            book.add_sheet(ds)

        with open(path, "wb") as f:
            f.write(book.export("xlsx"))
