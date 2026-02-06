"""Adapter for pylightxl library (read/write, zero-dependency)."""

import re
from datetime import date, datetime
from pathlib import Path
from typing import Any

import pylightxl

from excelbench.harness.adapters.base import ExcelAdapter
from excelbench.models import (
    BorderInfo,
    CellFormat,
    CellType,
    CellValue,
    LibraryInfo,
)


def _get_version() -> str:
    """Get pylightxl version."""
    try:
        from importlib.metadata import version
        return version("pylightxl")
    except Exception:
        return "unknown"


def _parse_cell_ref(cell: str) -> tuple[int, int]:
    """Parse a cell reference like 'A1' to (row_1based, col_1based)."""
    match = re.match(r"([A-Z]+)(\d+)", cell.upper())
    if not match:
        raise ValueError(f"Invalid cell reference: {cell}")
    col_str, row_str = match.groups()
    row = int(row_str)
    col = 0
    for char in col_str:
        col = col * 26 + (ord(char) - ord("A") + 1)
    return row, col


class PylightxlAdapter(ExcelAdapter):
    """Adapter for pylightxl library (read/write).

    pylightxl is a zero-dependency, lightweight library. It does NOT support
    formatting, borders, or most Tier 2 features. Formulas are preserved as
    strings. Dates are NOT auto-converted (serial numbers returned as floats).
    """

    @property
    def info(self) -> LibraryInfo:
        return LibraryInfo(
            name="pylightxl",
            version=_get_version(),
            language="python",
            capabilities={"read", "write"},
        )

    # =========================================================================
    # Read Operations
    # =========================================================================

    def open_workbook(self, path: Path) -> Any:
        return pylightxl.readxl(fn=str(path))

    def close_workbook(self, workbook: Any) -> None:
        pass  # No close needed

    def get_sheet_names(self, workbook: Any) -> list[str]:
        return workbook.ws_names

    def read_cell_value(
        self,
        workbook: Any,
        sheet: str,
        cell: str,
    ) -> CellValue:
        value = workbook.ws(ws=sheet).address(address=cell)

        # pylightxl returns "" for empty cells
        if value == "":
            return CellValue(type=CellType.BLANK)

        if value is None:
            return CellValue(type=CellType.BLANK)

        # Check bool BEFORE int (bool is subclass of int)
        if isinstance(value, bool):
            return CellValue(type=CellType.BOOLEAN, value=value)

        if isinstance(value, (int, float)):
            return CellValue(type=CellType.NUMBER, value=value)

        if isinstance(value, datetime):
            is_midnight = (
                value.hour == 0 and value.minute == 0
                and value.second == 0 and value.microsecond == 0
            )
            if is_midnight:
                return CellValue(type=CellType.DATE, value=value.date())
            return CellValue(type=CellType.DATETIME, value=value)

        if isinstance(value, date) and not isinstance(value, datetime):
            return CellValue(type=CellType.DATE, value=value)

        if isinstance(value, str):
            # Error values
            if value.startswith("#") and value.endswith("!"):
                return CellValue(type=CellType.ERROR, value=value)

            # Formulas — pylightxl preserves formula strings
            if value.startswith("="):
                return CellValue(type=CellType.FORMULA, value=value, formula=value)

            return CellValue(type=CellType.STRING, value=value)

        # Fallback
        return CellValue(type=CellType.STRING, value=str(value))

    def read_cell_format(
        self,
        workbook: Any,
        sheet: str,
        cell: str,
    ) -> CellFormat:
        return CellFormat()  # No formatting support

    def read_cell_border(
        self,
        workbook: Any,
        sheet: str,
        cell: str,
    ) -> BorderInfo:
        return BorderInfo()  # No border support

    def read_row_height(
        self,
        workbook: Any,
        sheet: str,
        row: int,
    ) -> float | None:
        return None

    def read_column_width(
        self,
        workbook: Any,
        sheet: str,
        column: str,
    ) -> float | None:
        return None

    # =========================================================================
    # Tier 2 Read Operations
    # =========================================================================

    def read_merged_ranges(self, workbook: Any, sheet: str) -> list[str]:
        return []

    def read_conditional_formats(self, workbook: Any, sheet: str) -> list[dict]:
        return []

    def read_data_validations(self, workbook: Any, sheet: str) -> list[dict]:
        return []

    def read_hyperlinks(self, workbook: Any, sheet: str) -> list[dict]:
        return []

    def read_images(self, workbook: Any, sheet: str) -> list[dict]:
        return []

    def read_pivot_tables(self, workbook: Any, sheet: str) -> list[dict]:
        return []

    def read_comments(self, workbook: Any, sheet: str) -> list[dict]:
        return []

    def read_freeze_panes(self, workbook: Any, sheet: str) -> dict:
        return {}

    # =========================================================================
    # Write Operations
    # =========================================================================

    def create_workbook(self) -> Any:
        return pylightxl.Database()

    def add_sheet(self, workbook: Any, name: str) -> None:
        workbook.add_ws(ws=name)

    def write_cell_value(
        self,
        workbook: Any,
        sheet: str,
        cell: str,
        value: CellValue,
    ) -> None:
        if value.type == CellType.BLANK:
            workbook.ws(ws=sheet).update_address(address=cell, val="")
        elif value.type == CellType.FORMULA:
            workbook.ws(ws=sheet).update_address(address=cell, val=value.formula or value.value)
        elif value.type == CellType.ERROR:
            workbook.ws(ws=sheet).update_address(address=cell, val=str(value.value))
        elif value.type == CellType.BOOLEAN:
            workbook.ws(ws=sheet).update_address(address=cell, val=value.value)
        elif value.type == CellType.DATE:
            val = value.value
            if isinstance(val, date):
                val = val.isoformat()
            workbook.ws(ws=sheet).update_address(address=cell, val=val)
        elif value.type == CellType.DATETIME:
            val = value.value
            if isinstance(val, datetime):
                val = val.isoformat()
            workbook.ws(ws=sheet).update_address(address=cell, val=val)
        else:
            workbook.ws(ws=sheet).update_address(address=cell, val=value.value)

    def write_cell_format(
        self,
        workbook: Any,
        sheet: str,
        cell: str,
        format: CellFormat,
    ) -> None:
        pass  # Not supported

    def write_cell_border(
        self,
        workbook: Any,
        sheet: str,
        cell: str,
        border: BorderInfo,
    ) -> None:
        pass  # Not supported

    def set_row_height(
        self,
        workbook: Any,
        sheet: str,
        row: int,
        height: float,
    ) -> None:
        pass  # Not supported

    def set_column_width(
        self,
        workbook: Any,
        sheet: str,
        column: str,
        width: float,
    ) -> None:
        pass  # Not supported

    def save_workbook(self, workbook: Any, path: Path) -> None:
        # pylightxl tries to read an existing file as a ZIP for in-place update.
        # Remove any pre-existing (possibly empty/invalid) file so it creates fresh.
        if path.exists():
            path.unlink()
        pylightxl.writexl(db=workbook, fn=str(path))

    # =========================================================================
    # Tier 2 Write Operations
    # =========================================================================

    def merge_cells(self, workbook: Any, sheet: str, cell_range: str) -> None:
        pass  # Not supported

    def add_conditional_format(self, workbook: Any, sheet: str, rule: dict) -> None:
        pass  # Not supported

    def add_data_validation(self, workbook: Any, sheet: str, validation: dict) -> None:
        pass  # Not supported

    def add_hyperlink(self, workbook: Any, sheet: str, link: dict) -> None:
        pass  # Not supported

    def add_image(self, workbook: Any, sheet: str, image: dict) -> None:
        pass  # Not supported

    def add_pivot_table(self, workbook: Any, sheet: str, pivot: dict) -> None:
        pass  # Not supported

    def add_comment(self, workbook: Any, sheet: str, comment: dict) -> None:
        pass  # Not supported

    def set_freeze_panes(self, workbook: Any, sheet: str, settings: dict) -> None:
        pass  # Not supported
