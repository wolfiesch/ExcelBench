"""Adapter for xlsxwriter library (write-only)."""

from datetime import date as _date
from datetime import datetime as _datetime
from pathlib import Path
from typing import Any

import xlsxwriter
from xlsxwriter import Workbook

from excelbench.harness.adapters.base import WriteOnlyAdapter
from excelbench.models import (
    BorderInfo,
    BorderStyle,
    CellFormat,
    CellType,
    CellValue,
    LibraryInfo,
)


def _get_version() -> str:
    """Get xlsxwriter version."""
    return xlsxwriter.__version__


class XlsxwriterAdapter(WriteOnlyAdapter):
    """Adapter for xlsxwriter library (write-only).

    Note: xlsxwriter uses a "format once, apply many" pattern where
    formats are created on the workbook and then applied to cells.
    This adapter creates formats on-demand for simplicity.
    """

    def __init__(self):
        self._workbooks: dict[int, dict] = {}  # wb id -> {sheets, formats, path}

    @property
    def info(self) -> LibraryInfo:
        return LibraryInfo(
            name="xlsxwriter",
            version=_get_version(),
            language="python",
            capabilities={"write"},
        )

    def create_workbook(self) -> dict:
        """Create a new workbook.

        Returns a wrapper dict because xlsxwriter workbooks need
        to be saved to a path at creation time.
        """
        # Return a placeholder - actual workbook created at save time
        wb_data = {
            "sheets": {},  # sheet_name -> list of (cell, value, format)
            "row_heights": {},  # sheet_name -> {row_index: height}
            "col_widths": {},  # sheet_name -> {col_index: width}
            "path": None,
            "workbook": None,
        }
        return wb_data

    def add_sheet(self, workbook: dict, name: str) -> None:
        """Add a new sheet to a workbook."""
        if name not in workbook["sheets"]:
            workbook["sheets"][name] = []

    def _ensure_sheet(self, workbook: dict, sheet: str) -> None:
        """Ensure a sheet exists."""
        if sheet not in workbook["sheets"]:
            workbook["sheets"][sheet] = []
        if sheet not in workbook["row_heights"]:
            workbook["row_heights"][sheet] = {}
        if sheet not in workbook["col_widths"]:
            workbook["col_widths"][sheet] = {}

    def _parse_cell(self, cell: str) -> tuple[int, int]:
        """Parse cell reference like 'A1' to (row, col) tuple."""
        import re
        match = re.match(r"([A-Z]+)(\d+)", cell.upper())
        if not match:
            raise ValueError(f"Invalid cell reference: {cell}")

        col_str, row_str = match.groups()
        row = int(row_str) - 1  # Convert to 0-indexed

        # Convert column letters to number
        col = 0
        for char in col_str:
            col = col * 26 + (ord(char) - ord('A') + 1)
        col -= 1  # Convert to 0-indexed

        return row, col

    def _col_to_index(self, column: str) -> int:
        """Convert column letter(s) to 0-indexed column number."""
        col = 0
        for char in column.upper():
            col = col * 26 + (ord(char) - ord('A') + 1)
        return col - 1

    def write_cell_value(
        self,
        workbook: dict,
        sheet: str,
        cell: str,
        value: CellValue,
    ) -> None:
        """Write a value to a cell."""
        self._ensure_sheet(workbook, sheet)
        row, col = self._parse_cell(cell)

        # Store the operation for later execution
        workbook["sheets"][sheet].append({
            "type": "value",
            "row": row,
            "col": col,
            "value": value,
        })

    def write_cell_format(
        self,
        workbook: dict,
        sheet: str,
        cell: str,
        format: CellFormat,
    ) -> None:
        """Apply formatting to a cell."""
        self._ensure_sheet(workbook, sheet)
        row, col = self._parse_cell(cell)

        workbook["sheets"][sheet].append({
            "type": "format",
            "row": row,
            "col": col,
            "format": format,
        })

    def write_cell_border(
        self,
        workbook: dict,
        sheet: str,
        cell: str,
        border: BorderInfo,
    ) -> None:
        """Apply border to a cell."""
        self._ensure_sheet(workbook, sheet)
        row, col = self._parse_cell(cell)

        workbook["sheets"][sheet].append({
            "type": "border",
            "row": row,
            "col": col,
            "border": border,
        })

    def _create_format(
        self,
        wb: Workbook,
        cell_format: CellFormat | None = None,
        border: BorderInfo | None = None,
    ) -> Any:
        """Create an xlsxwriter format from our models."""
        fmt_dict = {}

        if cell_format:
            if cell_format.bold:
                fmt_dict["bold"] = True
            if cell_format.italic:
                fmt_dict["italic"] = True
            if cell_format.underline:
                underline_map = {
                    "single": 1,
                    "double": 2,
                    "singleAccounting": 33,
                    "doubleAccounting": 34,
                }
                fmt_dict["underline"] = underline_map.get(cell_format.underline, 1)
            if cell_format.strikethrough:
                fmt_dict["font_strikeout"] = True
            if cell_format.font_name:
                fmt_dict["font_name"] = cell_format.font_name
            if cell_format.font_size:
                fmt_dict["font_size"] = cell_format.font_size
            if cell_format.font_color:
                fmt_dict["font_color"] = cell_format.font_color
            if cell_format.bg_color:
                fmt_dict["bg_color"] = cell_format.bg_color
            if cell_format.number_format:
                fmt_dict["num_format"] = cell_format.number_format
            if cell_format.h_align:
                h_align_map = {
                    "center": "center",
                    "left": "left",
                    "right": "right",
                    "justify": "justify",
                    "centerContinuous": "center_across",
                    "distributed": "distributed",
                    "general": "general",
                }
                fmt_dict["align"] = h_align_map.get(cell_format.h_align, cell_format.h_align)
            if cell_format.v_align:
                v_align_map = {
                    "top": "top",
                    "center": "vcenter",
                    "bottom": "bottom",
                    "justify": "vjustify",
                    "distributed": "vdistributed",
                }
                fmt_dict["valign"] = v_align_map.get(cell_format.v_align, cell_format.v_align)
            if cell_format.wrap:
                fmt_dict["text_wrap"] = True
            if cell_format.rotation is not None:
                fmt_dict["rotation"] = cell_format.rotation
            if cell_format.indent is not None:
                fmt_dict["indent"] = cell_format.indent

        if border:
            border_style_map = {
                BorderStyle.NONE: 0,
                BorderStyle.THIN: 1,
                BorderStyle.MEDIUM: 2,
                BorderStyle.DASHED: 3,
                BorderStyle.DOTTED: 4,
                BorderStyle.THICK: 5,
                BorderStyle.DOUBLE: 6,
                BorderStyle.HAIR: 7,
                BorderStyle.MEDIUM_DASHED: 8,
                BorderStyle.DASH_DOT: 9,
                BorderStyle.MEDIUM_DASH_DOT: 10,
                BorderStyle.DASH_DOT_DOT: 11,
                BorderStyle.MEDIUM_DASH_DOT_DOT: 12,
                BorderStyle.SLANT_DASH_DOT: 13,
            }

            if border.top:
                fmt_dict["top"] = border_style_map.get(border.top.style, 1)
                fmt_dict["top_color"] = border.top.color
            if border.bottom:
                fmt_dict["bottom"] = border_style_map.get(border.bottom.style, 1)
                fmt_dict["bottom_color"] = border.bottom.color
            if border.left:
                fmt_dict["left"] = border_style_map.get(border.left.style, 1)
                fmt_dict["left_color"] = border.left.color
            if border.right:
                fmt_dict["right"] = border_style_map.get(border.right.style, 1)
                fmt_dict["right_color"] = border.right.color

            # Diagonal borders
            if border.diagonal_up or border.diagonal_down:
                diag_border = border.diagonal_up or border.diagonal_down
                fmt_dict["diag_border"] = border_style_map.get(diag_border.style, 1)
                fmt_dict["diag_color"] = diag_border.color

                diag_type = 0
                if border.diagonal_up and border.diagonal_down:
                    diag_type = 3
                elif border.diagonal_up:
                    diag_type = 2
                elif border.diagonal_down:
                    diag_type = 1
                fmt_dict["diag_type"] = diag_type

        return wb.add_format(fmt_dict)

    def save_workbook(self, workbook: dict, path: Path) -> None:
        """Save a workbook to a file.

        This is where the actual xlsxwriter workbook is created and
        all queued operations are executed.
        """
        wb = xlsxwriter.Workbook(str(path))

        try:
            for sheet_name, operations in workbook["sheets"].items():
                ws = wb.add_worksheet(sheet_name)

                # Apply row heights / column widths
                for row_index, height in workbook["row_heights"].get(sheet_name, {}).items():
                    ws.set_row(row_index, height)
                for col_index, width in workbook["col_widths"].get(sheet_name, {}).items():
                    ws.set_column(col_index, col_index, width)

                # Group operations by cell to merge formats
                cell_ops: dict[tuple[int, int], dict] = {}

                for op in operations:
                    key = (op["row"], op["col"])
                    if key not in cell_ops:
                        cell_ops[key] = {"value": None, "format": None, "border": None}

                    if op["type"] == "value":
                        cell_ops[key]["value"] = op["value"]
                    elif op["type"] == "format":
                        cell_ops[key]["format"] = op["format"]
                    elif op["type"] == "border":
                        cell_ops[key]["border"] = op["border"]

                # Write all cells
                for (row, col), data in cell_ops.items():
                    cell_value = data["value"]
                    cell_format = data["format"]
                    cell_border = data["border"]

                    # Create format combining format and border
                    fmt = None
                    if cell_format or cell_border:
                        fmt = self._create_format(wb, cell_format, cell_border)

                    # Write value
                    if cell_value:
                        if cell_value.type in (CellType.DATE, CellType.DATETIME) and fmt is None:
                            default_format = (
                                "yyyy-mm-dd"
                                if cell_value.type == CellType.DATE
                                else "yyyy-mm-dd hh:mm:ss"
                            )
                            fmt = self._create_format(
                                wb,
                                CellFormat(number_format=default_format),
                                None,
                            )
                        if cell_value.type == CellType.BLANK:
                            ws.write_blank(row, col, None, fmt)
                        elif cell_value.type == CellType.FORMULA:
                            ws.write_formula(row, col, cell_value.formula or cell_value.value, fmt)
                        elif cell_value.type == CellType.BOOLEAN:
                            ws.write_boolean(row, col, cell_value.value, fmt)
                        elif cell_value.type == CellType.NUMBER:
                            ws.write_number(row, col, cell_value.value, fmt)
                        elif cell_value.type == CellType.DATE:
                            dt_value = cell_value.value
                            if isinstance(dt_value, _date) and not isinstance(dt_value, _datetime):
                                dt_value = _datetime.combine(dt_value, _datetime.min.time())
                            ws.write_datetime(row, col, dt_value, fmt)
                        elif cell_value.type == CellType.DATETIME:
                            ws.write_datetime(row, col, cell_value.value, fmt)
                        elif cell_value.type == CellType.ERROR:
                            # Write formula that produces error
                            error_formulas = {
                                "#DIV/0!": "=1/0",
                                "#N/A": "=NA()",
                                "#VALUE!": '="text"+1',
                            }
                            fallback = f'=ERROR("{cell_value.value}")'
                            formula = error_formulas.get(cell_value.value, fallback)
                            ws.write_formula(row, col, formula, fmt)
                        else:
                            ws.write_string(row, col, str(cell_value.value), fmt)
                    elif fmt:
                        # Write blank with format
                        ws.write_blank(row, col, None, fmt)

        finally:
            wb.close()

    def set_row_height(
        self,
        workbook: dict,
        sheet: str,
        row: int,
        height: float,
    ) -> None:
        self._ensure_sheet(workbook, sheet)
        workbook["row_heights"][sheet][row - 1] = height

    def set_column_width(
        self,
        workbook: dict,
        sheet: str,
        column: str,
        width: float,
    ) -> None:
        self._ensure_sheet(workbook, sheet)
        col_index = self._col_to_index(column)
        workbook["col_widths"][sheet][col_index] = width
