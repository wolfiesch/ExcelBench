"""Adapter for openpyxl library."""

from datetime import datetime, date
from pathlib import Path
from typing import Any

import openpyxl
from openpyxl import Workbook
from openpyxl.cell import Cell
from openpyxl.styles import Font, Border, Side, PatternFill
from openpyxl.utils import get_column_letter

from excelbench.harness.adapters.base import ExcelAdapter
from excelbench.models import (
    CellValue,
    CellType,
    CellFormat,
    BorderInfo,
    BorderEdge,
    BorderStyle,
    LibraryInfo,
)


def _get_version() -> str:
    """Get openpyxl version."""
    return openpyxl.__version__


class OpenpyxlAdapter(ExcelAdapter):
    """Adapter for openpyxl library (read/write support)."""

    @property
    def info(self) -> LibraryInfo:
        return LibraryInfo(
            name="openpyxl",
            version=_get_version(),
            language="python",
            capabilities={"read", "write"},
        )

    # =========================================================================
    # Read Operations
    # =========================================================================

    def open_workbook(self, path: Path) -> Workbook:
        """Open a workbook for reading."""
        return openpyxl.load_workbook(str(path), data_only=False)

    def close_workbook(self, workbook: Any) -> None:
        """Close an opened workbook."""
        workbook.close()

    def get_sheet_names(self, workbook: Workbook) -> list[str]:
        """Get list of sheet names in a workbook."""
        return workbook.sheetnames

    def read_cell_value(
        self,
        workbook: Workbook,
        sheet: str,
        cell: str,
    ) -> CellValue:
        """Read the value of a cell."""
        ws = workbook[sheet]
        c: Cell = ws[cell]

        # Handle different value types
        value = c.value

        if value is None:
            return CellValue(type=CellType.BLANK)

        if isinstance(value, bool):
            return CellValue(type=CellType.BOOLEAN, value=value)

        if isinstance(value, (int, float)):
            return CellValue(type=CellType.NUMBER, value=value)

        if isinstance(value, datetime):
            return CellValue(type=CellType.DATETIME, value=value)

        if isinstance(value, date):
            return CellValue(type=CellType.DATE, value=value)

        if isinstance(value, str):
            # Check if it's an error value
            if value.startswith("#") and value.endswith("!"):
                return CellValue(type=CellType.ERROR, value=value)

            # Check if there's a formula
            if c.data_type == "f" or (hasattr(c, "value") and str(c.value).startswith("=")):
                return CellValue(
                    type=CellType.FORMULA,
                    value=value,
                    formula=str(c.value) if str(c.value).startswith("=") else None,
                )

            return CellValue(type=CellType.STRING, value=value)

        # Fallback to string
        return CellValue(type=CellType.STRING, value=str(value))

    def read_cell_format(
        self,
        workbook: Workbook,
        sheet: str,
        cell: str,
    ) -> CellFormat:
        """Read the formatting of a cell."""
        ws = workbook[sheet]
        c: Cell = ws[cell]
        font: Font = c.font

        # Convert color to hex
        font_color = None
        if font.color and font.color.rgb:
            rgb = font.color.rgb
            if isinstance(rgb, str) and len(rgb) >= 6:
                # Handle ARGB format (8 chars) or RGB format (6 chars)
                if len(rgb) == 8:
                    font_color = f"#{rgb[2:]}"  # Skip alpha
                else:
                    font_color = f"#{rgb}"

        # Get background color
        bg_color = None
        fill = c.fill
        if fill and fill.patternType == "solid" and fill.fgColor and fill.fgColor.rgb:
            rgb = fill.fgColor.rgb
            if isinstance(rgb, str) and len(rgb) >= 6:
                if len(rgb) == 8:
                    bg_color = f"#{rgb[2:]}"
                else:
                    bg_color = f"#{rgb}"

        # Map underline
        underline = None
        if font.underline:
            underline_map = {
                "single": "single",
                "double": "double",
                "singleAccounting": "singleAccounting",
                "doubleAccounting": "doubleAccounting",
            }
            underline = underline_map.get(font.underline, font.underline)

        return CellFormat(
            bold=font.bold if font.bold else None,
            italic=font.italic if font.italic else None,
            underline=underline,
            strikethrough=font.strike if font.strike else None,
            font_name=font.name if font.name else None,
            font_size=font.size if font.size else None,
            font_color=font_color,
            bg_color=bg_color,
        )

    def read_cell_border(
        self,
        workbook: Workbook,
        sheet: str,
        cell: str,
    ) -> BorderInfo:
        """Read the border information of a cell."""
        ws = workbook[sheet]
        c: Cell = ws[cell]
        border: Border = c.border

        def parse_side(side: Side | None) -> BorderEdge | None:
            if side is None or side.style is None:
                return None

            # Map openpyxl style to our style
            style_map = {
                "thin": BorderStyle.THIN,
                "medium": BorderStyle.MEDIUM,
                "thick": BorderStyle.THICK,
                "double": BorderStyle.DOUBLE,
                "dashed": BorderStyle.DASHED,
                "dotted": BorderStyle.DOTTED,
                "hair": BorderStyle.HAIR,
                "mediumDashed": BorderStyle.MEDIUM_DASHED,
                "dashDot": BorderStyle.DASH_DOT,
                "mediumDashDot": BorderStyle.MEDIUM_DASH_DOT,
                "dashDotDot": BorderStyle.DASH_DOT_DOT,
                "mediumDashDotDot": BorderStyle.MEDIUM_DASH_DOT_DOT,
                "slantDashDot": BorderStyle.SLANT_DASH_DOT,
            }

            style = style_map.get(side.style, BorderStyle.THIN)

            # Get color
            color = "#000000"
            if side.color and side.color.rgb:
                rgb = side.color.rgb
                if isinstance(rgb, str) and len(rgb) >= 6:
                    if len(rgb) == 8:
                        color = f"#{rgb[2:]}"
                    else:
                        color = f"#{rgb}"

            return BorderEdge(style=style, color=color)

        return BorderInfo(
            top=parse_side(border.top),
            bottom=parse_side(border.bottom),
            left=parse_side(border.left),
            right=parse_side(border.right),
            diagonal_up=parse_side(border.diagonal) if border.diagonalUp else None,
            diagonal_down=parse_side(border.diagonal) if border.diagonalDown else None,
        )

    # =========================================================================
    # Write Operations
    # =========================================================================

    def create_workbook(self) -> Workbook:
        """Create a new workbook."""
        return Workbook()

    def add_sheet(self, workbook: Workbook, name: str) -> None:
        """Add a new sheet to a workbook."""
        workbook.create_sheet(name)

    def write_cell_value(
        self,
        workbook: Workbook,
        sheet: str,
        cell: str,
        value: CellValue,
    ) -> None:
        """Write a value to a cell."""
        ws = workbook[sheet]
        c: Cell = ws[cell]

        if value.type == CellType.BLANK:
            c.value = None
        elif value.type == CellType.FORMULA:
            c.value = value.formula or value.value
        elif value.type == CellType.ERROR:
            # Write a formula that produces the error
            error_formulas = {
                "#DIV/0!": "=1/0",
                "#N/A": "=NA()",
                "#VALUE!": '="text"+1',
                "#REF!": "=#REF!",
                "#NAME?": "=_undefined_name_",
                "#NUM!": "=SQRT(-1)",
                "#NULL!": "=A1:A2 B1:B2",
            }
            c.value = error_formulas.get(value.value, value.value)
        else:
            c.value = value.value

    def write_cell_format(
        self,
        workbook: Workbook,
        sheet: str,
        cell: str,
        format: CellFormat,
    ) -> None:
        """Apply formatting to a cell."""
        ws = workbook[sheet]
        c: Cell = ws[cell]

        # Build font kwargs
        font_kwargs = {}

        if format.bold is not None:
            font_kwargs["bold"] = format.bold
        if format.italic is not None:
            font_kwargs["italic"] = format.italic
        if format.underline is not None:
            font_kwargs["underline"] = format.underline
        if format.strikethrough is not None:
            font_kwargs["strike"] = format.strikethrough
        if format.font_name is not None:
            font_kwargs["name"] = format.font_name
        if format.font_size is not None:
            font_kwargs["size"] = format.font_size
        if format.font_color is not None:
            from openpyxl.styles import Color
            # Remove # prefix if present
            hex_color = format.font_color.lstrip("#")
            font_kwargs["color"] = Color(rgb=f"FF{hex_color}")

        if font_kwargs:
            c.font = Font(**font_kwargs)

        # Apply background color
        if format.bg_color is not None:
            hex_color = format.bg_color.lstrip("#")
            c.fill = PatternFill(
                start_color=f"FF{hex_color}",
                end_color=f"FF{hex_color}",
                fill_type="solid",
            )

    def write_cell_border(
        self,
        workbook: Workbook,
        sheet: str,
        cell: str,
        border: BorderInfo,
    ) -> None:
        """Apply border to a cell."""
        ws = workbook[sheet]
        c: Cell = ws[cell]

        def make_side(edge: BorderEdge | None) -> Side:
            if edge is None:
                return Side()

            # Map our style to openpyxl style
            style_map = {
                BorderStyle.NONE: None,
                BorderStyle.THIN: "thin",
                BorderStyle.MEDIUM: "medium",
                BorderStyle.THICK: "thick",
                BorderStyle.DOUBLE: "double",
                BorderStyle.DASHED: "dashed",
                BorderStyle.DOTTED: "dotted",
                BorderStyle.HAIR: "hair",
                BorderStyle.MEDIUM_DASHED: "mediumDashed",
                BorderStyle.DASH_DOT: "dashDot",
                BorderStyle.MEDIUM_DASH_DOT: "mediumDashDot",
                BorderStyle.DASH_DOT_DOT: "dashDotDot",
                BorderStyle.MEDIUM_DASH_DOT_DOT: "mediumDashDotDot",
                BorderStyle.SLANT_DASH_DOT: "slantDashDot",
            }

            style = style_map.get(edge.style)
            if style is None:
                return Side()

            hex_color = edge.color.lstrip("#")
            from openpyxl.styles import Color
            return Side(style=style, color=Color(rgb=f"FF{hex_color}"))

        # Determine diagonal settings
        diagonal_side = Side()
        diagonal_up = False
        diagonal_down = False

        if border.diagonal_up:
            diagonal_side = make_side(border.diagonal_up)
            diagonal_up = True
        if border.diagonal_down:
            diagonal_side = make_side(border.diagonal_down)
            diagonal_down = True

        c.border = Border(
            left=make_side(border.left),
            right=make_side(border.right),
            top=make_side(border.top),
            bottom=make_side(border.bottom),
            diagonal=diagonal_side,
            diagonalUp=diagonal_up,
            diagonalDown=diagonal_down,
        )

    def save_workbook(self, workbook: Workbook, path: Path) -> None:
        """Save a workbook to a file."""
        workbook.save(str(path))
