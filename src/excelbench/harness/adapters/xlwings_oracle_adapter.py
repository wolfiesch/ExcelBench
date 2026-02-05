"""Excel oracle adapter using xlwings (read-only)."""

from datetime import date, datetime
from pathlib import Path
from typing import Any

import xlwings as xw

from excelbench.harness.adapters.base import ReadOnlyAdapter
from excelbench.models import (
    BorderEdge,
    BorderInfo,
    BorderStyle,
    CellFormat,
    CellType,
    CellValue,
    LibraryInfo,
)


class XlLineStyle:
    CONTINUOUS = 1
    DASH = -4115
    DASH_DOT = 4
    DASH_DOT_DOT = 5
    DOT = -4118
    DOUBLE = -4119
    NONE = -4142
    SLANT_DASH_DOT = 13


class XlBorderWeight:
    HAIRLINE = 1
    THIN = 2
    MEDIUM = -4138
    THICK = 4


class XlBordersIndex:
    EDGE_LEFT = 7
    EDGE_TOP = 8
    EDGE_BOTTOM = 9
    EDGE_RIGHT = 10
    DIAGONAL_DOWN = 5
    DIAGONAL_UP = 6


H_ALIGN_MAP = {
    -4131: "left",     # xlLeft
    -4108: "center",   # xlCenter
    -4152: "right",    # xlRight
    -4130: "justify",  # xlJustify
    -4117: "distributed",  # xlDistributed
    7: "centerContinuous",
    1: "general",
}

V_ALIGN_MAP = {
    -4160: "top",      # xlTop
    -4108: "center",   # xlCenter
    -4107: "bottom",   # xlBottom
    -4130: "justify",  # xlJustify
    -4117: "distributed",  # xlDistributed
}


def _int_to_hex(color_int: int | None) -> str | None:
    if color_int is None:
        return None
    if isinstance(color_int, float):
        color_int = int(color_int)
    if color_int < 0:
        return None
    r = color_int & 0xFF
    g = (color_int >> 8) & 0xFF
    b = (color_int >> 16) & 0xFF
    return f"#{r:02X}{g:02X}{b:02X}"


class ExcelOracleAdapter(ReadOnlyAdapter):
    """Read-only adapter backed by Excel via xlwings."""

    @property
    def info(self) -> LibraryInfo:
        return LibraryInfo(
            name="excel_oracle",
            version="excel",
            language="excel",
            capabilities={"read"},
        )

    def open_workbook(self, path: Path) -> Any:
        return xw.Book(str(path))

    def close_workbook(self, workbook: Any) -> None:
        workbook.close()

    def get_sheet_names(self, workbook: Any) -> list[str]:
        return [s.name for s in workbook.sheets]

    def read_cell_value(
        self,
        workbook: Any,
        sheet: str,
        cell: str,
    ) -> CellValue:
        ws = workbook.sheets[sheet]
        rng = ws.range(cell)
        value = rng.value
        formula = rng.formula

        if value is None:
            return CellValue(type=CellType.BLANK)

        if isinstance(value, bool):
            return CellValue(type=CellType.BOOLEAN, value=value)

        if isinstance(value, (int, float)):
            return CellValue(type=CellType.NUMBER, value=value)

        if isinstance(value, date) and not isinstance(value, datetime):
            return CellValue(type=CellType.DATE, value=value)

        if isinstance(value, datetime):
            if (
                value.hour == 0
                and value.minute == 0
                and value.second == 0
                and value.microsecond == 0
            ):
                return CellValue(type=CellType.DATE, value=value.date())
            return CellValue(type=CellType.DATETIME, value=value)

        if isinstance(value, str):
            if value.startswith("#"):
                return CellValue(type=CellType.ERROR, value=value)
            if formula:
                return CellValue(type=CellType.FORMULA, value=value, formula=str(formula))
            return CellValue(type=CellType.STRING, value=value)

        if formula:
            return CellValue(type=CellType.FORMULA, value=value, formula=str(formula))

        return CellValue(type=CellType.STRING, value=str(value))

    def read_cell_format(
        self,
        workbook: Any,
        sheet: str,
        cell: str,
    ) -> CellFormat:
        ws = workbook.sheets[sheet]
        rng = ws.range(cell)

        font = rng.api.Font
        font_color = _int_to_hex(font.Color) if getattr(font, "Color", None) is not None else None

        bg_color = None
        if hasattr(rng.api, "Interior") and getattr(rng.api.Interior, "Color", None) is not None:
            bg_color = _int_to_hex(rng.api.Interior.Color)

        underline = None
        if font.Underline:
            if font.Underline == 2:
                underline = "single"
            elif font.Underline == -4119:
                underline = "double"
            else:
                underline = str(font.Underline)

        h_align = H_ALIGN_MAP.get(getattr(rng.api, "HorizontalAlignment", None))
        v_align = V_ALIGN_MAP.get(getattr(rng.api, "VerticalAlignment", None))

        wrap = getattr(rng.api, "WrapText", None)
        rotation = getattr(rng.api, "Orientation", None)
        indent = getattr(rng.api, "IndentLevel", None)

        return CellFormat(
            bold=bool(font.Bold) if font.Bold is not None else None,
            italic=bool(font.Italic) if font.Italic is not None else None,
            underline=underline,
            strikethrough=bool(font.Strikethrough) if font.Strikethrough is not None else None,
            font_name=font.Name if font.Name else None,
            font_size=font.Size if font.Size else None,
            font_color=font_color,
            bg_color=bg_color,
            number_format=rng.number_format if hasattr(rng, "number_format") else None,
            h_align=h_align,
            v_align=v_align,
            wrap=bool(wrap) if wrap is not None else None,
            rotation=int(rotation) if rotation not in (None, 0) else None,
            indent=int(indent) if indent not in (None, 0) else None,
        )

    def read_cell_border(
        self,
        workbook: Any,
        sheet: str,
        cell: str,
    ) -> BorderInfo:
        ws = workbook.sheets[sheet]
        rng = ws.range(cell)

        def parse_border(index: int) -> BorderEdge | None:
            border = rng.api.Borders(index)
            line_style = border.LineStyle
            if line_style in (None, 0, XlLineStyle.NONE):
                return None

            if line_style == XlLineStyle.DOUBLE:
                style = BorderStyle.DOUBLE
            elif line_style == XlLineStyle.DASH:
                style = BorderStyle.DASHED
            elif line_style == XlLineStyle.DOT:
                style = BorderStyle.DOTTED
            elif line_style == XlLineStyle.DASH_DOT:
                style = BorderStyle.DASH_DOT
            elif line_style == XlLineStyle.DASH_DOT_DOT:
                style = BorderStyle.DASH_DOT_DOT
            elif line_style == XlLineStyle.SLANT_DASH_DOT:
                style = BorderStyle.SLANT_DASH_DOT
            else:
                weight = border.Weight
                weight_map = {
                    XlBorderWeight.HAIRLINE: BorderStyle.HAIR,
                    XlBorderWeight.THIN: BorderStyle.THIN,
                    XlBorderWeight.MEDIUM: BorderStyle.MEDIUM,
                    XlBorderWeight.THICK: BorderStyle.THICK,
                }
                style = weight_map.get(weight, BorderStyle.THIN)

            color = _int_to_hex(border.Color) or "#000000"
            return BorderEdge(style=style, color=color)

        return BorderInfo(
            top=parse_border(XlBordersIndex.EDGE_TOP),
            bottom=parse_border(XlBordersIndex.EDGE_BOTTOM),
            left=parse_border(XlBordersIndex.EDGE_LEFT),
            right=parse_border(XlBordersIndex.EDGE_RIGHT),
            diagonal_up=parse_border(XlBordersIndex.DIAGONAL_UP),
            diagonal_down=parse_border(XlBordersIndex.DIAGONAL_DOWN),
        )

    def read_row_height(
        self,
        workbook: Any,
        sheet: str,
        row: int,
    ) -> float | None:
        ws = workbook.sheets[sheet]
        return ws.range(f"{row}:{row}").row_height

    def read_column_width(
        self,
        workbook: Any,
        sheet: str,
        column: str,
    ) -> float | None:
        ws = workbook.sheets[sheet]
        return ws.range(f"{column}:{column}").column_width
