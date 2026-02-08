"""Tests for XlsxwriterConstmemAdapter (write-only, constant_memory mode)."""

from __future__ import annotations

from datetime import date, datetime
from pathlib import Path

import pytest

from excelbench.harness.adapters.openpyxl_adapter import OpenpyxlAdapter
from excelbench.harness.adapters.xlsxwriter_constmem_adapter import XlsxwriterConstmemAdapter
from excelbench.models import (
    BorderEdge,
    BorderInfo,
    BorderStyle,
    CellFormat,
    CellType,
    CellValue,
)


@pytest.fixture
def opxl() -> OpenpyxlAdapter:
    return OpenpyxlAdapter()


@pytest.fixture
def cm() -> XlsxwriterConstmemAdapter:
    return XlsxwriterConstmemAdapter()


# ═════════════════════════════════════════════════════════════════════════
# TestConstmemInfo
# ═════════════════════════════════════════════════════════════════════════


class TestConstmemInfo:
    def test_name(self, cm: XlsxwriterConstmemAdapter) -> None:
        assert cm.info.name == "xlsxwriter-constmem"

    def test_version(self, cm: XlsxwriterConstmemAdapter) -> None:
        assert cm.info.version != "unknown"

    def test_capabilities(self, cm: XlsxwriterConstmemAdapter) -> None:
        assert cm.info.capabilities == {"write"}

    def test_is_write_only(self, cm: XlsxwriterConstmemAdapter) -> None:
        assert cm.can_write()
        assert not cm.can_read()

    def test_language(self, cm: XlsxwriterConstmemAdapter) -> None:
        assert cm.info.language == "python"


# ═════════════════════════════════════════════════════════════════════════
# TestConstmemWriteRoundtrip
# ═════════════════════════════════════════════════════════════════════════


class TestConstmemWriteRoundtrip:
    """Write via xlsxwriter-constmem, read back via openpyxl."""

    def test_string(
        self, cm: XlsxwriterConstmemAdapter, opxl: OpenpyxlAdapter, tmp_path: Path
    ) -> None:
        path = tmp_path / "cm_str.xlsx"
        wb = cm.create_workbook()
        cm.add_sheet(wb, "S1")
        cm.write_cell_value(wb, "S1", "A1", CellValue(type=CellType.STRING, value="hello"))
        cm.save_workbook(wb, path)

        rb = opxl.open_workbook(path)
        cv = opxl.read_cell_value(rb, "S1", "A1")
        assert cv.type == CellType.STRING
        assert cv.value == "hello"
        opxl.close_workbook(rb)

    def test_number(
        self, cm: XlsxwriterConstmemAdapter, opxl: OpenpyxlAdapter, tmp_path: Path
    ) -> None:
        path = tmp_path / "cm_num.xlsx"
        wb = cm.create_workbook()
        cm.add_sheet(wb, "S1")
        cm.write_cell_value(wb, "S1", "A1", CellValue(type=CellType.NUMBER, value=42.5))
        cm.save_workbook(wb, path)

        rb = opxl.open_workbook(path)
        cv = opxl.read_cell_value(rb, "S1", "A1")
        assert cv.type == CellType.NUMBER
        assert cv.value == 42.5
        opxl.close_workbook(rb)

    def test_boolean(
        self, cm: XlsxwriterConstmemAdapter, opxl: OpenpyxlAdapter, tmp_path: Path
    ) -> None:
        path = tmp_path / "cm_bool.xlsx"
        wb = cm.create_workbook()
        cm.add_sheet(wb, "S1")
        cm.write_cell_value(wb, "S1", "A1", CellValue(type=CellType.BOOLEAN, value=True))
        cm.save_workbook(wb, path)

        rb = opxl.open_workbook(path)
        cv = opxl.read_cell_value(rb, "S1", "A1")
        assert cv.type == CellType.BOOLEAN
        assert cv.value is True
        opxl.close_workbook(rb)

    def test_date(
        self, cm: XlsxwriterConstmemAdapter, opxl: OpenpyxlAdapter, tmp_path: Path
    ) -> None:
        path = tmp_path / "cm_date.xlsx"
        wb = cm.create_workbook()
        cm.add_sheet(wb, "S1")
        cm.write_cell_value(
            wb, "S1", "A1", CellValue(type=CellType.DATE, value=date(2024, 6, 15))
        )
        cm.save_workbook(wb, path)

        rb = opxl.open_workbook(path)
        cv = opxl.read_cell_value(rb, "S1", "A1")
        assert cv.type in (CellType.DATE, CellType.DATETIME)
        opxl.close_workbook(rb)

    def test_datetime(
        self, cm: XlsxwriterConstmemAdapter, opxl: OpenpyxlAdapter, tmp_path: Path
    ) -> None:
        path = tmp_path / "cm_dt.xlsx"
        wb = cm.create_workbook()
        cm.add_sheet(wb, "S1")
        cm.write_cell_value(
            wb,
            "S1",
            "A1",
            CellValue(type=CellType.DATETIME, value=datetime(2024, 6, 15, 14, 30, 0)),
        )
        cm.save_workbook(wb, path)

        rb = opxl.open_workbook(path)
        cv = opxl.read_cell_value(rb, "S1", "A1")
        assert cv.type == CellType.DATETIME
        opxl.close_workbook(rb)

    def test_blank(
        self, cm: XlsxwriterConstmemAdapter, opxl: OpenpyxlAdapter, tmp_path: Path
    ) -> None:
        path = tmp_path / "cm_blank.xlsx"
        wb = cm.create_workbook()
        cm.add_sheet(wb, "S1")
        cm.write_cell_value(wb, "S1", "A1", CellValue(type=CellType.BLANK))
        cm.write_cell_value(wb, "S1", "A2", CellValue(type=CellType.STRING, value="x"))
        cm.save_workbook(wb, path)

        rb = opxl.open_workbook(path)
        cv = opxl.read_cell_value(rb, "S1", "A1")
        assert cv.type == CellType.BLANK
        opxl.close_workbook(rb)

    def test_formula(
        self, cm: XlsxwriterConstmemAdapter, opxl: OpenpyxlAdapter, tmp_path: Path
    ) -> None:
        path = tmp_path / "cm_form.xlsx"
        wb = cm.create_workbook()
        cm.add_sheet(wb, "S1")
        cm.write_cell_value(
            wb,
            "S1",
            "A1",
            CellValue(type=CellType.FORMULA, value="=1+1", formula="=1+1"),
        )
        cm.save_workbook(wb, path)

        rb = opxl.open_workbook(path)
        cv = opxl.read_cell_value(rb, "S1", "A1")
        assert cv.type == CellType.FORMULA
        opxl.close_workbook(rb)

    def test_out_of_order_writes_sorted(
        self, cm: XlsxwriterConstmemAdapter, opxl: OpenpyxlAdapter, tmp_path: Path
    ) -> None:
        """Verify that writes queued out of order are sorted for constant_memory mode."""
        path = tmp_path / "cm_sort.xlsx"
        wb = cm.create_workbook()
        cm.add_sheet(wb, "S1")
        # Write in reverse order — should still work
        cm.write_cell_value(wb, "S1", "A3", CellValue(type=CellType.STRING, value="third"))
        cm.write_cell_value(wb, "S1", "A1", CellValue(type=CellType.STRING, value="first"))
        cm.write_cell_value(wb, "S1", "A2", CellValue(type=CellType.STRING, value="second"))
        cm.save_workbook(wb, path)

        rb = opxl.open_workbook(path)
        assert opxl.read_cell_value(rb, "S1", "A1").value == "first"
        assert opxl.read_cell_value(rb, "S1", "A2").value == "second"
        assert opxl.read_cell_value(rb, "S1", "A3").value == "third"
        opxl.close_workbook(rb)

    def test_multiple_sheets(
        self, cm: XlsxwriterConstmemAdapter, opxl: OpenpyxlAdapter, tmp_path: Path
    ) -> None:
        path = tmp_path / "cm_multi.xlsx"
        wb = cm.create_workbook()
        cm.add_sheet(wb, "S1")
        cm.add_sheet(wb, "S2")
        cm.write_cell_value(wb, "S1", "A1", CellValue(type=CellType.STRING, value="sheet1"))
        cm.write_cell_value(wb, "S2", "A1", CellValue(type=CellType.STRING, value="sheet2"))
        cm.save_workbook(wb, path)

        rb = opxl.open_workbook(path)
        assert opxl.read_cell_value(rb, "S1", "A1").value == "sheet1"
        assert opxl.read_cell_value(rb, "S2", "A1").value == "sheet2"
        opxl.close_workbook(rb)

    def test_formatting(
        self, cm: XlsxwriterConstmemAdapter, opxl: OpenpyxlAdapter, tmp_path: Path
    ) -> None:
        path = tmp_path / "cm_fmt.xlsx"
        wb = cm.create_workbook()
        cm.add_sheet(wb, "S1")
        cm.write_cell_value(wb, "S1", "A1", CellValue(type=CellType.STRING, value="bold"))
        cm.write_cell_format(wb, "S1", "A1", CellFormat(bold=True))
        cm.save_workbook(wb, path)

        rb = opxl.open_workbook(path)
        fmt = opxl.read_cell_format(rb, "S1", "A1")
        assert fmt.bold is True
        opxl.close_workbook(rb)

    def test_border(
        self, cm: XlsxwriterConstmemAdapter, opxl: OpenpyxlAdapter, tmp_path: Path
    ) -> None:
        path = tmp_path / "cm_bdr.xlsx"
        wb = cm.create_workbook()
        cm.add_sheet(wb, "S1")
        cm.write_cell_value(wb, "S1", "A1", CellValue(type=CellType.STRING, value="x"))
        cm.write_cell_border(
            wb, "S1", "A1",
            BorderInfo(top=BorderEdge(style=BorderStyle.THIN, color="#000000")),
        )
        cm.save_workbook(wb, path)

        rb = opxl.open_workbook(path)
        border = opxl.read_cell_border(rb, "S1", "A1")
        assert border.top is not None
        assert border.top.style == BorderStyle.THIN
        opxl.close_workbook(rb)


# ═════════════════════════════════════════════════════════════════════════
# TestConstmemReadRaises
# ═════════════════════════════════════════════════════════════════════════


class TestConstmemReadRaises:
    """Read operations should raise NotImplementedError (write-only)."""

    def test_open_workbook(self, cm: XlsxwriterConstmemAdapter, tmp_path: Path) -> None:
        with pytest.raises(NotImplementedError):
            cm.open_workbook(tmp_path / "nonexistent.xlsx")

    def test_read_cell_value(self, cm: XlsxwriterConstmemAdapter) -> None:
        with pytest.raises(NotImplementedError):
            cm.read_cell_value(None, "S1", "A1")
