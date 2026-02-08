"""Tests for TablibAdapter (read+write, value-only via tablib.Dataset)."""

from __future__ import annotations

from datetime import date, datetime
from pathlib import Path

import pytest

from excelbench.harness.adapters.openpyxl_adapter import OpenpyxlAdapter
from excelbench.harness.adapters.tablib_adapter import TablibAdapter
from excelbench.models import (
    BorderInfo,
    CellFormat,
    CellType,
    CellValue,
)


@pytest.fixture
def opxl() -> OpenpyxlAdapter:
    return OpenpyxlAdapter()


@pytest.fixture
def tbl() -> TablibAdapter:
    return TablibAdapter()


def _write_openpyxl_fixture(opxl: OpenpyxlAdapter, path: Path) -> None:
    """Write a multi-type .xlsx fixture using openpyxl for read tests."""
    wb = opxl.create_workbook()
    opxl.add_sheet(wb, "S1")
    opxl.write_cell_value(wb, "S1", "A1", CellValue(type=CellType.STRING, value="hello"))
    opxl.write_cell_value(wb, "S1", "A2", CellValue(type=CellType.NUMBER, value=42.5))
    opxl.write_cell_value(wb, "S1", "A3", CellValue(type=CellType.BOOLEAN, value=True))
    opxl.write_cell_value(wb, "S1", "A4", CellValue(type=CellType.DATE, value=date(2024, 6, 15)))
    opxl.write_cell_value(
        wb,
        "S1",
        "A5",
        CellValue(type=CellType.DATETIME, value=datetime(2024, 6, 15, 14, 30, 0)),
    )
    opxl.write_cell_value(wb, "S1", "A6", CellValue(type=CellType.ERROR, value="#N/A"))
    opxl.write_cell_value(
        wb, "S1", "A7", CellValue(type=CellType.FORMULA, value="=1+1", formula="=1+1")
    )
    opxl.write_cell_value(wb, "S1", "A8", CellValue(type=CellType.BLANK))
    opxl.save_workbook(wb, path)


# ═════════════════════════════════════════════════════════════════════════
# TestTablibInfo
# ═════════════════════════════════════════════════════════════════════════


class TestTablibInfo:
    def test_name(self, tbl: TablibAdapter) -> None:
        assert tbl.info.name == "tablib"

    def test_version(self, tbl: TablibAdapter) -> None:
        assert tbl.info.version != "unknown"

    def test_capabilities(self, tbl: TablibAdapter) -> None:
        assert "read" in tbl.info.capabilities
        assert "write" in tbl.info.capabilities

    def test_language(self, tbl: TablibAdapter) -> None:
        assert tbl.info.language == "python"


# ═════════════════════════════════════════════════════════════════════════
# TestTablibReadCellValue
# ═════════════════════════════════════════════════════════════════════════


class TestTablibReadCellValue:
    """Read via tablib from openpyxl-written fixtures."""

    def test_string(self, tbl: TablibAdapter, opxl: OpenpyxlAdapter, tmp_path: Path) -> None:
        path = tmp_path / "fixture.xlsx"
        _write_openpyxl_fixture(opxl, path)
        wb = tbl.open_workbook(path)
        cv = tbl.read_cell_value(wb, "S1", "A1")
        assert cv.type == CellType.STRING
        assert cv.value == "hello"
        tbl.close_workbook(wb)

    def test_number(self, tbl: TablibAdapter, opxl: OpenpyxlAdapter, tmp_path: Path) -> None:
        path = tmp_path / "fixture.xlsx"
        _write_openpyxl_fixture(opxl, path)
        wb = tbl.open_workbook(path)
        cv = tbl.read_cell_value(wb, "S1", "A2")
        assert cv.type == CellType.NUMBER
        assert cv.value == 42.5
        tbl.close_workbook(wb)

    def test_boolean(self, tbl: TablibAdapter, opxl: OpenpyxlAdapter, tmp_path: Path) -> None:
        path = tmp_path / "fixture.xlsx"
        _write_openpyxl_fixture(opxl, path)
        wb = tbl.open_workbook(path)
        cv = tbl.read_cell_value(wb, "S1", "A3")
        assert cv.type == CellType.BOOLEAN
        assert cv.value is True
        tbl.close_workbook(wb)

    def test_date(self, tbl: TablibAdapter, opxl: OpenpyxlAdapter, tmp_path: Path) -> None:
        path = tmp_path / "fixture.xlsx"
        _write_openpyxl_fixture(opxl, path)
        wb = tbl.open_workbook(path)
        cv = tbl.read_cell_value(wb, "S1", "A4")
        assert cv.type == CellType.DATE
        assert cv.value == date(2024, 6, 15)
        tbl.close_workbook(wb)

    def test_datetime(self, tbl: TablibAdapter, opxl: OpenpyxlAdapter, tmp_path: Path) -> None:
        path = tmp_path / "fixture.xlsx"
        _write_openpyxl_fixture(opxl, path)
        wb = tbl.open_workbook(path)
        cv = tbl.read_cell_value(wb, "S1", "A5")
        assert cv.type == CellType.DATETIME
        tbl.close_workbook(wb)

    def test_error(self, tbl: TablibAdapter, opxl: OpenpyxlAdapter, tmp_path: Path) -> None:
        path = tmp_path / "fixture.xlsx"
        _write_openpyxl_fixture(opxl, path)
        wb = tbl.open_workbook(path)
        cv = tbl.read_cell_value(wb, "S1", "A6")
        # tablib may return error as string or formula (openpyxl error not preserved)
        assert cv.type in (CellType.ERROR, CellType.STRING, CellType.FORMULA, CellType.BLANK)
        tbl.close_workbook(wb)

    def test_blank_out_of_bounds(
        self, tbl: TablibAdapter, opxl: OpenpyxlAdapter, tmp_path: Path
    ) -> None:
        path = tmp_path / "fixture.xlsx"
        _write_openpyxl_fixture(opxl, path)
        wb = tbl.open_workbook(path)
        cv = tbl.read_cell_value(wb, "S1", "A99")
        assert cv.type == CellType.BLANK
        cv2 = tbl.read_cell_value(wb, "S1", "Z1")
        assert cv2.type == CellType.BLANK
        tbl.close_workbook(wb)

    def test_sheet_not_found(
        self, tbl: TablibAdapter, opxl: OpenpyxlAdapter, tmp_path: Path
    ) -> None:
        path = tmp_path / "fixture.xlsx"
        _write_openpyxl_fixture(opxl, path)
        wb = tbl.open_workbook(path)
        cv = tbl.read_cell_value(wb, "NoSheet", "A1")
        assert cv.type == CellType.BLANK
        tbl.close_workbook(wb)

    def test_sheet_names(
        self, tbl: TablibAdapter, opxl: OpenpyxlAdapter, tmp_path: Path
    ) -> None:
        path = tmp_path / "fixture.xlsx"
        _write_openpyxl_fixture(opxl, path)
        wb = tbl.open_workbook(path)
        names = tbl.get_sheet_names(wb)
        assert "S1" in names
        tbl.close_workbook(wb)


# ═════════════════════════════════════════════════════════════════════════
# TestTablibReadStubs
# ═════════════════════════════════════════════════════════════════════════


class TestTablibReadStubs:
    """Tier-2 reads all return empty."""

    def test_all_stubs(self, tbl: TablibAdapter, opxl: OpenpyxlAdapter, tmp_path: Path) -> None:
        path = tmp_path / "fixture.xlsx"
        _write_openpyxl_fixture(opxl, path)
        wb = tbl.open_workbook(path)
        assert tbl.read_cell_format(wb, "S1", "A1") == CellFormat()
        assert tbl.read_cell_border(wb, "S1", "A1") == BorderInfo()
        assert tbl.read_merged_ranges(wb, "S1") == []
        assert tbl.read_conditional_formats(wb, "S1") == []
        assert tbl.read_data_validations(wb, "S1") == []
        assert tbl.read_hyperlinks(wb, "S1") == []
        assert tbl.read_images(wb, "S1") == []
        assert tbl.read_pivot_tables(wb, "S1") == []
        assert tbl.read_comments(wb, "S1") == []
        assert tbl.read_freeze_panes(wb, "S1") == {}
        assert tbl.read_row_height(wb, "S1", 1) is None
        assert tbl.read_column_width(wb, "S1", "A") is None
        tbl.close_workbook(wb)


# ═════════════════════════════════════════════════════════════════════════
# TestTablibWriteRoundtrip
# ═════════════════════════════════════════════════════════════════════════


class TestTablibWriteRoundtrip:
    """Write via tablib, read back via openpyxl."""

    def test_string(self, tbl: TablibAdapter, opxl: OpenpyxlAdapter, tmp_path: Path) -> None:
        path = tmp_path / "tbl_out.xlsx"
        wb = tbl.create_workbook()
        tbl.add_sheet(wb, "S1")
        tbl.write_cell_value(wb, "S1", "A1", CellValue(type=CellType.STRING, value="hi"))
        tbl.save_workbook(wb, path)

        rb = opxl.open_workbook(path)
        cv = opxl.read_cell_value(rb, "S1", "A1")
        assert cv.value == "hi"
        opxl.close_workbook(rb)

    def test_number(self, tbl: TablibAdapter, opxl: OpenpyxlAdapter, tmp_path: Path) -> None:
        path = tmp_path / "tbl_num.xlsx"
        wb = tbl.create_workbook()
        tbl.add_sheet(wb, "S1")
        tbl.write_cell_value(wb, "S1", "A1", CellValue(type=CellType.NUMBER, value=99))
        tbl.save_workbook(wb, path)

        rb = opxl.open_workbook(path)
        cv = opxl.read_cell_value(rb, "S1", "A1")
        assert cv.value == 99
        opxl.close_workbook(rb)

    def test_boolean(self, tbl: TablibAdapter, opxl: OpenpyxlAdapter, tmp_path: Path) -> None:
        path = tmp_path / "tbl_bool.xlsx"
        wb = tbl.create_workbook()
        tbl.add_sheet(wb, "S1")
        tbl.write_cell_value(wb, "S1", "A1", CellValue(type=CellType.BOOLEAN, value=False))
        tbl.save_workbook(wb, path)

        rb = opxl.open_workbook(path)
        cv = opxl.read_cell_value(rb, "S1", "A1")
        assert cv.value is False
        opxl.close_workbook(rb)

    def test_empty_sheet(
        self, tbl: TablibAdapter, opxl: OpenpyxlAdapter, tmp_path: Path
    ) -> None:
        path = tmp_path / "tbl_empty.xlsx"
        wb = tbl.create_workbook()
        tbl.add_sheet(wb, "EmptySheet")
        tbl.save_workbook(wb, path)

        rb = opxl.open_workbook(path)
        assert "EmptySheet" in opxl.get_sheet_names(rb)
        opxl.close_workbook(rb)

    def test_write_auto_creates_sheet(self, tbl: TablibAdapter) -> None:
        wb = tbl.create_workbook()
        tbl.write_cell_value(wb, "Auto", "A1", CellValue(type=CellType.STRING, value="x"))
        assert "Auto" in wb["sheets"]


# ═════════════════════════════════════════════════════════════════════════
# TestTablibWriteNoops
# ═════════════════════════════════════════════════════════════════════════


class TestTablibWriteNoops:
    """Formatting/tier-2 writes don't raise."""

    def test_noop_methods(self, tbl: TablibAdapter) -> None:
        wb = tbl.create_workbook()
        tbl.add_sheet(wb, "S1")
        tbl.write_cell_format(wb, "S1", "A1", CellFormat())
        tbl.write_cell_border(wb, "S1", "A1", BorderInfo())
        tbl.set_row_height(wb, "S1", 1, 30.0)
        tbl.set_column_width(wb, "S1", "A", 20.0)
        tbl.merge_cells(wb, "S1", "A1:B2")
        tbl.add_conditional_format(wb, "S1", {})
        tbl.add_data_validation(wb, "S1", {})
        tbl.add_hyperlink(wb, "S1", {})
        tbl.add_image(wb, "S1", {})
        tbl.add_pivot_table(wb, "S1", {})
        tbl.add_comment(wb, "S1", {})
        tbl.set_freeze_panes(wb, "S1", {})
