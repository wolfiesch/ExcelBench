"""Tests for python-calamine and pylightxl adapters."""

import tempfile
from datetime import date
from pathlib import Path

import pytest

from excelbench.harness.adapters import CalamineAdapter, PylightxlAdapter
from excelbench.harness.adapters.openpyxl_adapter import OpenpyxlAdapter
from excelbench.models import BorderInfo, CellFormat, CellType, CellValue

# =========================================================================
# Fixtures
# =========================================================================


@pytest.fixture
def openpyxl_adapter():
    return OpenpyxlAdapter()


@pytest.fixture
def calamine_adapter():
    return CalamineAdapter()


@pytest.fixture
def pylightxl_adapter():
    return PylightxlAdapter()


@pytest.fixture
def sample_xlsx(openpyxl_adapter):
    """Create a sample xlsx with various cell types via openpyxl."""
    path = Path(tempfile.mktemp(suffix=".xlsx"))
    wb = openpyxl_adapter.create_workbook()
    openpyxl_adapter.add_sheet(wb, "Types")

    openpyxl_adapter.write_cell_value(
        wb, "Types", "A1", CellValue(type=CellType.STRING, value="hello")
    )
    openpyxl_adapter.write_cell_value(
        wb, "Types", "A2", CellValue(type=CellType.NUMBER, value=42)
    )
    openpyxl_adapter.write_cell_value(
        wb, "Types", "A3", CellValue(type=CellType.NUMBER, value=3.14)
    )
    openpyxl_adapter.write_cell_value(
        wb, "Types", "A4", CellValue(type=CellType.BOOLEAN, value=True)
    )
    openpyxl_adapter.write_cell_value(
        wb, "Types", "A5", CellValue(type=CellType.BOOLEAN, value=False)
    )
    openpyxl_adapter.write_cell_value(
        wb, "Types", "A6", CellValue(type=CellType.BLANK)
    )
    openpyxl_adapter.write_cell_value(
        wb, "Types", "A7",
        CellValue(type=CellType.FORMULA, value="=A2*2", formula="=A2*2"),
    )
    openpyxl_adapter.write_cell_value(
        wb, "Types", "A8",
        CellValue(type=CellType.DATE, value=date(2024, 6, 15)),
    )

    # Second sheet for multi-sheet test
    openpyxl_adapter.add_sheet(wb, "Sheet2")
    openpyxl_adapter.write_cell_value(
        wb, "Sheet2", "B1", CellValue(type=CellType.STRING, value="second")
    )

    openpyxl_adapter.save_workbook(wb, path)
    yield path
    path.unlink(missing_ok=True)


# =========================================================================
# CalamineAdapter Tests
# =========================================================================


class TestCalamineInfo:
    def test_name(self, calamine_adapter):
        assert calamine_adapter.name == "python-calamine"

    def test_capabilities(self, calamine_adapter):
        assert calamine_adapter.can_read() is True
        assert calamine_adapter.can_write() is False

    def test_info_language(self, calamine_adapter):
        assert calamine_adapter.info.language == "python"


class TestCalamineRead:
    def test_sheet_names(self, calamine_adapter, sample_xlsx):
        wb = calamine_adapter.open_workbook(sample_xlsx)
        names = calamine_adapter.get_sheet_names(wb)
        assert "Types" in names
        assert "Sheet2" in names
        calamine_adapter.close_workbook(wb)

    def test_read_string(self, calamine_adapter, sample_xlsx):
        wb = calamine_adapter.open_workbook(sample_xlsx)
        val = calamine_adapter.read_cell_value(wb, "Types", "A1")
        assert val.type == CellType.STRING
        assert val.value == "hello"
        calamine_adapter.close_workbook(wb)

    def test_read_number_int(self, calamine_adapter, sample_xlsx):
        wb = calamine_adapter.open_workbook(sample_xlsx)
        val = calamine_adapter.read_cell_value(wb, "Types", "A2")
        assert val.type == CellType.NUMBER
        assert val.value == 42 or val.value == 42.0
        calamine_adapter.close_workbook(wb)

    def test_read_number_float(self, calamine_adapter, sample_xlsx):
        wb = calamine_adapter.open_workbook(sample_xlsx)
        val = calamine_adapter.read_cell_value(wb, "Types", "A3")
        assert val.type == CellType.NUMBER
        assert abs(val.value - 3.14) < 0.001
        calamine_adapter.close_workbook(wb)

    def test_read_boolean_true(self, calamine_adapter, sample_xlsx):
        wb = calamine_adapter.open_workbook(sample_xlsx)
        val = calamine_adapter.read_cell_value(wb, "Types", "A4")
        assert val.type == CellType.BOOLEAN
        assert val.value is True
        calamine_adapter.close_workbook(wb)

    def test_read_boolean_false(self, calamine_adapter, sample_xlsx):
        wb = calamine_adapter.open_workbook(sample_xlsx)
        val = calamine_adapter.read_cell_value(wb, "Types", "A5")
        assert val.type == CellType.BOOLEAN
        assert val.value is False
        calamine_adapter.close_workbook(wb)

    def test_read_blank(self, calamine_adapter, sample_xlsx):
        wb = calamine_adapter.open_workbook(sample_xlsx)
        val = calamine_adapter.read_cell_value(wb, "Types", "A6")
        assert val.type == CellType.BLANK
        calamine_adapter.close_workbook(wb)

    def test_read_out_of_bounds_blank(self, calamine_adapter, sample_xlsx):
        wb = calamine_adapter.open_workbook(sample_xlsx)
        val = calamine_adapter.read_cell_value(wb, "Types", "Z99")
        assert val.type == CellType.BLANK
        calamine_adapter.close_workbook(wb)

    def test_read_date(self, calamine_adapter, sample_xlsx):
        wb = calamine_adapter.open_workbook(sample_xlsx)
        val = calamine_adapter.read_cell_value(wb, "Types", "A8")
        # calamine may return DATE or NUMBER depending on how it interprets
        # the cell. Both are valid outcomes for a read-only library.
        assert val.type in (CellType.DATE, CellType.DATETIME, CellType.NUMBER)
        calamine_adapter.close_workbook(wb)

    def test_formatting_returns_empty(self, calamine_adapter, sample_xlsx):
        wb = calamine_adapter.open_workbook(sample_xlsx)
        fmt = calamine_adapter.read_cell_format(wb, "Types", "A1")
        assert isinstance(fmt, CellFormat)
        assert fmt.bold is None
        calamine_adapter.close_workbook(wb)

    def test_border_returns_empty(self, calamine_adapter, sample_xlsx):
        wb = calamine_adapter.open_workbook(sample_xlsx)
        border = calamine_adapter.read_cell_border(wb, "Types", "A1")
        assert isinstance(border, BorderInfo)
        assert border.top is None
        calamine_adapter.close_workbook(wb)

    def test_dimensions_return_none(self, calamine_adapter, sample_xlsx):
        wb = calamine_adapter.open_workbook(sample_xlsx)
        assert calamine_adapter.read_row_height(wb, "Types", 1) is None
        assert calamine_adapter.read_column_width(wb, "Types", "A") is None
        calamine_adapter.close_workbook(wb)

    def test_tier2_returns_empty(self, calamine_adapter, sample_xlsx):
        wb = calamine_adapter.open_workbook(sample_xlsx)
        assert calamine_adapter.read_merged_ranges(wb, "Types") == []
        assert calamine_adapter.read_conditional_formats(wb, "Types") == []
        assert calamine_adapter.read_data_validations(wb, "Types") == []
        assert calamine_adapter.read_hyperlinks(wb, "Types") == []
        assert calamine_adapter.read_images(wb, "Types") == []
        assert calamine_adapter.read_pivot_tables(wb, "Types") == []
        assert calamine_adapter.read_comments(wb, "Types") == []
        assert calamine_adapter.read_freeze_panes(wb, "Types") == {}
        calamine_adapter.close_workbook(wb)


class TestCalamineWriteBlocked:
    def test_create_workbook_raises(self, calamine_adapter):
        with pytest.raises(NotImplementedError, match="read-only"):
            calamine_adapter.create_workbook()

    def test_save_workbook_raises(self, calamine_adapter):
        with pytest.raises(NotImplementedError, match="read-only"):
            calamine_adapter.save_workbook(None, Path("/tmp/test.xlsx"))


# =========================================================================
# PylightxlAdapter Tests
# =========================================================================


class TestPylightxlInfo:
    def test_name(self, pylightxl_adapter):
        assert pylightxl_adapter.name == "pylightxl"

    def test_capabilities(self, pylightxl_adapter):
        assert pylightxl_adapter.can_read() is True
        assert pylightxl_adapter.can_write() is True

    def test_info_language(self, pylightxl_adapter):
        assert pylightxl_adapter.info.language == "python"


class TestPylightxlWriteRead:
    """Test pylightxl write→read roundtrip (self-consistency)."""

    def test_string_roundtrip(self, pylightxl_adapter):
        path = Path(tempfile.mktemp(suffix=".xlsx"))
        try:
            wb = pylightxl_adapter.create_workbook()
            pylightxl_adapter.add_sheet(wb, "S1")
            pylightxl_adapter.write_cell_value(
                wb, "S1", "A1", CellValue(type=CellType.STRING, value="test")
            )
            pylightxl_adapter.save_workbook(wb, path)

            wb2 = pylightxl_adapter.open_workbook(path)
            val = pylightxl_adapter.read_cell_value(wb2, "S1", "A1")
            assert val.type == CellType.STRING
            assert val.value == "test"
        finally:
            path.unlink(missing_ok=True)

    def test_number_roundtrip(self, pylightxl_adapter):
        path = Path(tempfile.mktemp(suffix=".xlsx"))
        try:
            wb = pylightxl_adapter.create_workbook()
            pylightxl_adapter.add_sheet(wb, "S1")
            pylightxl_adapter.write_cell_value(
                wb, "S1", "A1", CellValue(type=CellType.NUMBER, value=42)
            )
            pylightxl_adapter.write_cell_value(
                wb, "S1", "A2", CellValue(type=CellType.NUMBER, value=3.14)
            )
            pylightxl_adapter.save_workbook(wb, path)

            wb2 = pylightxl_adapter.open_workbook(path)
            val1 = pylightxl_adapter.read_cell_value(wb2, "S1", "A1")
            assert val1.type == CellType.NUMBER
            assert val1.value == 42

            val2 = pylightxl_adapter.read_cell_value(wb2, "S1", "A2")
            assert val2.type == CellType.NUMBER
            assert abs(val2.value - 3.14) < 0.001
        finally:
            path.unlink(missing_ok=True)

    def test_blank_cell(self, pylightxl_adapter):
        path = Path(tempfile.mktemp(suffix=".xlsx"))
        try:
            wb = pylightxl_adapter.create_workbook()
            pylightxl_adapter.add_sheet(wb, "S1")
            pylightxl_adapter.write_cell_value(
                wb, "S1", "A1", CellValue(type=CellType.STRING, value="data")
            )
            pylightxl_adapter.save_workbook(wb, path)

            wb2 = pylightxl_adapter.open_workbook(path)
            val = pylightxl_adapter.read_cell_value(wb2, "S1", "B1")
            assert val.type == CellType.BLANK
        finally:
            path.unlink(missing_ok=True)

    def test_multiple_sheets(self, pylightxl_adapter):
        path = Path(tempfile.mktemp(suffix=".xlsx"))
        try:
            wb = pylightxl_adapter.create_workbook()
            pylightxl_adapter.add_sheet(wb, "Alpha")
            pylightxl_adapter.add_sheet(wb, "Beta")
            pylightxl_adapter.write_cell_value(
                wb, "Alpha", "A1", CellValue(type=CellType.STRING, value="first")
            )
            pylightxl_adapter.write_cell_value(
                wb, "Beta", "A1", CellValue(type=CellType.STRING, value="second")
            )
            pylightxl_adapter.save_workbook(wb, path)

            wb2 = pylightxl_adapter.open_workbook(path)
            sheets = pylightxl_adapter.get_sheet_names(wb2)
            assert "Alpha" in sheets
            assert "Beta" in sheets

            v1 = pylightxl_adapter.read_cell_value(wb2, "Alpha", "A1")
            assert v1.value == "first"
            v2 = pylightxl_adapter.read_cell_value(wb2, "Beta", "A1")
            assert v2.value == "second"
        finally:
            path.unlink(missing_ok=True)

    def test_save_overwrites_existing(self, pylightxl_adapter):
        """Verify save works even when the target file already exists."""
        path = Path(tempfile.mktemp(suffix=".xlsx"))
        try:
            # First write
            wb = pylightxl_adapter.create_workbook()
            pylightxl_adapter.add_sheet(wb, "S1")
            pylightxl_adapter.write_cell_value(
                wb, "S1", "A1", CellValue(type=CellType.STRING, value="v1")
            )
            pylightxl_adapter.save_workbook(wb, path)

            # Second write to same path
            wb2 = pylightxl_adapter.create_workbook()
            pylightxl_adapter.add_sheet(wb2, "S1")
            pylightxl_adapter.write_cell_value(
                wb2, "S1", "A1", CellValue(type=CellType.STRING, value="v2")
            )
            pylightxl_adapter.save_workbook(wb2, path)

            wb3 = pylightxl_adapter.open_workbook(path)
            val = pylightxl_adapter.read_cell_value(wb3, "S1", "A1")
            assert val.value == "v2"
        finally:
            path.unlink(missing_ok=True)


class TestPylightxlFormatNoop:
    """Verify formatting methods don't crash (they're no-ops)."""

    def test_write_format_noop(self, pylightxl_adapter):
        wb = pylightxl_adapter.create_workbook()
        pylightxl_adapter.add_sheet(wb, "S1")
        pylightxl_adapter.write_cell_format(wb, "S1", "A1", CellFormat(bold=True))

    def test_write_border_noop(self, pylightxl_adapter):
        wb = pylightxl_adapter.create_workbook()
        pylightxl_adapter.add_sheet(wb, "S1")
        pylightxl_adapter.write_cell_border(wb, "S1", "A1", BorderInfo())

    def test_tier2_write_noops(self, pylightxl_adapter):
        wb = pylightxl_adapter.create_workbook()
        pylightxl_adapter.add_sheet(wb, "S1")
        pylightxl_adapter.merge_cells(wb, "S1", "A1:B1")
        pylightxl_adapter.add_conditional_format(wb, "S1", {})
        pylightxl_adapter.add_data_validation(wb, "S1", {})
        pylightxl_adapter.add_hyperlink(wb, "S1", {})
        pylightxl_adapter.add_image(wb, "S1", {})
        pylightxl_adapter.add_pivot_table(wb, "S1", {})
        pylightxl_adapter.add_comment(wb, "S1", {})
        pylightxl_adapter.set_freeze_panes(wb, "S1", {})


# =========================================================================
# Cross-adapter Tests
# =========================================================================


class TestCrossAdapterRead:
    """Test that calamine and pylightxl files are readable cross-adapter."""

    def test_calamine_reads_pylightxl_file(self, calamine_adapter, pylightxl_adapter):
        path = Path(tempfile.mktemp(suffix=".xlsx"))
        try:
            wb = pylightxl_adapter.create_workbook()
            pylightxl_adapter.add_sheet(wb, "Cross")
            pylightxl_adapter.write_cell_value(
                wb, "Cross", "A1", CellValue(type=CellType.STRING, value="cross-test")
            )
            pylightxl_adapter.write_cell_value(
                wb, "Cross", "B1", CellValue(type=CellType.NUMBER, value=7)
            )
            pylightxl_adapter.save_workbook(wb, path)

            wb2 = calamine_adapter.open_workbook(path)
            sheets = calamine_adapter.get_sheet_names(wb2)
            assert "Cross" in sheets

            val = calamine_adapter.read_cell_value(wb2, "Cross", "A1")
            assert val.type == CellType.STRING
            assert val.value == "cross-test"

            val2 = calamine_adapter.read_cell_value(wb2, "Cross", "B1")
            assert val2.type == CellType.NUMBER
            assert val2.value == 7 or val2.value == 7.0
            calamine_adapter.close_workbook(wb2)
        finally:
            path.unlink(missing_ok=True)
