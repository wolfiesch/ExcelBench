"""Tests for tables feature (Tier 3)."""

from __future__ import annotations

from pathlib import Path

import pytest

from excelbench.harness.adapters.openpyxl_adapter import OpenpyxlAdapter
from excelbench.models import CellType, CellValue
from excelbench.test_support import StubExcelAdapter

FIXTURE = Path("fixtures/excel/tier3/19_tables.xlsx")


class _StubAdapter(StubExcelAdapter):
    pass


class TestTablesBase:
    """Base adapter API surface for tables."""

    def test_read_tables_default_returns_empty(self) -> None:
        adapter = _StubAdapter()
        assert adapter.read_tables(object(), "S1") == []

    def test_add_table_default_is_noop(self) -> None:
        adapter = _StubAdapter()
        adapter.add_table(object(), "S1", {"table": {"name": "T", "ref": "A1:B2"}})


class TestOpenpyxlTables:
    def test_roundtrip_table_in_memory(self, tmp_path: Path) -> None:
        adapter = OpenpyxlAdapter()
        wb = adapter.create_workbook()
        adapter.add_sheet(wb, "tables")

        adapter.write_cell_value(
            wb,
            "tables",
            "A1",
            CellValue(type=CellType.STRING, value="Name"),
        )
        adapter.write_cell_value(
            wb,
            "tables",
            "B1",
            CellValue(type=CellType.STRING, value="Qty"),
        )
        adapter.write_cell_value(
            wb,
            "tables",
            "A2",
            CellValue(type=CellType.STRING, value="X"),
        )
        adapter.write_cell_value(
            wb,
            "tables",
            "B2",
            CellValue(type=CellType.NUMBER, value=10),
        )

        adapter.add_table(
            wb,
            "tables",
            {
                "table": {
                    "name": "TestTable",
                    "ref": "A1:B2",
                    "style": "TableStyleMedium9",
                    "columns": ["Name", "Qty"],
                }
            },
        )

        path = tmp_path / "tables.xlsx"
        adapter.save_workbook(wb, path)

        wb2 = adapter.open_workbook(path)
        try:
            tables = adapter.read_tables(wb2, "tables")
            assert isinstance(tables, list)
            assert len(tables) == 1
            assert tables[0]["name"] == "TestTable"
            assert tables[0]["ref"] == "A1:B2"
            assert tables[0]["columns"] == ["Name", "Qty"]
        finally:
            adapter.close_workbook(wb2)

    @pytest.mark.skipif(not FIXTURE.exists(), reason="Tables fixture not generated yet")
    def test_read_tables_returns_list_from_fixture(self) -> None:
        adapter = OpenpyxlAdapter()
        wb = adapter.open_workbook(FIXTURE)
        try:
            tables = adapter.read_tables(wb, "tables")
            assert isinstance(tables, list)
            assert len(tables) >= 1
        finally:
            adapter.close_workbook(wb)
