"""Tests for named ranges feature (Tier 3)."""

from __future__ import annotations

from pathlib import Path

import pytest

from excelbench.harness.adapters.openpyxl_adapter import OpenpyxlAdapter
from excelbench.models import CellType, CellValue
from excelbench.test_support import StubExcelAdapter

FIXTURE = Path("fixtures/excel/tier3/18_named_ranges.xlsx")


class TestNamedRangesBase:
    """Base adapter API surface for named ranges."""

    def test_read_named_ranges_default_returns_empty(self) -> None:
        adapter = StubExcelAdapter()
        assert adapter.read_named_ranges(object(), "S1") == []

    def test_add_named_range_default_is_noop(self) -> None:
        adapter = StubExcelAdapter()
        adapter.add_named_range(object(), "S1", {"name": "X", "refers_to": "S1!$A$1"})


class TestOpenpyxlNamedRanges:
    def test_roundtrip_named_ranges_in_memory(self, tmp_path: Path) -> None:
        adapter = OpenpyxlAdapter()
        wb = adapter.create_workbook()
        adapter.add_sheet(wb, "named_ranges")
        adapter.add_sheet(wb, "Targets")
        adapter.write_cell_value(
            wb, "named_ranges", "B2", CellValue(type=CellType.NUMBER, value=42)
        )

        adapter.add_named_range(
            wb,
            "named_ranges",
            {"name": "SingleCell", "scope": "workbook", "refers_to": "named_ranges!$B$2"},
        )
        adapter.add_named_range(
            wb,
            "named_ranges",
            {"name": "LocalName", "scope": "sheet", "refers_to": "named_ranges!$B$2"},
        )

        path = tmp_path / "named_ranges.xlsx"
        adapter.save_workbook(wb, path)

        wb2 = adapter.open_workbook(path)
        try:
            names = adapter.read_named_ranges(wb2, "named_ranges")
            assert isinstance(names, list)
            assert any(
                n.get("name") == "SingleCell" and n.get("scope") == "workbook" for n in names
            )
            assert any(n.get("name") == "LocalName" and n.get("scope") == "sheet" for n in names)
        finally:
            adapter.close_workbook(wb2)

    @pytest.mark.skipif(not FIXTURE.exists(), reason="Named ranges fixture not generated yet")
    def test_read_named_ranges_returns_list_from_fixture(self) -> None:
        adapter = OpenpyxlAdapter()
        wb = adapter.open_workbook(FIXTURE)
        try:
            names = adapter.read_named_ranges(wb, "named_ranges")
            assert isinstance(names, list)
            assert len(names) >= 1
        finally:
            adapter.close_workbook(wb)
