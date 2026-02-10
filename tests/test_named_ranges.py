"""Tests for named ranges feature (Tier 3)."""

from __future__ import annotations

from pathlib import Path
from typing import Any

import pytest

from excelbench.harness.adapters.base import ExcelAdapter
from excelbench.harness.adapters.openpyxl_adapter import OpenpyxlAdapter
from excelbench.models import BorderInfo, CellFormat, CellType, CellValue, LibraryInfo

FIXTURE = Path("fixtures/excel/tier3/18_named_ranges.xlsx")


class _StubAdapter(ExcelAdapter):
    """Concrete adapter to exercise base-class defaults."""

    @property
    def info(self) -> LibraryInfo:
        return LibraryInfo(
            name="stub",
            version="0",
            language="python",
            capabilities={"read", "write"},
        )

    # Read
    def open_workbook(self, path: Path) -> Any:  # pragma: no cover
        raise NotImplementedError

    def close_workbook(self, workbook: Any) -> None:  # pragma: no cover
        raise NotImplementedError

    def get_sheet_names(self, workbook: Any) -> list[str]:  # pragma: no cover
        raise NotImplementedError

    def read_cell_value(
        self, workbook: Any, sheet: str, cell: str
    ) -> CellValue:  # pragma: no cover
        return CellValue(type=CellType.BLANK)

    def read_cell_format(
        self, workbook: Any, sheet: str, cell: str
    ) -> CellFormat:  # pragma: no cover
        return CellFormat()

    def read_cell_border(
        self, workbook: Any, sheet: str, cell: str
    ) -> BorderInfo:  # pragma: no cover
        return BorderInfo()

    def read_row_height(
        self, workbook: Any, sheet: str, row: int
    ) -> float | None:  # pragma: no cover
        return None

    def read_column_width(
        self, workbook: Any, sheet: str, column: str
    ) -> float | None:  # pragma: no cover
        return None

    def read_merged_ranges(self, workbook: Any, sheet: str) -> list[str]:  # pragma: no cover
        return []

    def read_conditional_formats(
        self, workbook: Any, sheet: str
    ) -> list[dict[str, Any]]:  # pragma: no cover
        return []

    def read_data_validations(
        self, workbook: Any, sheet: str
    ) -> list[dict[str, Any]]:  # pragma: no cover
        return []

    def read_hyperlinks(
        self, workbook: Any, sheet: str
    ) -> list[dict[str, Any]]:  # pragma: no cover
        return []

    def read_images(self, workbook: Any, sheet: str) -> list[dict[str, Any]]:  # pragma: no cover
        return []

    def read_pivot_tables(
        self, workbook: Any, sheet: str
    ) -> list[dict[str, Any]]:  # pragma: no cover
        return []

    def read_comments(self, workbook: Any, sheet: str) -> list[dict[str, Any]]:  # pragma: no cover
        return []

    def read_freeze_panes(self, workbook: Any, sheet: str) -> dict[str, Any]:  # pragma: no cover
        return {}

    # Write
    def create_workbook(self) -> Any:  # pragma: no cover
        raise NotImplementedError

    def add_sheet(self, workbook: Any, name: str) -> None:  # pragma: no cover
        raise NotImplementedError

    def write_cell_value(
        self, workbook: Any, sheet: str, cell: str, value: CellValue
    ) -> None:  # pragma: no cover
        raise NotImplementedError

    def write_cell_format(
        self, workbook: Any, sheet: str, cell: str, format: CellFormat
    ) -> None:  # pragma: no cover
        raise NotImplementedError

    def write_cell_border(
        self, workbook: Any, sheet: str, cell: str, border: BorderInfo
    ) -> None:  # pragma: no cover
        raise NotImplementedError

    def set_row_height(
        self, workbook: Any, sheet: str, row: int, height: float
    ) -> None:  # pragma: no cover
        raise NotImplementedError

    def set_column_width(
        self, workbook: Any, sheet: str, column: str, width: float
    ) -> None:  # pragma: no cover
        raise NotImplementedError

    def merge_cells(self, workbook: Any, sheet: str, cell_range: str) -> None:  # pragma: no cover
        raise NotImplementedError

    def add_conditional_format(
        self, workbook: Any, sheet: str, rule: dict[str, Any]
    ) -> None:  # pragma: no cover
        raise NotImplementedError

    def add_data_validation(
        self, workbook: Any, sheet: str, validation: dict[str, Any]
    ) -> None:  # pragma: no cover
        raise NotImplementedError

    def add_hyperlink(
        self, workbook: Any, sheet: str, link: dict[str, Any]
    ) -> None:  # pragma: no cover
        raise NotImplementedError

    def add_image(
        self, workbook: Any, sheet: str, image: dict[str, Any]
    ) -> None:  # pragma: no cover
        raise NotImplementedError

    def add_pivot_table(
        self, workbook: Any, sheet: str, pivot: dict[str, Any]
    ) -> None:  # pragma: no cover
        raise NotImplementedError

    def add_comment(
        self, workbook: Any, sheet: str, comment: dict[str, Any]
    ) -> None:  # pragma: no cover
        raise NotImplementedError

    def set_freeze_panes(
        self, workbook: Any, sheet: str, settings: dict[str, Any]
    ) -> None:  # pragma: no cover
        raise NotImplementedError

    def save_workbook(self, workbook: Any, path: Path) -> None:  # pragma: no cover
        raise NotImplementedError


class TestNamedRangesBase:
    """Base adapter API surface for named ranges."""

    def test_read_named_ranges_default_returns_empty(self) -> None:
        adapter = _StubAdapter()
        assert adapter.read_named_ranges(object(), "S1") == []

    def test_add_named_range_default_is_noop(self) -> None:
        adapter = _StubAdapter()
        assert (
            adapter.add_named_range(object(), "S1", {"name": "X", "refers_to": "S1!$A$1"}) is None
        )


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
