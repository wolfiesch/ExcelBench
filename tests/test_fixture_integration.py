"""Integration tests: adapters reading canonical Excel-generated fixtures.

These tests verify that each adapter correctly reads the tracked fixture
files in fixtures/excel/, which were produced by real Excel via xlwings.
This catches discrepancies that synthetic (openpyxl-written) test files miss.
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from excelbench.harness.adapters.calamine_adapter import CalamineAdapter
from excelbench.harness.adapters.openpyxl_adapter import OpenpyxlAdapter
from excelbench.harness.adapters.pyexcel_adapter import PyexcelAdapter
from excelbench.harness.adapters.pylightxl_adapter import PylightxlAdapter
from excelbench.models import CellType, CellValue

FIXTURES_DIR = Path(__file__).parent.parent / "fixtures" / "excel"
MANIFEST_PATH = FIXTURES_DIR / "manifest.json"

# Skip all tests if fixtures haven't been generated
pytestmark = pytest.mark.skipif(
    not MANIFEST_PATH.exists(),
    reason="Canonical fixtures not found (run excelbench generate first)",
)

JSONDict = dict[str, Any]


# ─────────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────────


def load_manifest() -> JSONDict:
    with open(MANIFEST_PATH) as f:
        result: JSONDict = json.load(f)
        return result


def fixture_path(relative: str) -> Path:
    return FIXTURES_DIR / relative


def get_test_cases(feature: str) -> list[JSONDict]:
    """Get test cases for a feature from the manifest."""
    manifest = load_manifest()
    for file_entry in manifest["files"]:
        if file_entry["feature"] == feature:
            cases: list[JSONDict] = file_entry["test_cases"]
            return cases
    return []


# ─────────────────────────────────────────────────
# Adapter fixtures
# ─────────────────────────────────────────────────


@pytest.fixture
def openpyxl_adapter() -> OpenpyxlAdapter:
    return OpenpyxlAdapter()


@pytest.fixture
def calamine_adapter() -> CalamineAdapter:
    return CalamineAdapter()


@pytest.fixture
def pylightxl_adapter() -> PylightxlAdapter:
    return PylightxlAdapter()


@pytest.fixture
def pyexcel_adapter() -> PyexcelAdapter:
    return PyexcelAdapter()


# ═════════════════════════════════════════════════
# Tier 1: Cell Values
# ═════════════════════════════════════════════════

CELL_VALUES_PATH = "tier1/01_cell_values.xlsx"


class TestCellValuesOpenpyxl:
    """openpyxl reading Excel-generated cell_values fixture."""

    def test_string_simple(self, openpyxl_adapter: OpenpyxlAdapter) -> None:
        wb = openpyxl_adapter.open_workbook(fixture_path(CELL_VALUES_PATH))
        val = openpyxl_adapter.read_cell_value(wb, "cell_values", "B2")
        assert val.type == CellType.STRING
        assert val.value == "Hello World"
        openpyxl_adapter.close_workbook(wb)

    def test_string_unicode(self, openpyxl_adapter: OpenpyxlAdapter) -> None:
        wb = openpyxl_adapter.open_workbook(fixture_path(CELL_VALUES_PATH))
        val = openpyxl_adapter.read_cell_value(wb, "cell_values", "B3")
        assert val.type == CellType.STRING
        assert "\u65e5\u672c\u8a9e" in val.value  # Japanese chars
        openpyxl_adapter.close_workbook(wb)

    def test_number_integer(self, openpyxl_adapter: OpenpyxlAdapter) -> None:
        wb = openpyxl_adapter.open_workbook(fixture_path(CELL_VALUES_PATH))
        val = openpyxl_adapter.read_cell_value(wb, "cell_values", "B7")
        assert val.type == CellType.NUMBER
        assert val.value == 42
        openpyxl_adapter.close_workbook(wb)

    def test_number_float(self, openpyxl_adapter: OpenpyxlAdapter) -> None:
        wb = openpyxl_adapter.open_workbook(fixture_path(CELL_VALUES_PATH))
        val = openpyxl_adapter.read_cell_value(wb, "cell_values", "B8")
        assert val.type == CellType.NUMBER
        assert abs(val.value - 3.14159265358979) < 0.0001
        openpyxl_adapter.close_workbook(wb)

    def test_boolean_true(self, openpyxl_adapter: OpenpyxlAdapter) -> None:
        wb = openpyxl_adapter.open_workbook(fixture_path(CELL_VALUES_PATH))
        val = openpyxl_adapter.read_cell_value(wb, "cell_values", "B14")
        assert val.type == CellType.BOOLEAN
        assert val.value is True
        openpyxl_adapter.close_workbook(wb)

    def test_blank(self, openpyxl_adapter: OpenpyxlAdapter) -> None:
        wb = openpyxl_adapter.open_workbook(fixture_path(CELL_VALUES_PATH))
        val = openpyxl_adapter.read_cell_value(wb, "cell_values", "B4")
        assert val.type == CellType.BLANK
        openpyxl_adapter.close_workbook(wb)


class TestCellValuesCalamine:
    """python-calamine reading Excel-generated cell_values fixture."""

    def test_string_simple(self, calamine_adapter: CalamineAdapter) -> None:
        wb = calamine_adapter.open_workbook(fixture_path(CELL_VALUES_PATH))
        val = calamine_adapter.read_cell_value(wb, "cell_values", "B2")
        assert val.type == CellType.STRING
        assert val.value == "Hello World"
        calamine_adapter.close_workbook(wb)

    def test_number_integer(self, calamine_adapter: CalamineAdapter) -> None:
        wb = calamine_adapter.open_workbook(fixture_path(CELL_VALUES_PATH))
        val = calamine_adapter.read_cell_value(wb, "cell_values", "B7")
        assert val.type == CellType.NUMBER
        assert val.value == 42 or val.value == 42.0
        calamine_adapter.close_workbook(wb)

    def test_boolean_true(self, calamine_adapter: CalamineAdapter) -> None:
        wb = calamine_adapter.open_workbook(fixture_path(CELL_VALUES_PATH))
        val = calamine_adapter.read_cell_value(wb, "cell_values", "B14")
        assert val.type == CellType.BOOLEAN
        assert val.value is True
        calamine_adapter.close_workbook(wb)


class TestCellValuesPylightxl:
    """pylightxl reading Excel-generated cell_values fixture."""

    def test_string_simple(self, pylightxl_adapter: PylightxlAdapter) -> None:
        wb = pylightxl_adapter.open_workbook(fixture_path(CELL_VALUES_PATH))
        val = pylightxl_adapter.read_cell_value(wb, "cell_values", "B2")
        assert val.type == CellType.STRING
        assert val.value == "Hello World"
        pylightxl_adapter.close_workbook(wb)

    def test_number_integer(self, pylightxl_adapter: PylightxlAdapter) -> None:
        wb = pylightxl_adapter.open_workbook(fixture_path(CELL_VALUES_PATH))
        val = pylightxl_adapter.read_cell_value(wb, "cell_values", "B7")
        assert val.type == CellType.NUMBER
        assert val.value == 42
        pylightxl_adapter.close_workbook(wb)


class TestCellValuesPyexcel:
    """pyexcel reading Excel-generated cell_values fixture."""

    def test_string_simple(self, pyexcel_adapter: PyexcelAdapter) -> None:
        wb = pyexcel_adapter.open_workbook(fixture_path(CELL_VALUES_PATH))
        val = pyexcel_adapter.read_cell_value(wb, "cell_values", "B2")
        assert val.type == CellType.STRING
        assert val.value == "Hello World"
        pyexcel_adapter.close_workbook(wb)

    def test_number_integer(self, pyexcel_adapter: PyexcelAdapter) -> None:
        wb = pyexcel_adapter.open_workbook(fixture_path(CELL_VALUES_PATH))
        val = pyexcel_adapter.read_cell_value(wb, "cell_values", "B7")
        assert val.type == CellType.NUMBER
        assert val.value == 42 or val.value == 42.0
        pyexcel_adapter.close_workbook(wb)


# ═════════════════════════════════════════════════
# Tier 1: Multiple Sheets
# ═════════════════════════════════════════════════

MULTIPLE_SHEETS_PATH = "tier1/09_multiple_sheets.xlsx"


class TestMultipleSheetsOpenpyxl:
    def test_sheet_names(self, openpyxl_adapter: OpenpyxlAdapter) -> None:
        wb = openpyxl_adapter.open_workbook(fixture_path(MULTIPLE_SHEETS_PATH))
        names = openpyxl_adapter.get_sheet_names(wb)
        assert len(names) >= 2
        openpyxl_adapter.close_workbook(wb)

    def test_cross_sheet_read(self, openpyxl_adapter: OpenpyxlAdapter) -> None:
        wb = openpyxl_adapter.open_workbook(fixture_path(MULTIPLE_SHEETS_PATH))
        names = openpyxl_adapter.get_sheet_names(wb)
        # Every sheet should be openable
        for name in names:
            val = openpyxl_adapter.read_cell_value(wb, name, "A1")
            assert val is not None  # No crash
        openpyxl_adapter.close_workbook(wb)


class TestMultipleSheetsCalamine:
    def test_sheet_names(self, calamine_adapter: CalamineAdapter) -> None:
        wb = calamine_adapter.open_workbook(fixture_path(MULTIPLE_SHEETS_PATH))
        names = calamine_adapter.get_sheet_names(wb)
        assert len(names) >= 2
        calamine_adapter.close_workbook(wb)


class TestMultipleSheetsPylightxl:
    def test_sheet_names(self, pylightxl_adapter: PylightxlAdapter) -> None:
        wb = pylightxl_adapter.open_workbook(fixture_path(MULTIPLE_SHEETS_PATH))
        names = pylightxl_adapter.get_sheet_names(wb)
        assert len(names) >= 2
        pylightxl_adapter.close_workbook(wb)


# ═════════════════════════════════════════════════
# Tier 1: Text Formatting (openpyxl only — others don't support)
# ═════════════════════════════════════════════════

TEXT_FORMATTING_PATH = "tier1/03_text_formatting.xlsx"


class TestTextFormattingOpenpyxl:
    def test_bold(self, openpyxl_adapter: OpenpyxlAdapter) -> None:
        wb = openpyxl_adapter.open_workbook(fixture_path(TEXT_FORMATTING_PATH))
        fmt = openpyxl_adapter.read_cell_format(wb, "text_formatting", "B2")
        assert fmt.bold is True
        openpyxl_adapter.close_workbook(wb)

    def test_italic(self, openpyxl_adapter: OpenpyxlAdapter) -> None:
        wb = openpyxl_adapter.open_workbook(fixture_path(TEXT_FORMATTING_PATH))
        fmt = openpyxl_adapter.read_cell_format(wb, "text_formatting", "B3")
        assert fmt.italic is True
        openpyxl_adapter.close_workbook(wb)

    def test_font_color(self, openpyxl_adapter: OpenpyxlAdapter) -> None:
        wb = openpyxl_adapter.open_workbook(fixture_path(TEXT_FORMATTING_PATH))
        fmt = openpyxl_adapter.read_cell_format(wb, "text_formatting", "B15")
        assert fmt.font_color is not None
        assert fmt.font_color == "#FF0000"
        openpyxl_adapter.close_workbook(wb)


# ═════════════════════════════════════════════════
# Tier 1: Background Colors (openpyxl only)
# ═════════════════════════════════════════════════

BG_COLORS_PATH = "tier1/04_background_colors.xlsx"


class TestBackgroundColorsOpenpyxl:
    def test_has_bg_color(self, openpyxl_adapter: OpenpyxlAdapter) -> None:
        wb = openpyxl_adapter.open_workbook(fixture_path(BG_COLORS_PATH))
        fmt = openpyxl_adapter.read_cell_format(wb, "background_colors", "B2")
        assert fmt.bg_color is not None
        assert fmt.bg_color.startswith("#")
        openpyxl_adapter.close_workbook(wb)


# ═════════════════════════════════════════════════
# Tier 1: Borders (openpyxl only)
# ═════════════════════════════════════════════════

BORDERS_PATH = "tier1/07_borders.xlsx"


class TestBordersOpenpyxl:
    def test_thin_border(self, openpyxl_adapter: OpenpyxlAdapter) -> None:
        wb = openpyxl_adapter.open_workbook(fixture_path(BORDERS_PATH))
        border = openpyxl_adapter.read_cell_border(wb, "borders", "B2")
        assert border.top is not None
        assert border.top.style.value == "thin"
        openpyxl_adapter.close_workbook(wb)


# ═════════════════════════════════════════════════
# Tier 1: Dimensions (openpyxl only)
# ═════════════════════════════════════════════════

DIMENSIONS_PATH = "tier1/08_dimensions.xlsx"


class TestDimensionsOpenpyxl:
    def test_row_height(self, openpyxl_adapter: OpenpyxlAdapter) -> None:
        wb = openpyxl_adapter.open_workbook(fixture_path(DIMENSIONS_PATH))
        height = openpyxl_adapter.read_row_height(wb, "dimensions", 2)
        assert height is not None
        assert height > 0
        openpyxl_adapter.close_workbook(wb)

    def test_column_width(self, openpyxl_adapter: OpenpyxlAdapter) -> None:
        wb = openpyxl_adapter.open_workbook(fixture_path(DIMENSIONS_PATH))
        width = openpyxl_adapter.read_column_width(wb, "dimensions", "B")
        # Width should be set (non-default)
        assert width is not None
        openpyxl_adapter.close_workbook(wb)


# ═════════════════════════════════════════════════
# Tier 1: Formulas (openpyxl only)
# ═════════════════════════════════════════════════

FORMULAS_PATH = "tier1/02_formulas.xlsx"


class TestFormulasOpenpyxl:
    def test_formula_preserved(self, openpyxl_adapter: OpenpyxlAdapter) -> None:
        wb = openpyxl_adapter.open_workbook(fixture_path(FORMULAS_PATH))
        val = openpyxl_adapter.read_cell_value(wb, "formulas", "B2")
        assert val.type == CellType.FORMULA
        assert val.formula is not None
        assert "SUM" in val.formula.upper()
        openpyxl_adapter.close_workbook(wb)


# ═════════════════════════════════════════════════
# Tier 2: Merged Cells
# ═════════════════════════════════════════════════

MERGED_CELLS_PATH = "tier2/10_merged_cells.xlsx"


class TestMergedCellsOpenpyxl:
    def test_has_merged_ranges(self, openpyxl_adapter: OpenpyxlAdapter) -> None:
        wb = openpyxl_adapter.open_workbook(fixture_path(MERGED_CELLS_PATH))
        ranges = openpyxl_adapter.read_merged_ranges(wb, "merged_cells")
        assert len(ranges) > 0
        openpyxl_adapter.close_workbook(wb)

    def test_merged_value(self, openpyxl_adapter: OpenpyxlAdapter) -> None:
        wb = openpyxl_adapter.open_workbook(fixture_path(MERGED_CELLS_PATH))
        val = openpyxl_adapter.read_cell_value(wb, "merged_cells", "B2")
        assert val.type != CellType.BLANK  # Top-left of merge should have value
        openpyxl_adapter.close_workbook(wb)


# ═════════════════════════════════════════════════
# Tier 2: Conditional Formatting
# ═════════════════════════════════════════════════

COND_FORMAT_PATH = "tier2/11_conditional_formatting.xlsx"


class TestConditionalFormattingOpenpyxl:
    def test_has_rules(self, openpyxl_adapter: OpenpyxlAdapter) -> None:
        wb = openpyxl_adapter.open_workbook(fixture_path(COND_FORMAT_PATH))
        rules = openpyxl_adapter.read_conditional_formats(wb, "conditional_formatting")
        assert len(rules) > 0
        openpyxl_adapter.close_workbook(wb)

    def test_rule_has_range(self, openpyxl_adapter: OpenpyxlAdapter) -> None:
        wb = openpyxl_adapter.open_workbook(fixture_path(COND_FORMAT_PATH))
        rules = openpyxl_adapter.read_conditional_formats(wb, "conditional_formatting")
        for rule in rules:
            assert "range" in rule
            assert rule["range"] is not None
        openpyxl_adapter.close_workbook(wb)

    def test_rule_has_type(self, openpyxl_adapter: OpenpyxlAdapter) -> None:
        wb = openpyxl_adapter.open_workbook(fixture_path(COND_FORMAT_PATH))
        rules = openpyxl_adapter.read_conditional_formats(wb, "conditional_formatting")
        for rule in rules:
            assert "rule_type" in rule
        openpyxl_adapter.close_workbook(wb)


# ═════════════════════════════════════════════════
# Tier 2: Data Validation
# ═════════════════════════════════════════════════

DATA_VALIDATION_PATH = "tier2/12_data_validation.xlsx"


class TestDataValidationOpenpyxl:
    def test_has_validations(self, openpyxl_adapter: OpenpyxlAdapter) -> None:
        wb = openpyxl_adapter.open_workbook(fixture_path(DATA_VALIDATION_PATH))
        validations = openpyxl_adapter.read_data_validations(wb, "data_validation")
        assert len(validations) > 0
        openpyxl_adapter.close_workbook(wb)

    def test_validation_has_type(self, openpyxl_adapter: OpenpyxlAdapter) -> None:
        wb = openpyxl_adapter.open_workbook(fixture_path(DATA_VALIDATION_PATH))
        validations = openpyxl_adapter.read_data_validations(wb, "data_validation")
        for v in validations:
            assert "validation_type" in v
            assert v["validation_type"] is not None
        openpyxl_adapter.close_workbook(wb)


# ═════════════════════════════════════════════════
# Tier 2: Hyperlinks
# ═════════════════════════════════════════════════

HYPERLINKS_PATH = "tier2/13_hyperlinks.xlsx"


class TestHyperlinksOpenpyxl:
    def test_has_hyperlinks(self, openpyxl_adapter: OpenpyxlAdapter) -> None:
        wb = openpyxl_adapter.open_workbook(fixture_path(HYPERLINKS_PATH))
        links = openpyxl_adapter.read_hyperlinks(wb, "hyperlinks")
        assert len(links) > 0
        openpyxl_adapter.close_workbook(wb)

    def test_hyperlink_has_target(self, openpyxl_adapter: OpenpyxlAdapter) -> None:
        wb = openpyxl_adapter.open_workbook(fixture_path(HYPERLINKS_PATH))
        links = openpyxl_adapter.read_hyperlinks(wb, "hyperlinks")
        for link in links:
            assert "target" in link
            assert link["target"] is not None
        openpyxl_adapter.close_workbook(wb)

    def test_hyperlink_has_cell(self, openpyxl_adapter: OpenpyxlAdapter) -> None:
        wb = openpyxl_adapter.open_workbook(fixture_path(HYPERLINKS_PATH))
        links = openpyxl_adapter.read_hyperlinks(wb, "hyperlinks")
        for link in links:
            assert "cell" in link
        openpyxl_adapter.close_workbook(wb)


# ═════════════════════════════════════════════════
# Tier 2: Images
# ═════════════════════════════════════════════════

IMAGES_PATH = "tier2/14_images.xlsx"


class TestImagesOpenpyxl:
    def test_has_images(self, openpyxl_adapter: OpenpyxlAdapter) -> None:
        wb = openpyxl_adapter.open_workbook(fixture_path(IMAGES_PATH))
        images = openpyxl_adapter.read_images(wb, "images")
        assert len(images) > 0
        openpyxl_adapter.close_workbook(wb)

    def test_image_has_cell(self, openpyxl_adapter: OpenpyxlAdapter) -> None:
        wb = openpyxl_adapter.open_workbook(fixture_path(IMAGES_PATH))
        images = openpyxl_adapter.read_images(wb, "images")
        for img in images:
            assert "cell" in img
            assert img["cell"] is not None
        openpyxl_adapter.close_workbook(wb)


# ═════════════════════════════════════════════════
# Tier 2: Comments
# ═════════════════════════════════════════════════

COMMENTS_PATH = "tier2/16_comments.xlsx"


class TestCommentsOpenpyxl:
    def test_has_comments(self, openpyxl_adapter: OpenpyxlAdapter) -> None:
        wb = openpyxl_adapter.open_workbook(fixture_path(COMMENTS_PATH))
        comments = openpyxl_adapter.read_comments(wb, "comments")
        assert len(comments) > 0
        openpyxl_adapter.close_workbook(wb)

    def test_comment_has_text(self, openpyxl_adapter: OpenpyxlAdapter) -> None:
        wb = openpyxl_adapter.open_workbook(fixture_path(COMMENTS_PATH))
        comments = openpyxl_adapter.read_comments(wb, "comments")
        for comment in comments:
            assert "text" in comment
            assert comment["text"]  # non-empty
        openpyxl_adapter.close_workbook(wb)

    def test_comment_has_cell(self, openpyxl_adapter: OpenpyxlAdapter) -> None:
        wb = openpyxl_adapter.open_workbook(fixture_path(COMMENTS_PATH))
        comments = openpyxl_adapter.read_comments(wb, "comments")
        for comment in comments:
            assert "cell" in comment
        openpyxl_adapter.close_workbook(wb)


# ═════════════════════════════════════════════════
# Tier 2: Freeze Panes
# ═════════════════════════════════════════════════

FREEZE_PANES_PATH = "tier2/17_freeze_panes.xlsx"


class TestFreezePanesOpenpyxl:
    def test_has_freeze(self, openpyxl_adapter: OpenpyxlAdapter) -> None:
        wb = openpyxl_adapter.open_workbook(fixture_path(FREEZE_PANES_PATH))
        names = openpyxl_adapter.get_sheet_names(wb)
        # At least one sheet should have freeze/split settings
        found_freeze = False
        for name in names:
            result = openpyxl_adapter.read_freeze_panes(wb, name)
            if result.get("mode"):
                found_freeze = True
                break
        assert found_freeze, "No freeze/split panes found in any sheet"
        openpyxl_adapter.close_workbook(wb)


# ═════════════════════════════════════════════════
# Tier 2: Calamine/pylightxl return empty for Tier 2
# ═════════════════════════════════════════════════


class TestCalamineTier2Empty:
    """Calamine doesn't support Tier 2 features — verify graceful empty returns."""

    def test_merged_cells_empty(self, calamine_adapter: CalamineAdapter) -> None:
        wb = calamine_adapter.open_workbook(fixture_path(MERGED_CELLS_PATH))
        assert calamine_adapter.read_merged_ranges(wb, "merged_cells") == []
        calamine_adapter.close_workbook(wb)

    def test_conditional_formatting_empty(self, calamine_adapter: CalamineAdapter) -> None:
        wb = calamine_adapter.open_workbook(fixture_path(COND_FORMAT_PATH))
        assert calamine_adapter.read_conditional_formats(wb, "conditional_formatting") == []
        calamine_adapter.close_workbook(wb)

    def test_comments_empty(self, calamine_adapter: CalamineAdapter) -> None:
        wb = calamine_adapter.open_workbook(fixture_path(COMMENTS_PATH))
        assert calamine_adapter.read_comments(wb, "comments") == []
        calamine_adapter.close_workbook(wb)

    def test_hyperlinks_empty(self, calamine_adapter: CalamineAdapter) -> None:
        wb = calamine_adapter.open_workbook(fixture_path(HYPERLINKS_PATH))
        assert calamine_adapter.read_hyperlinks(wb, "hyperlinks") == []
        calamine_adapter.close_workbook(wb)

    def test_freeze_panes_empty(self, calamine_adapter: CalamineAdapter) -> None:
        wb = calamine_adapter.open_workbook(fixture_path(FREEZE_PANES_PATH))
        names = calamine_adapter.get_sheet_names(wb)
        for name in names:
            assert calamine_adapter.read_freeze_panes(wb, name) == {}
        calamine_adapter.close_workbook(wb)


# ═════════════════════════════════════════════════
# Write → Read Roundtrip (openpyxl)
# ═════════════════════════════════════════════════


class TestOpenpyxlWriteRoundtrip:
    """Test openpyxl write → read roundtrip for Tier 2 features."""

    def test_merged_cells_roundtrip(
        self, openpyxl_adapter: OpenpyxlAdapter, tmp_path: Path
    ) -> None:
        path = tmp_path / "merged.xlsx"
        wb = openpyxl_adapter.create_workbook()
        openpyxl_adapter.add_sheet(wb, "S1")
        openpyxl_adapter.write_cell_value(
            wb, "S1", "A1", CellValue(type=CellType.STRING, value="merged")
        )
        openpyxl_adapter.merge_cells(wb, "S1", "A1:C1")
        openpyxl_adapter.save_workbook(wb, path)

        wb2 = openpyxl_adapter.open_workbook(path)
        ranges = openpyxl_adapter.read_merged_ranges(wb2, "S1")
        assert any("A1" in r and "C1" in r for r in ranges)
        openpyxl_adapter.close_workbook(wb2)

    def test_comment_roundtrip(
        self, openpyxl_adapter: OpenpyxlAdapter, tmp_path: Path
    ) -> None:
        path = tmp_path / "comments.xlsx"
        wb = openpyxl_adapter.create_workbook()
        openpyxl_adapter.add_sheet(wb, "S1")
        openpyxl_adapter.add_comment(
            wb, "S1", {"cell": "A1", "text": "Test note", "author": "Bot"}
        )
        openpyxl_adapter.save_workbook(wb, path)

        wb2 = openpyxl_adapter.open_workbook(path)
        comments = openpyxl_adapter.read_comments(wb2, "S1")
        assert len(comments) == 1
        assert comments[0]["text"] == "Test note"
        assert comments[0]["author"] == "Bot"
        openpyxl_adapter.close_workbook(wb2)

    def test_hyperlink_roundtrip(
        self, openpyxl_adapter: OpenpyxlAdapter, tmp_path: Path
    ) -> None:
        path = tmp_path / "links.xlsx"
        wb = openpyxl_adapter.create_workbook()
        openpyxl_adapter.add_sheet(wb, "S1")
        openpyxl_adapter.add_hyperlink(
            wb,
            "S1",
            {
                "cell": "A1",
                "target": "https://example.com",
                "display": "Example",
            },
        )
        openpyxl_adapter.save_workbook(wb, path)

        wb2 = openpyxl_adapter.open_workbook(path)
        links = openpyxl_adapter.read_hyperlinks(wb2, "S1")
        assert len(links) >= 1
        assert links[0]["target"] == "https://example.com"
        openpyxl_adapter.close_workbook(wb2)

    def test_data_validation_roundtrip(
        self, openpyxl_adapter: OpenpyxlAdapter, tmp_path: Path
    ) -> None:
        path = tmp_path / "dv.xlsx"
        wb = openpyxl_adapter.create_workbook()
        openpyxl_adapter.add_sheet(wb, "S1")
        openpyxl_adapter.add_data_validation(
            wb,
            "S1",
            {
                "range": "A1:A10",
                "validation_type": "list",
                "formula1": '"Yes,No,Maybe"',
            },
        )
        openpyxl_adapter.save_workbook(wb, path)

        wb2 = openpyxl_adapter.open_workbook(path)
        validations = openpyxl_adapter.read_data_validations(wb2, "S1")
        assert len(validations) >= 1
        assert validations[0]["validation_type"] == "list"
        openpyxl_adapter.close_workbook(wb2)

    def test_freeze_panes_roundtrip(
        self, openpyxl_adapter: OpenpyxlAdapter, tmp_path: Path
    ) -> None:
        path = tmp_path / "freeze.xlsx"
        wb = openpyxl_adapter.create_workbook()
        openpyxl_adapter.add_sheet(wb, "S1")
        openpyxl_adapter.set_freeze_panes(
            wb, "S1", {"mode": "freeze", "top_left_cell": "B2"}
        )
        openpyxl_adapter.save_workbook(wb, path)

        wb2 = openpyxl_adapter.open_workbook(path)
        result = openpyxl_adapter.read_freeze_panes(wb2, "S1")
        assert result.get("mode") == "freeze"
        assert result.get("top_left_cell") == "B2"
        openpyxl_adapter.close_workbook(wb2)

    def test_conditional_format_roundtrip(
        self, openpyxl_adapter: OpenpyxlAdapter, tmp_path: Path
    ) -> None:
        path = tmp_path / "cf.xlsx"
        wb = openpyxl_adapter.create_workbook()
        openpyxl_adapter.add_sheet(wb, "S1")
        openpyxl_adapter.add_conditional_format(
            wb,
            "S1",
            {
                "range": "A1:A10",
                "rule_type": "cellIs",
                "operator": "greaterThan",
                "formula": "100",
                "format": {"bg_color": "#FF0000"},
            },
        )
        openpyxl_adapter.save_workbook(wb, path)

        wb2 = openpyxl_adapter.open_workbook(path)
        rules = openpyxl_adapter.read_conditional_formats(wb2, "S1")
        assert len(rules) >= 1
        assert rules[0]["rule_type"] == "cellIs"
        openpyxl_adapter.close_workbook(wb2)
