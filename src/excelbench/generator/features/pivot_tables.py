"""Generator for pivot table test cases (Tier 2)."""

import shutil
import sys
from pathlib import Path

import xlwings as xw

from excelbench.generator.base import FeatureGenerator
from excelbench.models import TestCase


class PivotTablesGenerator(FeatureGenerator):
    """Generates test cases for pivot tables."""

    feature_name = "pivot_tables"
    tier = 2
    filename = "15_pivot_tables.xlsx"

    def __init__(self) -> None:
        self._fixture_path = Path("fixtures/excel/tier2/15_pivot_tables.xlsx")

    def generate(self, sheet: xw.Sheet) -> list[TestCase]:
        self.setup_header(sheet)

        if sys.platform == "darwin":
            if self._fixture_path.exists():
                return self._fixture_test_cases(sheet)
            print("  Pivot fixture not found; skipping pivot tests on macOS.")
            return []

        wb = sheet.book
        data_sheet = wb.sheets.add("Data")
        pivot_sheet = wb.sheets.add("Pivot")

        # Seed data
        data = [
            ["Region", "Product", "Date", "Sales"],
            ["North", "A", "2026-01-05", 100],
            ["North", "B", "2026-01-08", 150],
            ["South", "A", "2026-02-03", 90],
            ["South", "B", "2026-02-10", 120],
            ["West", "A", "2026-03-15", 200],
        ]
        data_sheet.range("A1").value = data

        source_range = data_sheet.range("A1:D6").api
        dest = pivot_sheet.range("B3").api

        # Create pivot cache and table
        pivot_cache = wb.api.PivotCaches().Create(SourceType=1, SourceData=source_range)
        pivot_table = pivot_cache.CreatePivotTable(TableDestination=dest, TableName="SalesPivot")

        # Field layout
        pivot_table.PivotFields("Region").Orientation = 1  # xlRowField
        pivot_table.PivotFields("Product").Orientation = 2  # xlColumnField
        pivot_table.PivotFields("Date").Orientation = 3  # xlPageField

        # Data field (sum) to ensure the pivot is materialized.
        pivot_table.AddDataField(pivot_table.PivotFields("Sales"), "Sum of Sales", -4157)

        return self._minimal_test_cases(sheet)

    def post_process(self, output_path: Path) -> None:
        if sys.platform != "darwin":
            return
        if self._fixture_path.exists():
            shutil.copyfile(self._fixture_path, output_path)

    def _fixture_test_cases(self, sheet: xw.Sheet) -> list[TestCase]:
        return self._minimal_test_cases(sheet)

    def _minimal_test_cases(self, sheet: xw.Sheet) -> list[TestCase]:
        test_cases: list[TestCase] = []
        row = 2

        expected = {
            "pivot": {
                "name": "SalesPivot",
                "source_range": "Data!A1:D6",
                "target_cell": "Pivot!B3",
            }
        }
        self.write_test_case(sheet, row, "Pivot: basic layout", expected)
        test_cases.append(
            TestCase(
                id="pivot_basic",
                label="Pivot: basic layout",
                row=row,
                expected=expected,
                sheet="Pivot",
            )
        )

        return test_cases
