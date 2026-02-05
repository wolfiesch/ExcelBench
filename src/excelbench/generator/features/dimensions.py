"""Generator for row height and column width test cases."""

import xlwings as xw

from excelbench.generator.base import FeatureGenerator
from excelbench.models import TestCase


class DimensionsGenerator(FeatureGenerator):
    """Generates test cases for row heights and column widths."""

    feature_name = "dimensions"
    tier = 1
    filename = "08_dimensions.xlsx"

    def generate(self, sheet: xw.Sheet) -> list[TestCase]:
        self.setup_header(sheet)

        test_cases: list[TestCase] = []
        row = 2

        # Row height tests (use column B for visibility)
        test_cases.append(self._test_row_height(sheet, row, 30, "row_height_30"))
        row += 1
        test_cases.append(self._test_row_height(sheet, row, 45, "row_height_45"))
        row += 1

        # Column width tests (use columns D/E to avoid collisions)
        test_cases.append(self._test_column_width(sheet, row, "D", 20, "col_width_20"))
        row += 1
        test_cases.append(self._test_column_width(sheet, row, "E", 8, "col_width_8"))
        row += 1

        return test_cases

    def _test_row_height(self, sheet: xw.Sheet, row: int, height: float, case_id: str) -> TestCase:
        label = f"Row height - {height}"
        expected = {"row_height": height}
        self.write_test_case(sheet, row, label, expected)
        cell = sheet.range(f"B{row}")
        cell.value = label
        sheet.range(f"{row}:{row}").row_height = height
        return TestCase(id=case_id, label=label, row=row, expected=expected, cell=f"B{row}")

    def _test_column_width(
        self,
        sheet: xw.Sheet,
        row: int,
        column: str,
        width: float,
        case_id: str,
    ) -> TestCase:
        label = f"Column width - {column} = {width}"
        expected = {"column_width": width}
        self.write_test_case(sheet, row, label, expected)
        cell = sheet.range(f"{column}{row}")
        cell.value = label
        sheet.range(f"{column}:{column}").column_width = width
        return TestCase(id=case_id, label=label, row=row, expected=expected, cell=f"{column}{row}")
