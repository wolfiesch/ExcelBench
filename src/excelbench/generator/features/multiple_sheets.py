"""Generator for multiple sheets test cases."""

import xlwings as xw

from excelbench.generator.base import FeatureGenerator
from excelbench.models import TestCase


class MultipleSheetsGenerator(FeatureGenerator):
    """Generates test cases for multiple sheet handling."""

    feature_name = "multiple_sheets"
    tier = 1
    filename = "09_multiple_sheets.xlsx"

    def generate(self, sheet: xw.Sheet) -> list[TestCase]:
        wb = sheet.book

        # Rename the first sheet and add two more
        sheet.name = "Alpha"
        sheet_alpha = sheet
        sheet_beta = wb.sheets.add("Beta", after=sheet_alpha)
        sheet_gamma = wb.sheets.add("Gamma", after=sheet_beta)

        for ws in [sheet_alpha, sheet_beta, sheet_gamma]:
            self.setup_header(ws)

        test_cases: list[TestCase] = []

        # Sheet names test (stored on Alpha)
        expected_names = ["Alpha", "Beta", "Gamma"]
        expected = {"sheet_names": expected_names}
        self.write_test_case(sheet_alpha, 2, "Sheet names", expected)
        test_cases.append(TestCase(
            id="sheet_names",
            label="Sheet names",
            row=2,
            expected=expected,
            sheet="Alpha",
        ))

        # Value on each sheet
        test_cases.append(self._sheet_value_case(sheet_alpha, "Alpha", 3))
        test_cases.append(self._sheet_value_case(sheet_beta, "Beta", 3))
        test_cases.append(self._sheet_value_case(sheet_gamma, "Gamma", 3))

        return test_cases

    def _sheet_value_case(self, sheet: xw.Sheet, name: str, row: int) -> TestCase:
        label = f"{name} value"
        expected = {"type": "string", "value": name}
        self.write_test_case(sheet, row, label, expected)
        sheet.range(f"B{row}").value = name
        return TestCase(
            id=f"value_{name.lower()}",
            label=label,
            row=row,
            expected=expected,
            sheet=name,
            cell=f"B{row}",
        )
