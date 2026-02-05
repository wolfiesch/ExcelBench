"""Generator for number format test cases."""

from datetime import date

import xlwings as xw

from excelbench.generator.base import FeatureGenerator
from excelbench.models import TestCase


class NumberFormatsGenerator(FeatureGenerator):
    """Generates test cases for number formats."""

    feature_name = "number_formats"
    tier = 1
    filename = "05_number_formats.xlsx"

    def generate(self, sheet: xw.Sheet) -> list[TestCase]:
        self.setup_header(sheet)

        test_cases: list[TestCase] = []
        row = 2

        test_cases.append(
            self._test_format(
                sheet,
                row,
                "Format - currency",
                1234.56,
                "$#,##0.00",
                "currency",
            )
        )
        row += 1
        test_cases.append(
            self._test_format(
                sheet,
                row,
                "Format - percent",
                0.256,
                "0.00%",
                "percent",
            )
        )
        row += 1
        test_cases.append(
            self._test_format(
                sheet,
                row,
                "Format - date",
                date(2026, 2, 4),
                "yyyy-mm-dd",
                "date",
            )
        )
        row += 1
        test_cases.append(
            self._test_format(
                sheet,
                row,
                "Format - scientific",
                12345.678,
                "0.00E+00",
                "scientific",
            )
        )
        row += 1
        test_cases.append(
            self._test_format(
                sheet,
                row,
                "Format - custom text",
                12.3,
                "\"USD\" 0.00",
                "custom_text",
            )
        )
        row += 1

        return test_cases

    def _test_format(
        self,
        sheet: xw.Sheet,
        row: int,
        label: str,
        value: object,
        number_format: str,
        case_id: str,
    ) -> TestCase:
        expected = {"number_format": number_format}
        self.write_test_case(sheet, row, label, expected)
        cell = sheet.range(f"B{row}")
        cell.value = value
        cell.number_format = number_format
        return TestCase(id=f"numfmt_{case_id}", label=label, row=row, expected=expected)
