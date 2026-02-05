"""Generator for cell values test cases."""

from datetime import date, datetime

import xlwings as xw

from excelbench.generator.base import FeatureGenerator
from excelbench.models import TestCase


class CellValuesGenerator(FeatureGenerator):
    """Generates test cases for cell value types.

    Tests: strings, numbers, dates, booleans, errors, blanks.
    """

    feature_name = "cell_values"
    tier = 1
    filename = "01_cell_values.xlsx"

    def generate(self, sheet: xw.Sheet) -> list[TestCase]:
        """Generate cell value test cases."""
        self.setup_header(sheet)

        test_cases = []
        row = 2  # Start after header

        # String tests
        test_cases.append(self._test_string_simple(sheet, row))
        row += 1

        test_cases.append(self._test_string_unicode(sheet, row))
        row += 1

        test_cases.append(self._test_string_empty(sheet, row))
        row += 1

        test_cases.append(self._test_string_long(sheet, row))
        row += 1

        test_cases.append(self._test_string_newline(sheet, row))
        row += 1

        # Number tests
        test_cases.append(self._test_number_integer(sheet, row))
        row += 1

        test_cases.append(self._test_number_float(sheet, row))
        row += 1

        test_cases.append(self._test_number_negative(sheet, row))
        row += 1

        test_cases.append(self._test_number_large(sheet, row))
        row += 1

        test_cases.append(self._test_number_scientific(sheet, row))
        row += 1

        # Date tests
        test_cases.append(self._test_date_standard(sheet, row))
        row += 1

        test_cases.append(self._test_datetime(sheet, row))
        row += 1

        # Boolean tests
        test_cases.append(self._test_boolean_true(sheet, row))
        row += 1

        test_cases.append(self._test_boolean_false(sheet, row))
        row += 1

        # Error tests
        test_cases.append(self._test_error_div0(sheet, row))
        row += 1

        test_cases.append(self._test_error_na(sheet, row))
        row += 1

        test_cases.append(self._test_error_value(sheet, row))
        row += 1

        # Blank test
        test_cases.append(self._test_blank(sheet, row))
        row += 1

        return test_cases

    def _test_string_simple(self, sheet: xw.Sheet, row: int) -> TestCase:
        label = "String - simple"
        value = "Hello World"
        expected = {"type": "string", "value": value}

        self.write_test_case(sheet, row, label, expected)
        sheet.range(f"B{row}").value = value

        return TestCase(id="string_simple", label=label, row=row, expected=expected)

    def _test_string_unicode(self, sheet: xw.Sheet, row: int) -> TestCase:
        label = "String - unicode"
        value = "日本語🎉émojis"
        expected = {"type": "string", "value": value}

        self.write_test_case(sheet, row, label, expected)
        sheet.range(f"B{row}").value = value

        return TestCase(id="string_unicode", label=label, row=row, expected=expected)

    def _test_string_empty(self, sheet: xw.Sheet, row: int) -> TestCase:
        label = "String - empty"
        value = ""
        expected = {"type": "blank"}

        self.write_test_case(sheet, row, label, expected)
        # Write empty string explicitly (not None/blank)
        sheet.range(f"B{row}").value = value

        return TestCase(id="string_empty", label=label, row=row, expected=expected)

    def _test_string_long(self, sheet: xw.Sheet, row: int) -> TestCase:
        label = "String - long (1000 chars)"
        value = "A" * 1000
        expected = {"type": "string", "value": value}

        self.write_test_case(sheet, row, label, expected)
        sheet.range(f"B{row}").value = value

        return TestCase(id="string_long", label=label, row=row, expected=expected)

    def _test_string_newline(self, sheet: xw.Sheet, row: int) -> TestCase:
        label = "String - with newlines"
        value = "Line 1\nLine 2\nLine 3"
        expected = {"type": "string", "value": value}

        self.write_test_case(sheet, row, label, expected)
        sheet.range(f"B{row}").value = value

        return TestCase(id="string_newline", label=label, row=row, expected=expected)

    def _test_number_integer(self, sheet: xw.Sheet, row: int) -> TestCase:
        label = "Number - integer"
        value = 42
        expected = {"type": "number", "value": value}

        self.write_test_case(sheet, row, label, expected)
        sheet.range(f"B{row}").value = value

        return TestCase(id="number_integer", label=label, row=row, expected=expected)

    def _test_number_float(self, sheet: xw.Sheet, row: int) -> TestCase:
        label = "Number - float"
        value = 3.14159265358979
        expected = {"type": "number", "value": value}

        self.write_test_case(sheet, row, label, expected)
        sheet.range(f"B{row}").value = value

        return TestCase(id="number_float", label=label, row=row, expected=expected)

    def _test_number_negative(self, sheet: xw.Sheet, row: int) -> TestCase:
        label = "Number - negative"
        value = -100.5
        expected = {"type": "number", "value": value}

        self.write_test_case(sheet, row, label, expected)
        sheet.range(f"B{row}").value = value

        return TestCase(id="number_negative", label=label, row=row, expected=expected)

    def _test_number_large(self, sheet: xw.Sheet, row: int) -> TestCase:
        label = "Number - large"
        value = 1234567890123456
        expected = {"type": "number", "value": value}

        self.write_test_case(sheet, row, label, expected)
        sheet.range(f"B{row}").value = value

        return TestCase(id="number_large", label=label, row=row, expected=expected)

    def _test_number_scientific(self, sheet: xw.Sheet, row: int) -> TestCase:
        label = "Number - scientific notation"
        value = 1.23e-10
        expected = {"type": "number", "value": value}

        self.write_test_case(sheet, row, label, expected)
        sheet.range(f"B{row}").value = value

        return TestCase(id="number_scientific", label=label, row=row, expected=expected)

    def _test_date_standard(self, sheet: xw.Sheet, row: int) -> TestCase:
        label = "Date - standard"
        value = date(2026, 2, 4)
        expected = {"type": "date", "value": "2026-02-04"}

        self.write_test_case(sheet, row, label, expected)
        sheet.range(f"B{row}").value = value
        # Format as date
        sheet.range(f"B{row}").number_format = "yyyy-mm-dd"

        return TestCase(id="date_standard", label=label, row=row, expected=expected)

    def _test_datetime(self, sheet: xw.Sheet, row: int) -> TestCase:
        label = "DateTime - with time"
        value = datetime(2026, 2, 4, 10, 30, 45)
        expected = {"type": "datetime", "value": "2026-02-04T10:30:45"}

        self.write_test_case(sheet, row, label, expected)
        sheet.range(f"B{row}").value = value
        # Format as datetime
        sheet.range(f"B{row}").number_format = "yyyy-mm-dd hh:mm:ss"

        return TestCase(id="datetime", label=label, row=row, expected=expected)

    def _test_boolean_true(self, sheet: xw.Sheet, row: int) -> TestCase:
        label = "Boolean - TRUE"
        value = True
        expected = {"type": "boolean", "value": True}

        self.write_test_case(sheet, row, label, expected)
        sheet.range(f"B{row}").value = value

        return TestCase(id="boolean_true", label=label, row=row, expected=expected)

    def _test_boolean_false(self, sheet: xw.Sheet, row: int) -> TestCase:
        label = "Boolean - FALSE"
        value = False
        expected = {"type": "boolean", "value": False}

        self.write_test_case(sheet, row, label, expected)
        sheet.range(f"B{row}").value = value

        return TestCase(id="boolean_false", label=label, row=row, expected=expected)

    def _test_error_div0(self, sheet: xw.Sheet, row: int) -> TestCase:
        label = "Error - #DIV/0!"
        expected = {"type": "error", "value": "#DIV/0!"}

        self.write_test_case(sheet, row, label, expected)
        # Use a formula that produces the error
        sheet.range(f"B{row}").formula = "=1/0"

        return TestCase(id="error_div0", label=label, row=row, expected=expected)

    def _test_error_na(self, sheet: xw.Sheet, row: int) -> TestCase:
        label = "Error - #N/A"
        expected = {"type": "error", "value": "#N/A"}

        self.write_test_case(sheet, row, label, expected)
        sheet.range(f"B{row}").formula = "=NA()"

        return TestCase(id="error_na", label=label, row=row, expected=expected)

    def _test_error_value(self, sheet: xw.Sheet, row: int) -> TestCase:
        label = "Error - #VALUE!"
        expected = {"type": "error", "value": "#VALUE!"}

        self.write_test_case(sheet, row, label, expected)
        # Formula that produces #VALUE! error
        sheet.range(f"B{row}").formula = '="text"+1'

        return TestCase(id="error_value", label=label, row=row, expected=expected)

    def _test_blank(self, sheet: xw.Sheet, row: int) -> TestCase:
        label = "Blank cell"
        expected = {"type": "blank"}

        self.write_test_case(sheet, row, label, expected)
        # Leave B column empty (don't write anything)

        return TestCase(id="blank", label=label, row=row, expected=expected)
