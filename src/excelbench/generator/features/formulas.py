"""Generator for formula test cases."""

import xlwings as xw

from excelbench.generator.base import FeatureGenerator
from excelbench.models import TestCase


class FormulasGenerator(FeatureGenerator):
    """Generates test cases for formulas (formula text fidelity)."""

    feature_name = "formulas"
    tier = 1
    filename = "02_formulas.xlsx"

    def generate(self, sheet: xw.Sheet) -> list[TestCase]:
        self.setup_header(sheet)

        test_cases: list[TestCase] = []
        row = 2

        # Simple SUM formula
        test_cases.append(self._test_formula_sum(sheet, row))
        row += 1

        # Cell reference
        test_cases.append(self._test_formula_cell_ref(sheet, row))
        row += 1

        # String concatenation
        test_cases.append(self._test_formula_concat(sheet, row))
        row += 1

        # Cross-sheet reference
        test_cases.append(self._test_formula_cross_sheet(sheet, row))
        row += 1

        return test_cases

    def _test_formula_sum(self, sheet: xw.Sheet, row: int) -> TestCase:
        label = "Formula - SUM"
        formula = "=SUM(1,2,3)"
        expected = {"type": "formula", "formula": formula}

        self.write_test_case(sheet, row, label, expected)
        sheet.range(f"B{row}").formula = formula

        return TestCase(id="formula_sum", label=label, row=row, expected=expected)

    def _test_formula_cell_ref(self, sheet: xw.Sheet, row: int) -> TestCase:
        label = "Formula - cell reference"
        sheet.range(f"A{row}").value = 10
        formula = f"=A{row}*2"
        expected = {"type": "formula", "formula": formula}

        self.write_test_case(sheet, row, label, expected)
        sheet.range(f"B{row}").formula = formula

        return TestCase(id="formula_cell_ref", label=label, row=row, expected=expected)

    def _test_formula_concat(self, sheet: xw.Sheet, row: int) -> TestCase:
        label = "Formula - concat"
        sheet.range(f"A{row}").value = "Hello"
        sheet.range(f"A{row+1}").value = "World"
        formula = f'=A{row}&" "&A{row+1}'
        expected = {"type": "formula", "formula": formula}

        self.write_test_case(sheet, row, label, expected)
        sheet.range(f"B{row}").formula = formula

        return TestCase(id="formula_concat", label=label, row=row, expected=expected)

    def _test_formula_cross_sheet(self, sheet: xw.Sheet, row: int) -> TestCase:
        label = "Formula - cross sheet"
        wb = sheet.book
        ref_sheet = wb.sheets.add("References", after=sheet)
        self.setup_header(ref_sheet)
        ref_sheet.range("B2").value = 42
        ref_sheet.range("A2").value = "Reference Value"
        ref_sheet.range("C2").value = '{"type": "number", "value": 42}'

        formula = "='References'!B2"
        expected = {"type": "formula", "formula": formula}

        self.write_test_case(sheet, row, label, expected)
        sheet.range(f"B{row}").formula = formula

        return TestCase(id="formula_cross_sheet", label=label, row=row, expected=expected)
