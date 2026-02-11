"""Generator for named range test cases (Tier 3)."""

from pathlib import Path

import xlwings as xw

from excelbench.generator.base import FeatureGenerator
from excelbench.models import Importance, TestCase


class NamedRangesGenerator(FeatureGenerator):
    """Generates test cases for workbook and sheet-scoped named ranges."""

    feature_name = "named_ranges"
    tier = 3
    filename = "18_named_ranges.xlsx"

    def __init__(self) -> None:
        self._ops: list[dict[str, object]] = []

    def generate(self, sheet: xw.Sheet) -> list[TestCase]:
        self.setup_header(sheet)

        test_cases: list[TestCase] = []
        row = 2

        # Add a second sheet for cross-sheet references.
        try:
            targets = sheet.book.sheets["Targets"]
        except KeyError:
            targets = sheet.book.sheets.add("Targets")
        targets.range("A1").value = "Target"

        # 1) Workbook-scoped: single cell
        label = "Named range: single cell"
        sheet.range("B2").value = 42
        expected = {
            "name": "SingleCell",
            "scope": "workbook",
            "refers_to": "named_ranges!$B$2",
            "value": 42,
        }
        self.write_test_case(sheet, row, label, expected)
        test_cases.append(TestCase(id="nr_simple_cell", label=label, row=row, expected=expected))
        self._ops.append(
            {"name": "SingleCell", "scope": "workbook", "refers_to": expected["refers_to"]}
        )
        row += 1

        # 2) Workbook-scoped: range
        label = "Named range: cell range"
        sheet.range("B3").value = 1
        sheet.range("C3").value = 2
        sheet.range("D3").value = 3
        expected = {
            "name": "DataRange",
            "scope": "workbook",
            "refers_to": "named_ranges!$B$3:$D$3",
        }
        self.write_test_case(sheet, row, label, expected)
        test_cases.append(TestCase(id="nr_cell_range", label=label, row=row, expected=expected))
        self._ops.append(
            {"name": "DataRange", "scope": "workbook", "refers_to": expected["refers_to"]}
        )
        row += 1

        # 3) Workbook-scoped: used in a formula (name definition + a formula that references it)
        label = "Named range: used in formula"
        sheet.range("B4").value = 0.08
        # Best-effort: reference the name so it is exercised in Excel.
        sheet.range("B6").formula = "=TaxRate*100"
        expected = {
            "name": "TaxRate",
            "scope": "workbook",
            "refers_to": "named_ranges!$B$4",
            "value": 0.08,
        }
        self.write_test_case(sheet, row, label, expected)
        test_cases.append(TestCase(id="nr_formula_ref", label=label, row=row, expected=expected))
        self._ops.append(
            {"name": "TaxRate", "scope": "workbook", "refers_to": expected["refers_to"]}
        )
        row += 1

        # 4) Sheet-scoped name
        label = "Named range: sheet-scoped"
        sheet.range("B5").value = "local"
        expected = {
            "name": "LocalName",
            "scope": "sheet",
            "refers_to": "named_ranges!$B$5",
            "value": "local",
        }
        self.write_test_case(sheet, row, label, expected)
        test_cases.append(
            TestCase(
                id="nr_sheet_scope",
                label=label,
                row=row,
                expected=expected,
                importance=Importance.EDGE,
            )
        )
        self._ops.append(
            {
                "name": "LocalName",
                "scope": "sheet",
                "sheet": self.feature_name,
                "refers_to": expected["refers_to"],
            }
        )
        row += 1

        # 5) Workbook-scoped: cross-sheet reference
        label = "Named range: cross-sheet reference"
        expected = {
            "name": "OtherSheet",
            "scope": "workbook",
            "refers_to": "Targets!$A$1",
        }
        self.write_test_case(sheet, row, label, expected)
        test_cases.append(
            TestCase(
                id="nr_cross_sheet",
                label=label,
                row=row,
                expected=expected,
                importance=Importance.EDGE,
            )
        )
        self._ops.append(
            {"name": "OtherSheet", "scope": "workbook", "refers_to": expected["refers_to"]}
        )
        row += 1

        # 6) Workbook-scoped: underscore name
        label = "Named range: underscore name"
        sheet.range("B7").value = "x"
        expected = {
            "name": "_my_range",
            "scope": "workbook",
            "refers_to": "named_ranges!$B$7",
        }
        self.write_test_case(sheet, row, label, expected)
        test_cases.append(
            TestCase(
                id="nr_special_chars",
                label=label,
                row=row,
                expected=expected,
                importance=Importance.EDGE,
            )
        )
        self._ops.append(
            {"name": "_my_range", "scope": "workbook", "refers_to": expected["refers_to"]}
        )

        return test_cases

    def post_process(self, output_path: Path) -> None:
        # xlwings named range APIs are inconsistent across platforms; define names
        # post-save using openpyxl for reproducibility.
        if not self._ops:
            return

        from openpyxl import load_workbook
        from openpyxl.workbook.defined_name import DefinedName

        wb = load_workbook(output_path)
        try:
            for op in self._ops:
                name = str(op["name"])
                scope = str(op.get("scope", "workbook"))
                refers_to = str(op["refers_to"])

                refers_to_str = str(refers_to)
                attr_text = (
                    f"={refers_to_str}" if not refers_to_str.startswith("=") else refers_to_str
                )

                dn = DefinedName(name, attr_text=attr_text)
                if scope == "sheet":
                    sheet_name = str(op.get("sheet") or self.feature_name)
                    ws = wb[sheet_name]
                    ws.defined_names.add(dn)
                else:
                    wb.defined_names.add(dn)

            wb.save(output_path)
        finally:
            wb.close()
