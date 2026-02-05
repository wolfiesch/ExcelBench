"""Generator for background color test cases."""

import xlwings as xw

from excelbench.generator.base import FeatureGenerator
from excelbench.models import TestCase


class BackgroundColorsGenerator(FeatureGenerator):
    """Generates test cases for cell background colors."""

    feature_name = "background_colors"
    tier = 1
    filename = "04_background_colors.xlsx"

    def generate(self, sheet: xw.Sheet) -> list[TestCase]:
        self.setup_header(sheet)

        test_cases: list[TestCase] = []
        row = 2

        test_cases.append(
            self._test_color(
                sheet,
                row,
                "Background - red",
                (255, 0, 0),
                "#FF0000",
                "bg_red",
            )
        )
        row += 1
        test_cases.append(
            self._test_color(
                sheet,
                row,
                "Background - blue",
                (0, 0, 255),
                "#0000FF",
                "bg_blue",
            )
        )
        row += 1
        test_cases.append(
            self._test_color(
                sheet,
                row,
                "Background - green",
                (0, 255, 0),
                "#00FF00",
                "bg_green",
            )
        )
        row += 1
        test_cases.append(
            self._test_color(
                sheet,
                row,
                "Background - custom (#8B4513)",
                (139, 69, 19),
                "#8B4513",
                "bg_custom",
            )
        )
        row += 1

        return test_cases

    def _test_color(
        self,
        sheet: xw.Sheet,
        row: int,
        label: str,
        rgb: tuple[int, int, int],
        expected_color: str,
        case_id: str,
    ) -> TestCase:
        expected = {"bg_color": expected_color}
        self.write_test_case(sheet, row, label, expected)
        cell = sheet.range(f"B{row}")
        cell.value = label
        cell.color = rgb
        return TestCase(id=case_id, label=label, row=row, expected=expected)
