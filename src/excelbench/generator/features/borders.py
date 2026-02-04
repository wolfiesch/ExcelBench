"""Generator for border test cases."""

import xlwings as xw

from excelbench.generator.base import FeatureGenerator
from excelbench.models import TestCase


# Excel border style constants (from xlwings/Excel API)
class XlBorderWeight:
    HAIRLINE = 1
    THIN = 2
    MEDIUM = -4138
    THICK = 4


class XlLineStyle:
    CONTINUOUS = 1
    DASH = -4115
    DASH_DOT = 4
    DASH_DOT_DOT = 5
    DOT = -4118
    DOUBLE = -4119
    NONE = -4142
    SLANT_DASH_DOT = 13


class XlBordersIndex:
    EDGE_LEFT = 7
    EDGE_TOP = 8
    EDGE_BOTTOM = 9
    EDGE_RIGHT = 10
    INSIDE_VERTICAL = 11
    INSIDE_HORIZONTAL = 12
    DIAGONAL_DOWN = 5
    DIAGONAL_UP = 6


class BordersGenerator(FeatureGenerator):
    """Generates test cases for cell borders.

    Tests: styles, weights, colors, positions, diagonals.
    """

    feature_name = "borders"
    tier = 1
    filename = "07_borders.xlsx"

    def generate(self, sheet: xw.Sheet) -> list[TestCase]:
        """Generate border test cases."""
        self.setup_header(sheet)

        test_cases = []
        row = 2

        # Border styles
        test_cases.append(self._test_thin_all(sheet, row))
        row += 1

        test_cases.append(self._test_medium_all(sheet, row))
        row += 1

        test_cases.append(self._test_thick_all(sheet, row))
        row += 1

        test_cases.append(self._test_double(sheet, row))
        row += 1

        test_cases.append(self._test_dashed(sheet, row))
        row += 1

        test_cases.append(self._test_dotted(sheet, row))
        row += 1

        test_cases.append(self._test_dash_dot(sheet, row))
        row += 1

        test_cases.append(self._test_dash_dot_dot(sheet, row))
        row += 1

        # Individual edges
        test_cases.append(self._test_top_only(sheet, row))
        row += 1

        test_cases.append(self._test_bottom_only(sheet, row))
        row += 1

        test_cases.append(self._test_left_only(sheet, row))
        row += 1

        test_cases.append(self._test_right_only(sheet, row))
        row += 1

        # Diagonal borders
        test_cases.append(self._test_diagonal_up(sheet, row))
        row += 1

        test_cases.append(self._test_diagonal_down(sheet, row))
        row += 1

        test_cases.append(self._test_diagonal_both(sheet, row))
        row += 1

        # Colors
        test_cases.append(self._test_color_red(sheet, row))
        row += 1

        test_cases.append(self._test_color_blue(sheet, row))
        row += 1

        test_cases.append(self._test_color_custom(sheet, row))
        row += 1

        # Mixed edges
        test_cases.append(self._test_mixed_styles(sheet, row))
        row += 1

        test_cases.append(self._test_mixed_colors(sheet, row))
        row += 1

        return test_cases

    def _set_all_borders(
        self,
        cell: xw.Range,
        weight: int = XlBorderWeight.THIN,
        line_style: int = XlLineStyle.CONTINUOUS,
        color: tuple[int, int, int] | None = None,
    ) -> None:
        """Set all four edges of a cell border."""
        for edge in [
            XlBordersIndex.EDGE_TOP,
            XlBordersIndex.EDGE_BOTTOM,
            XlBordersIndex.EDGE_LEFT,
            XlBordersIndex.EDGE_RIGHT,
        ]:
            border = cell.api.Borders(edge)
            border.LineStyle = line_style
            border.Weight = weight
            if color:
                border.Color = self._rgb_to_int(color)

    def _rgb_to_int(self, rgb: tuple[int, int, int]) -> int:
        """Convert RGB tuple to Excel color integer."""
        return rgb[0] + (rgb[1] * 256) + (rgb[2] * 256 * 256)

    def _test_thin_all(self, sheet: xw.Sheet, row: int) -> TestCase:
        label = "Border - thin all edges"
        expected = {"border_style": "thin", "border_color": "#000000"}

        self.write_test_case(sheet, row, label, expected)
        cell = sheet.range(f"B{row}")
        cell.value = "Thin"
        self._set_all_borders(cell, XlBorderWeight.THIN)

        return TestCase(id="thin_all", label=label, row=row, expected=expected)

    def _test_medium_all(self, sheet: xw.Sheet, row: int) -> TestCase:
        label = "Border - medium all edges"
        expected = {"border_style": "medium", "border_color": "#000000"}

        self.write_test_case(sheet, row, label, expected)
        cell = sheet.range(f"B{row}")
        cell.value = "Medium"
        self._set_all_borders(cell, XlBorderWeight.MEDIUM)

        return TestCase(id="medium_all", label=label, row=row, expected=expected)

    def _test_thick_all(self, sheet: xw.Sheet, row: int) -> TestCase:
        label = "Border - thick all edges"
        expected = {"border_style": "thick", "border_color": "#000000"}

        self.write_test_case(sheet, row, label, expected)
        cell = sheet.range(f"B{row}")
        cell.value = "Thick"
        self._set_all_borders(cell, XlBorderWeight.THICK)

        return TestCase(id="thick_all", label=label, row=row, expected=expected)

    def _test_double(self, sheet: xw.Sheet, row: int) -> TestCase:
        label = "Border - double line"
        expected = {"border_style": "double", "border_color": "#000000"}

        self.write_test_case(sheet, row, label, expected)
        cell = sheet.range(f"B{row}")
        cell.value = "Double"
        self._set_all_borders(cell, line_style=XlLineStyle.DOUBLE)

        return TestCase(id="double", label=label, row=row, expected=expected)

    def _test_dashed(self, sheet: xw.Sheet, row: int) -> TestCase:
        label = "Border - dashed"
        expected = {"border_style": "dashed", "border_color": "#000000"}

        self.write_test_case(sheet, row, label, expected)
        cell = sheet.range(f"B{row}")
        cell.value = "Dashed"
        self._set_all_borders(cell, line_style=XlLineStyle.DASH)

        return TestCase(id="dashed", label=label, row=row, expected=expected)

    def _test_dotted(self, sheet: xw.Sheet, row: int) -> TestCase:
        label = "Border - dotted"
        expected = {"border_style": "dotted", "border_color": "#000000"}

        self.write_test_case(sheet, row, label, expected)
        cell = sheet.range(f"B{row}")
        cell.value = "Dotted"
        self._set_all_borders(cell, line_style=XlLineStyle.DOT)

        return TestCase(id="dotted", label=label, row=row, expected=expected)

    def _test_dash_dot(self, sheet: xw.Sheet, row: int) -> TestCase:
        label = "Border - dash-dot"
        expected = {"border_style": "dashDot", "border_color": "#000000"}

        self.write_test_case(sheet, row, label, expected)
        cell = sheet.range(f"B{row}")
        cell.value = "Dash-Dot"
        self._set_all_borders(cell, line_style=XlLineStyle.DASH_DOT)

        return TestCase(id="dash_dot", label=label, row=row, expected=expected)

    def _test_dash_dot_dot(self, sheet: xw.Sheet, row: int) -> TestCase:
        label = "Border - dash-dot-dot"
        expected = {"border_style": "dashDotDot", "border_color": "#000000"}

        self.write_test_case(sheet, row, label, expected)
        cell = sheet.range(f"B{row}")
        cell.value = "Dash-Dot-Dot"
        self._set_all_borders(cell, line_style=XlLineStyle.DASH_DOT_DOT)

        return TestCase(id="dash_dot_dot", label=label, row=row, expected=expected)

    def _test_top_only(self, sheet: xw.Sheet, row: int) -> TestCase:
        label = "Border - top only"
        expected = {"border_top": "thin", "border_bottom": None, "border_left": None, "border_right": None}

        self.write_test_case(sheet, row, label, expected)
        cell = sheet.range(f"B{row}")
        cell.value = "Top Only"
        border = cell.api.Borders(XlBordersIndex.EDGE_TOP)
        border.LineStyle = XlLineStyle.CONTINUOUS
        border.Weight = XlBorderWeight.THIN

        return TestCase(id="top_only", label=label, row=row, expected=expected)

    def _test_bottom_only(self, sheet: xw.Sheet, row: int) -> TestCase:
        label = "Border - bottom only"
        expected = {"border_top": None, "border_bottom": "thin", "border_left": None, "border_right": None}

        self.write_test_case(sheet, row, label, expected)
        cell = sheet.range(f"B{row}")
        cell.value = "Bottom Only"
        border = cell.api.Borders(XlBordersIndex.EDGE_BOTTOM)
        border.LineStyle = XlLineStyle.CONTINUOUS
        border.Weight = XlBorderWeight.THIN

        return TestCase(id="bottom_only", label=label, row=row, expected=expected)

    def _test_left_only(self, sheet: xw.Sheet, row: int) -> TestCase:
        label = "Border - left only"
        expected = {"border_top": None, "border_bottom": None, "border_left": "thin", "border_right": None}

        self.write_test_case(sheet, row, label, expected)
        cell = sheet.range(f"B{row}")
        cell.value = "Left Only"
        border = cell.api.Borders(XlBordersIndex.EDGE_LEFT)
        border.LineStyle = XlLineStyle.CONTINUOUS
        border.Weight = XlBorderWeight.THIN

        return TestCase(id="left_only", label=label, row=row, expected=expected)

    def _test_right_only(self, sheet: xw.Sheet, row: int) -> TestCase:
        label = "Border - right only"
        expected = {"border_top": None, "border_bottom": None, "border_left": None, "border_right": "thin"}

        self.write_test_case(sheet, row, label, expected)
        cell = sheet.range(f"B{row}")
        cell.value = "Right Only"
        border = cell.api.Borders(XlBordersIndex.EDGE_RIGHT)
        border.LineStyle = XlLineStyle.CONTINUOUS
        border.Weight = XlBorderWeight.THIN

        return TestCase(id="right_only", label=label, row=row, expected=expected)

    def _test_diagonal_up(self, sheet: xw.Sheet, row: int) -> TestCase:
        label = "Border - diagonal up"
        expected = {"border_diagonal_up": "thin"}

        self.write_test_case(sheet, row, label, expected)
        cell = sheet.range(f"B{row}")
        cell.value = "Diag Up"
        border = cell.api.Borders(XlBordersIndex.DIAGONAL_UP)
        border.LineStyle = XlLineStyle.CONTINUOUS
        border.Weight = XlBorderWeight.THIN

        return TestCase(id="diagonal_up", label=label, row=row, expected=expected)

    def _test_diagonal_down(self, sheet: xw.Sheet, row: int) -> TestCase:
        label = "Border - diagonal down"
        expected = {"border_diagonal_down": "thin"}

        self.write_test_case(sheet, row, label, expected)
        cell = sheet.range(f"B{row}")
        cell.value = "Diag Down"
        border = cell.api.Borders(XlBordersIndex.DIAGONAL_DOWN)
        border.LineStyle = XlLineStyle.CONTINUOUS
        border.Weight = XlBorderWeight.THIN

        return TestCase(id="diagonal_down", label=label, row=row, expected=expected)

    def _test_diagonal_both(self, sheet: xw.Sheet, row: int) -> TestCase:
        label = "Border - diagonal both"
        expected = {"border_diagonal_up": "thin", "border_diagonal_down": "thin"}

        self.write_test_case(sheet, row, label, expected)
        cell = sheet.range(f"B{row}")
        cell.value = "Both Diag"
        for diag in [XlBordersIndex.DIAGONAL_UP, XlBordersIndex.DIAGONAL_DOWN]:
            border = cell.api.Borders(diag)
            border.LineStyle = XlLineStyle.CONTINUOUS
            border.Weight = XlBorderWeight.THIN

        return TestCase(id="diagonal_both", label=label, row=row, expected=expected)

    def _test_color_red(self, sheet: xw.Sheet, row: int) -> TestCase:
        label = "Border - red color"
        expected = {"border_style": "thin", "border_color": "#FF0000"}

        self.write_test_case(sheet, row, label, expected)
        cell = sheet.range(f"B{row}")
        cell.value = "Red Border"
        self._set_all_borders(cell, XlBorderWeight.THIN, color=(255, 0, 0))

        return TestCase(id="color_red", label=label, row=row, expected=expected)

    def _test_color_blue(self, sheet: xw.Sheet, row: int) -> TestCase:
        label = "Border - blue color"
        expected = {"border_style": "thin", "border_color": "#0000FF"}

        self.write_test_case(sheet, row, label, expected)
        cell = sheet.range(f"B{row}")
        cell.value = "Blue Border"
        self._set_all_borders(cell, XlBorderWeight.THIN, color=(0, 0, 255))

        return TestCase(id="color_blue", label=label, row=row, expected=expected)

    def _test_color_custom(self, sheet: xw.Sheet, row: int) -> TestCase:
        label = "Border - custom color (#8B4513)"
        expected = {"border_style": "thin", "border_color": "#8B4513"}

        self.write_test_case(sheet, row, label, expected)
        cell = sheet.range(f"B{row}")
        cell.value = "Custom Color"
        self._set_all_borders(cell, XlBorderWeight.THIN, color=(139, 69, 19))

        return TestCase(id="color_custom", label=label, row=row, expected=expected)

    def _test_mixed_styles(self, sheet: xw.Sheet, row: int) -> TestCase:
        label = "Border - mixed styles per edge"
        expected = {
            "border_top": "thick",
            "border_bottom": "thin",
            "border_left": "medium",
            "border_right": "dashed",
        }

        self.write_test_case(sheet, row, label, expected)
        cell = sheet.range(f"B{row}")
        cell.value = "Mixed Styles"

        # Top - thick
        border = cell.api.Borders(XlBordersIndex.EDGE_TOP)
        border.LineStyle = XlLineStyle.CONTINUOUS
        border.Weight = XlBorderWeight.THICK

        # Bottom - thin
        border = cell.api.Borders(XlBordersIndex.EDGE_BOTTOM)
        border.LineStyle = XlLineStyle.CONTINUOUS
        border.Weight = XlBorderWeight.THIN

        # Left - medium
        border = cell.api.Borders(XlBordersIndex.EDGE_LEFT)
        border.LineStyle = XlLineStyle.CONTINUOUS
        border.Weight = XlBorderWeight.MEDIUM

        # Right - dashed
        border = cell.api.Borders(XlBordersIndex.EDGE_RIGHT)
        border.LineStyle = XlLineStyle.DASH

        return TestCase(id="mixed_styles", label=label, row=row, expected=expected)

    def _test_mixed_colors(self, sheet: xw.Sheet, row: int) -> TestCase:
        label = "Border - mixed colors per edge"
        expected = {
            "border_top_color": "#FF0000",
            "border_bottom_color": "#00FF00",
            "border_left_color": "#0000FF",
            "border_right_color": "#FFFF00",
        }

        self.write_test_case(sheet, row, label, expected)
        cell = sheet.range(f"B{row}")
        cell.value = "Mixed Colors"

        colors = [
            (XlBordersIndex.EDGE_TOP, (255, 0, 0)),      # Red
            (XlBordersIndex.EDGE_BOTTOM, (0, 255, 0)),   # Green
            (XlBordersIndex.EDGE_LEFT, (0, 0, 255)),     # Blue
            (XlBordersIndex.EDGE_RIGHT, (255, 255, 0)),  # Yellow
        ]

        for edge, color in colors:
            border = cell.api.Borders(edge)
            border.LineStyle = XlLineStyle.CONTINUOUS
            border.Weight = XlBorderWeight.THIN
            border.Color = self._rgb_to_int(color)

        return TestCase(id="mixed_colors", label=label, row=row, expected=expected)
