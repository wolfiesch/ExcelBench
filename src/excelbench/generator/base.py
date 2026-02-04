"""Base classes for test file generators."""

from abc import ABC, abstractmethod
from pathlib import Path
from typing import Any

import xlwings as xw

from excelbench.models import TestCase


class FeatureGenerator(ABC):
    """Abstract base class for feature test file generators.

    Each feature generator creates a single Excel file containing
    test cases for that feature. The file follows a 3-column format:
    - Column A: Label describing the test case
    - Column B: The test cell with the feature applied
    - Column C: Expected values in JSON format
    """

    feature_name: str
    tier: int
    filename: str

    @abstractmethod
    def generate(self, sheet: xw.Sheet) -> list[TestCase]:
        """Generate test cases in the given worksheet.

        Args:
            sheet: The xlwings Sheet to write test cases to.

        Returns:
            List of TestCase objects describing what was generated.
        """
        ...

    def setup_header(self, sheet: xw.Sheet) -> None:
        """Set up the header row with column labels."""
        sheet.range("A1").value = "Label"
        sheet.range("B1").value = "Test Cell"
        sheet.range("C1").value = "Expected"

        # Format header row
        header_range = sheet.range("A1:C1")
        header_range.font.bold = True
        header_range.color = (220, 220, 220)  # Light gray background

        # Set column widths
        sheet.range("A:A").column_width = 30
        sheet.range("B:B").column_width = 25
        sheet.range("C:C").column_width = 50

    def write_test_case(
        self,
        sheet: xw.Sheet,
        row: int,
        label: str,
        expected: dict[str, Any],
    ) -> None:
        """Write the label and expected columns for a test case.

        The test cell (column B) should be written separately with
        the actual feature being tested.

        Args:
            sheet: The worksheet to write to.
            row: The row number (1-indexed).
            label: Description of the test case.
            expected: Dictionary of expected values.
        """
        import json

        sheet.range(f"A{row}").value = label
        sheet.range(f"C{row}").value = json.dumps(expected)

    def create_workbook(self, output_dir: Path) -> tuple[xw.Book, Path]:
        """Create a new workbook for this feature.

        Args:
            output_dir: Directory to save the file in.

        Returns:
            Tuple of (workbook, output_path).
        """
        output_path = output_dir / f"tier{self.tier}" / self.filename
        output_path.parent.mkdir(parents=True, exist_ok=True)

        # Create new workbook
        wb = xw.Book()

        # Rename first sheet to feature name
        wb.sheets[0].name = self.feature_name

        return wb, output_path

    def save_and_close(self, wb: xw.Book, output_path: Path) -> None:
        """Save the workbook and close it.

        Args:
            wb: The workbook to save.
            output_path: Path to save to.
        """
        wb.save(str(output_path))
        wb.close()
