"""Base adapter protocol for Excel libraries."""

from abc import ABC, abstractmethod
from pathlib import Path
from typing import Any

from excelbench.models import (
    CellValue,
    CellFormat,
    BorderInfo,
    LibraryInfo,
)


class ExcelAdapter(ABC):
    """Abstract base class for Excel library adapters.

    Each adapter wraps a specific Excel library and provides a
    unified interface for reading and writing cell data.
    """

    @property
    @abstractmethod
    def info(self) -> LibraryInfo:
        """Get information about this library."""
        ...

    @property
    def name(self) -> str:
        """Library name."""
        return self.info.name

    @property
    def capabilities(self) -> set[str]:
        """Library capabilities (read, write)."""
        return self.info.capabilities

    def can_read(self) -> bool:
        """Check if this adapter supports reading."""
        return "read" in self.capabilities

    def can_write(self) -> bool:
        """Check if this adapter supports writing."""
        return "write" in self.capabilities

    # =========================================================================
    # Read Operations
    # =========================================================================

    @abstractmethod
    def open_workbook(self, path: Path) -> Any:
        """Open a workbook for reading.

        Args:
            path: Path to the Excel file.

        Returns:
            Library-specific workbook object.
        """
        ...

    @abstractmethod
    def close_workbook(self, workbook: Any) -> None:
        """Close an opened workbook.

        Args:
            workbook: The workbook object to close.
        """
        ...

    @abstractmethod
    def get_sheet_names(self, workbook: Any) -> list[str]:
        """Get list of sheet names in a workbook.

        Args:
            workbook: The workbook object.

        Returns:
            List of sheet names.
        """
        ...

    @abstractmethod
    def read_cell_value(
        self,
        workbook: Any,
        sheet: str,
        cell: str,
    ) -> CellValue:
        """Read the value of a cell.

        Args:
            workbook: The workbook object.
            sheet: Sheet name.
            cell: Cell reference (e.g., "A1", "B2").

        Returns:
            CellValue with type and value.
        """
        ...

    @abstractmethod
    def read_cell_format(
        self,
        workbook: Any,
        sheet: str,
        cell: str,
    ) -> CellFormat:
        """Read the formatting of a cell.

        Args:
            workbook: The workbook object.
            sheet: Sheet name.
            cell: Cell reference.

        Returns:
            CellFormat with formatting properties.
        """
        ...

    @abstractmethod
    def read_cell_border(
        self,
        workbook: Any,
        sheet: str,
        cell: str,
    ) -> BorderInfo:
        """Read the border information of a cell.

        Args:
            workbook: The workbook object.
            sheet: Sheet name.
            cell: Cell reference.

        Returns:
            BorderInfo with border properties.
        """
        ...

    # =========================================================================
    # Write Operations
    # =========================================================================

    @abstractmethod
    def create_workbook(self) -> Any:
        """Create a new workbook.

        Returns:
            Library-specific workbook object.
        """
        ...

    @abstractmethod
    def add_sheet(self, workbook: Any, name: str) -> None:
        """Add a new sheet to a workbook.

        Args:
            workbook: The workbook object.
            name: Name for the new sheet.
        """
        ...

    @abstractmethod
    def write_cell_value(
        self,
        workbook: Any,
        sheet: str,
        cell: str,
        value: CellValue,
    ) -> None:
        """Write a value to a cell.

        Args:
            workbook: The workbook object.
            sheet: Sheet name.
            cell: Cell reference.
            value: The value to write.
        """
        ...

    @abstractmethod
    def write_cell_format(
        self,
        workbook: Any,
        sheet: str,
        cell: str,
        format: CellFormat,
    ) -> None:
        """Apply formatting to a cell.

        Args:
            workbook: The workbook object.
            sheet: Sheet name.
            cell: Cell reference.
            format: The formatting to apply.
        """
        ...

    @abstractmethod
    def write_cell_border(
        self,
        workbook: Any,
        sheet: str,
        cell: str,
        border: BorderInfo,
    ) -> None:
        """Apply border to a cell.

        Args:
            workbook: The workbook object.
            sheet: Sheet name.
            cell: Cell reference.
            border: The border to apply.
        """
        ...

    @abstractmethod
    def save_workbook(self, workbook: Any, path: Path) -> None:
        """Save a workbook to a file.

        Args:
            workbook: The workbook object.
            path: Path to save to.
        """
        ...


class ReadOnlyAdapter(ExcelAdapter):
    """Base class for read-only adapters.

    Provides default implementations that raise NotImplementedError
    for all write operations.
    """

    def create_workbook(self) -> Any:
        raise NotImplementedError(f"{self.name} is read-only")

    def add_sheet(self, workbook: Any, name: str) -> None:
        raise NotImplementedError(f"{self.name} is read-only")

    def write_cell_value(
        self,
        workbook: Any,
        sheet: str,
        cell: str,
        value: CellValue,
    ) -> None:
        raise NotImplementedError(f"{self.name} is read-only")

    def write_cell_format(
        self,
        workbook: Any,
        sheet: str,
        cell: str,
        format: CellFormat,
    ) -> None:
        raise NotImplementedError(f"{self.name} is read-only")

    def write_cell_border(
        self,
        workbook: Any,
        sheet: str,
        cell: str,
        border: BorderInfo,
    ) -> None:
        raise NotImplementedError(f"{self.name} is read-only")

    def save_workbook(self, workbook: Any, path: Path) -> None:
        raise NotImplementedError(f"{self.name} is read-only")


class WriteOnlyAdapter(ExcelAdapter):
    """Base class for write-only adapters.

    Provides default implementations that raise NotImplementedError
    for all read operations.
    """

    def open_workbook(self, path: Path) -> Any:
        raise NotImplementedError(f"{self.name} is write-only")

    def close_workbook(self, workbook: Any) -> None:
        pass  # Nothing to close for write-only

    def get_sheet_names(self, workbook: Any) -> list[str]:
        raise NotImplementedError(f"{self.name} is write-only")

    def read_cell_value(
        self,
        workbook: Any,
        sheet: str,
        cell: str,
    ) -> CellValue:
        raise NotImplementedError(f"{self.name} is write-only")

    def read_cell_format(
        self,
        workbook: Any,
        sheet: str,
        cell: str,
    ) -> CellFormat:
        raise NotImplementedError(f"{self.name} is write-only")

    def read_cell_border(
        self,
        workbook: Any,
        sheet: str,
        cell: str,
    ) -> BorderInfo:
        raise NotImplementedError(f"{self.name} is write-only")
