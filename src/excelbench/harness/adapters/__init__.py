"""Excel library adapters."""

from excelbench.harness.adapters.base import ExcelAdapter, ReadOnlyAdapter, WriteOnlyAdapter
from excelbench.harness.adapters.openpyxl_adapter import OpenpyxlAdapter
from excelbench.harness.adapters.xlsxwriter_adapter import XlsxwriterAdapter

__all__ = [
    "ExcelAdapter",
    "ReadOnlyAdapter",
    "WriteOnlyAdapter",
    "OpenpyxlAdapter",
    "XlsxwriterAdapter",
]


def get_all_adapters() -> list[ExcelAdapter]:
    """Get all available adapters."""
    return [
        OpenpyxlAdapter(),
        XlsxwriterAdapter(),
    ]
