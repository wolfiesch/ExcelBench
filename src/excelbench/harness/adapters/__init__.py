"""Excel library adapters."""

from excelbench.harness.adapters.base import ExcelAdapter, ReadOnlyAdapter, WriteOnlyAdapter
from excelbench.harness.adapters.openpyxl_adapter import OpenpyxlAdapter
from excelbench.harness.adapters.xlsxwriter_adapter import XlsxwriterAdapter
from excelbench.harness.adapters.xlwings_oracle_adapter import ExcelOracleAdapter

__all__ = [
    "ExcelAdapter",
    "ReadOnlyAdapter",
    "WriteOnlyAdapter",
    "OpenpyxlAdapter",
    "XlsxwriterAdapter",
    "ExcelOracleAdapter",
]


def get_all_adapters() -> list[ExcelAdapter]:
    """Get all available adapters."""
    return [
        OpenpyxlAdapter(),
        XlsxwriterAdapter(),
    ]
