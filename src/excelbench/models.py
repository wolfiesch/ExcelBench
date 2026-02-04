"""Core data models for ExcelBench."""

from dataclasses import dataclass, field
from datetime import datetime
from enum import Enum
from typing import Any


class CellType(str, Enum):
    STRING = "string"
    NUMBER = "number"
    DATE = "date"
    DATETIME = "datetime"
    BOOLEAN = "boolean"
    ERROR = "error"
    BLANK = "blank"
    FORMULA = "formula"


class BorderStyle(str, Enum):
    NONE = "none"
    THIN = "thin"
    MEDIUM = "medium"
    THICK = "thick"
    DOUBLE = "double"
    DASHED = "dashed"
    DOTTED = "dotted"
    HAIR = "hair"
    MEDIUM_DASHED = "mediumDashed"
    DASH_DOT = "dashDot"
    MEDIUM_DASH_DOT = "mediumDashDot"
    DASH_DOT_DOT = "dashDotDot"
    MEDIUM_DASH_DOT_DOT = "mediumDashDotDot"
    SLANT_DASH_DOT = "slantDashDot"


@dataclass
class CellValue:
    """Represents a cell's value and type."""
    type: CellType
    value: Any = None
    formula: str | None = None  # If type is FORMULA, this holds the formula string


@dataclass
class CellFormat:
    """Represents text formatting for a cell."""
    bold: bool | None = None
    italic: bool | None = None
    underline: str | None = None  # "single", "double", etc.
    strikethrough: bool | None = None
    font_name: str | None = None
    font_size: float | None = None
    font_color: str | None = None  # Hex color like "#FF0000"
    bg_color: str | None = None  # Background/fill color


@dataclass
class BorderEdge:
    """Represents one edge of a cell border."""
    style: BorderStyle = BorderStyle.NONE
    color: str = "#000000"


@dataclass
class BorderInfo:
    """Represents all borders of a cell."""
    top: BorderEdge | None = None
    bottom: BorderEdge | None = None
    left: BorderEdge | None = None
    right: BorderEdge | None = None
    diagonal_up: BorderEdge | None = None
    diagonal_down: BorderEdge | None = None


@dataclass
class TestCase:
    """A single test case within a feature test file."""
    id: str
    label: str
    row: int
    expected: dict[str, Any]


@dataclass
class TestFile:
    """Metadata for a generated test file."""
    path: str
    feature: str
    tier: int
    test_cases: list[TestCase] = field(default_factory=list)


@dataclass
class Manifest:
    """Index of all generated test files."""
    generated_at: datetime
    excel_version: str
    generator_version: str
    files: list[TestFile] = field(default_factory=list)


@dataclass
class TestResult:
    """Result of testing a single test case."""
    test_case_id: str
    passed: bool
    expected: dict[str, Any]
    actual: dict[str, Any]
    notes: str | None = None


@dataclass
class FeatureScore:
    """Fidelity score for a feature."""
    feature: str
    library: str
    read_score: int | None = None  # 0-3, None if not applicable
    write_score: int | None = None  # 0-3, None if not applicable
    test_results: list[TestResult] = field(default_factory=list)
    notes: str | None = None


@dataclass
class LibraryInfo:
    """Information about a library being tested."""
    name: str
    version: str
    language: str  # "python" or "rust"
    capabilities: set[str] = field(default_factory=set)  # {"read", "write"}


@dataclass
class BenchmarkMetadata:
    """Metadata about a benchmark run."""
    benchmark_version: str
    run_date: datetime
    excel_version: str
    platform: str


@dataclass
class BenchmarkResults:
    """Complete benchmark results."""
    metadata: BenchmarkMetadata
    libraries: dict[str, LibraryInfo] = field(default_factory=dict)
    scores: list[FeatureScore] = field(default_factory=list)
