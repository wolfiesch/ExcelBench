"""Test runner for executing benchmarks."""

import json
from datetime import datetime, timezone
from pathlib import Path
from typing import Any
import platform

from excelbench.generator.generate import load_manifest
from excelbench.harness.adapters import ExcelAdapter, get_all_adapters
from excelbench.models import (
    BenchmarkMetadata,
    BenchmarkResults,
    FeatureScore,
    LibraryInfo,
    TestResult,
    Manifest,
    TestFile,
    TestCase,
)

BENCHMARK_VERSION = "0.1.0"


def run_benchmark(
    test_dir: Path,
    adapters: list[ExcelAdapter] | None = None,
) -> BenchmarkResults:
    """Run the full benchmark suite.

    Args:
        test_dir: Directory containing test files and manifest.json.
        adapters: List of adapters to test. If None, uses all available.

    Returns:
        BenchmarkResults with all scores.
    """
    test_dir = Path(test_dir)

    # Load manifest
    manifest_path = test_dir / "manifest.json"
    if not manifest_path.exists():
        raise FileNotFoundError(f"Manifest not found: {manifest_path}")

    manifest = load_manifest(manifest_path)

    # Get adapters
    if adapters is None:
        adapters = get_all_adapters()

    # Create metadata
    metadata = BenchmarkMetadata(
        benchmark_version=BENCHMARK_VERSION,
        run_date=datetime.now(timezone.utc),
        excel_version=manifest.excel_version,
        platform=f"{platform.system()}-{platform.machine()}",
    )

    # Collect library info
    libraries = {adapter.name: adapter.info for adapter in adapters}

    # Run tests for each file
    all_scores: list[FeatureScore] = []

    for test_file in manifest.files:
        file_path = test_dir / test_file.path

        if not file_path.exists():
            print(f"Warning: Test file not found: {file_path}")
            continue

        print(f"Testing {test_file.feature}...")

        for adapter in adapters:
            score = test_feature(
                adapter=adapter,
                test_file=test_file,
                file_path=file_path,
            )
            all_scores.append(score)
            print(f"  {adapter.name}: read={score.read_score}, write={score.write_score}")

    return BenchmarkResults(
        metadata=metadata,
        libraries=libraries,
        scores=all_scores,
    )


def test_feature(
    adapter: ExcelAdapter,
    test_file: TestFile,
    file_path: Path,
) -> FeatureScore:
    """Test a single feature with a single adapter.

    Args:
        adapter: The adapter to test.
        test_file: TestFile metadata.
        file_path: Path to the test file.

    Returns:
        FeatureScore with results.
    """
    read_results: list[TestResult] = []
    write_results: list[TestResult] = []

    # Test reading
    if adapter.can_read():
        read_results = test_read(adapter, test_file, file_path)

    # Test writing
    if adapter.can_write():
        write_results = test_write(adapter, test_file, file_path)

    # Calculate scores
    read_score = calculate_score(read_results) if read_results else None
    write_score = calculate_score(write_results) if write_results else None

    return FeatureScore(
        feature=test_file.feature,
        library=adapter.name,
        read_score=read_score,
        write_score=write_score,
        test_results=read_results + write_results,
    )


def test_read(
    adapter: ExcelAdapter,
    test_file: TestFile,
    file_path: Path,
) -> list[TestResult]:
    """Test reading a feature.

    Args:
        adapter: The adapter to test.
        test_file: TestFile metadata.
        file_path: Path to the test file.

    Returns:
        List of TestResult for each test case.
    """
    results: list[TestResult] = []

    try:
        workbook = adapter.open_workbook(file_path)
    except Exception as e:
        # Can't even open the file
        for tc in test_file.test_cases:
            results.append(TestResult(
                test_case_id=tc.id,
                passed=False,
                expected=tc.expected,
                actual={"error": str(e)},
                notes="Failed to open workbook",
            ))
        return results

    try:
        sheet_names = adapter.get_sheet_names(workbook)
        if not sheet_names:
            raise ValueError("No sheets found in workbook")

        sheet_name = sheet_names[0]  # Use first sheet

        for tc in test_file.test_cases:
            result = test_read_case(adapter, workbook, sheet_name, tc, test_file.feature)
            results.append(result)

    finally:
        adapter.close_workbook(workbook)

    return results


def test_read_case(
    adapter: ExcelAdapter,
    workbook: Any,
    sheet: str,
    test_case: TestCase,
    feature: str,
) -> TestResult:
    """Test reading a single test case.

    Args:
        adapter: The adapter to test.
        workbook: The opened workbook.
        sheet: Sheet name.
        test_case: The test case to verify.
        feature: The feature being tested.

    Returns:
        TestResult for this test case.
    """
    cell = f"B{test_case.row}"
    expected = test_case.expected

    try:
        if feature == "cell_values":
            actual = read_cell_value_actual(adapter, workbook, sheet, cell, expected)
        elif feature == "text_formatting":
            actual = read_text_format_actual(adapter, workbook, sheet, cell)
        elif feature == "borders":
            actual = read_border_actual(adapter, workbook, sheet, cell)
        else:
            actual = {"error": f"Unknown feature: {feature}"}

        passed = compare_results(expected, actual)

        return TestResult(
            test_case_id=test_case.id,
            passed=passed,
            expected=expected,
            actual=actual,
        )

    except Exception as e:
        return TestResult(
            test_case_id=test_case.id,
            passed=False,
            expected=expected,
            actual={"error": str(e)},
            notes=f"Exception: {type(e).__name__}",
        )


def read_cell_value_actual(
    adapter: ExcelAdapter,
    workbook: Any,
    sheet: str,
    cell: str,
    expected: dict,
) -> dict:
    """Read cell value and return as comparable dict."""
    cell_value = adapter.read_cell_value(workbook, sheet, cell)

    result = {"type": cell_value.type.value}

    if cell_value.type.value != "blank":
        value = cell_value.value
        # Format dates/datetimes as strings for comparison
        if cell_value.type.value == "date" and hasattr(value, "isoformat"):
            result["value"] = value.strftime("%Y-%m-%d")
        elif cell_value.type.value == "datetime" and hasattr(value, "isoformat"):
            result["value"] = value.strftime("%Y-%m-%dT%H:%M:%S")
        else:
            result["value"] = value

    return result


def read_text_format_actual(
    adapter: ExcelAdapter,
    workbook: Any,
    sheet: str,
    cell: str,
) -> dict:
    """Read cell formatting and return as comparable dict."""
    fmt = adapter.read_cell_format(workbook, sheet, cell)

    result = {}
    if fmt.bold:
        result["bold"] = True
    if fmt.italic:
        result["italic"] = True
    if fmt.underline:
        result["underline"] = fmt.underline
    if fmt.strikethrough:
        result["strikethrough"] = True
    if fmt.font_name:
        result["font_name"] = fmt.font_name
    if fmt.font_size:
        result["font_size"] = fmt.font_size
    if fmt.font_color:
        result["font_color"] = fmt.font_color.upper()

    return result


def read_border_actual(
    adapter: ExcelAdapter,
    workbook: Any,
    sheet: str,
    cell: str,
) -> dict:
    """Read cell border and return as comparable dict."""
    border = adapter.read_cell_border(workbook, sheet, cell)

    result = {}

    # Check for uniform border style
    styles = []
    colors = []

    for edge_name, edge in [
        ("top", border.top),
        ("bottom", border.bottom),
        ("left", border.left),
        ("right", border.right),
    ]:
        if edge and edge.style.value != "none":
            result[f"border_{edge_name}"] = edge.style.value
            result[f"border_{edge_name}_color"] = edge.color.upper()
            styles.append(edge.style.value)
            colors.append(edge.color.upper())

    # If all edges have the same style, simplify
    if styles and len(set(styles)) == 1:
        result["border_style"] = styles[0]
        # Remove individual entries
        for edge in ["top", "bottom", "left", "right"]:
            result.pop(f"border_{edge}", None)

    if colors and len(set(colors)) == 1:
        result["border_color"] = colors[0]
        for edge in ["top", "bottom", "left", "right"]:
            result.pop(f"border_{edge}_color", None)

    # Diagonal borders
    if border.diagonal_up:
        result["border_diagonal_up"] = border.diagonal_up.style.value
    if border.diagonal_down:
        result["border_diagonal_down"] = border.diagonal_down.style.value

    return result


def test_write(
    adapter: ExcelAdapter,
    test_file: TestFile,
    file_path: Path,
) -> list[TestResult]:
    """Test writing a feature.

    Creates a new file with the adapter, then reads it back with
    openpyxl to verify correctness.

    Args:
        adapter: The adapter to test.
        test_file: TestFile metadata.
        file_path: Path to the original test file (for reference).

    Returns:
        List of TestResult for each test case.
    """
    # For write tests, we'd need to:
    # 1. Create a new file with the adapter
    # 2. Read it back with a trusted reader (openpyxl or xlwings)
    # 3. Compare results

    # For now, return empty results (write testing requires more infrastructure)
    # TODO: Implement write testing
    return []


def compare_results(expected: dict, actual: dict) -> bool:
    """Compare expected and actual results.

    Args:
        expected: Expected values.
        actual: Actual values read.

    Returns:
        True if results match.
    """
    if "error" in actual:
        return False

    # Check each expected key
    for key, exp_value in expected.items():
        if key not in actual:
            # Expected key missing
            if exp_value is not None:
                return False
            continue

        act_value = actual[key]

        # Handle color comparison (case-insensitive)
        if isinstance(exp_value, str) and exp_value.startswith("#"):
            if not isinstance(act_value, str):
                return False
            if exp_value.upper() != act_value.upper():
                return False
            continue

        # Handle numeric comparison with tolerance
        if isinstance(exp_value, (int, float)) and isinstance(act_value, (int, float)):
            if abs(exp_value - act_value) > 0.0001:
                return False
            continue

        # Direct comparison
        if exp_value != act_value:
            return False

    return True


def calculate_score(results: list[TestResult]) -> int:
    """Calculate fidelity score from test results.

    Simplified scoring for Phase 1:
    - 3: All test cases pass
    - 2: >80% pass, no critical failures
    - 1: >50% pass or basic cases work
    - 0: <50% pass or critical failures

    Args:
        results: List of test results.

    Returns:
        Score from 0-3.
    """
    if not results:
        return 0

    passed = sum(1 for r in results if r.passed)
    total = len(results)
    pass_rate = passed / total

    if pass_rate == 1.0:
        return 3
    elif pass_rate >= 0.8:
        return 2
    elif pass_rate >= 0.5:
        return 1
    else:
        return 0
