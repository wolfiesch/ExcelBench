"""Test runner for executing benchmarks."""

import os
import platform
import tempfile
from datetime import UTC, datetime
from pathlib import Path
from typing import Any

from excelbench.generator.generate import load_manifest
from excelbench.harness.adapters import (
    ExcelAdapter,
    ExcelOracleAdapter,
    OpenpyxlAdapter,
    get_all_adapters,
)
from excelbench.models import (
    BenchmarkMetadata,
    BenchmarkResults,
    BorderEdge,
    BorderInfo,
    BorderStyle,
    CellFormat,
    CellType,
    CellValue,
    FeatureScore,
    OperationType,
    TestCase,
    TestFile,
    TestResult,
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
        run_date=datetime.now(UTC),
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
                operation=OperationType.READ,
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

        default_sheet = sheet_names[0]

        for tc in test_file.test_cases:
            result = test_read_case(
                adapter,
                workbook,
                default_sheet,
                tc,
                test_file.feature,
                OperationType.READ,
            )
            results.append(result)

    finally:
        adapter.close_workbook(workbook)

    return results


def test_read_case(
    adapter: ExcelAdapter,
    workbook: Any,
    default_sheet: str,
    test_case: TestCase,
    feature: str,
    operation: OperationType,
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
    expected = test_case.expected

    if feature == "multiple_sheets" and "sheet_names" in expected:
        actual = read_sheet_names_actual(adapter, workbook)
        passed = compare_results(expected, actual)
        return TestResult(
            test_case_id=test_case.id,
            operation=operation,
            passed=passed,
            expected=expected,
            actual=actual,
        )

    sheet = test_case.sheet or feature or default_sheet
    cell = test_case.cell or f"B{test_case.row}"

    try:
        if feature == "cell_values":
            actual = read_cell_value_actual(adapter, workbook, sheet, cell, expected)
        elif feature == "formulas":
            actual = read_formula_actual(adapter, workbook, sheet, cell)
        elif feature == "text_formatting":
            actual = read_text_format_actual(adapter, workbook, sheet, cell)
        elif feature == "background_colors":
            actual = read_background_color_actual(adapter, workbook, sheet, cell)
        elif feature == "number_formats":
            actual = read_number_format_actual(adapter, workbook, sheet, cell)
        elif feature == "alignment":
            actual = read_alignment_actual(adapter, workbook, sheet, cell)
        elif feature == "borders":
            actual = read_border_actual(adapter, workbook, sheet, cell)
        elif feature == "dimensions":
            actual = read_dimensions_actual(adapter, workbook, sheet, cell, test_case)
        elif feature == "multiple_sheets":
            # Non-sheet_names test cases are cell value reads from specific sheets
            actual = read_cell_value_actual(adapter, workbook, sheet, cell, expected)
        else:
            actual = {"error": f"Unknown feature: {feature}"}

        passed = compare_results(expected, actual)

        return TestResult(
            test_case_id=test_case.id,
            operation=operation,
            passed=passed,
            expected=expected,
            actual=actual,
        )

    except Exception as e:
        return TestResult(
            test_case_id=test_case.id,
            operation=operation,
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


def read_formula_actual(
    adapter: ExcelAdapter,
    workbook: Any,
    sheet: str,
    cell: str,
) -> dict:
    cell_value = adapter.read_cell_value(workbook, sheet, cell)
    if cell_value.type != CellType.FORMULA:
        return {"error": f"Expected formula, got {cell_value.type.value}"}
    formula = cell_value.formula or cell_value.value
    return {"type": "formula", "formula": formula}


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


def read_background_color_actual(
    adapter: ExcelAdapter,
    workbook: Any,
    sheet: str,
    cell: str,
) -> dict:
    fmt = adapter.read_cell_format(workbook, sheet, cell)
    result = {}
    if fmt.bg_color:
        result["bg_color"] = fmt.bg_color.upper()
    return result


def read_number_format_actual(
    adapter: ExcelAdapter,
    workbook: Any,
    sheet: str,
    cell: str,
) -> dict:
    fmt = adapter.read_cell_format(workbook, sheet, cell)
    result = {}
    if fmt.number_format:
        result["number_format"] = fmt.number_format
    return result


def read_alignment_actual(
    adapter: ExcelAdapter,
    workbook: Any,
    sheet: str,
    cell: str,
) -> dict:
    fmt = adapter.read_cell_format(workbook, sheet, cell)
    result = {}
    if fmt.h_align:
        result["h_align"] = fmt.h_align
    if fmt.v_align:
        result["v_align"] = fmt.v_align
    if fmt.wrap is not None:
        result["wrap"] = fmt.wrap
    if fmt.rotation is not None:
        result["rotation"] = fmt.rotation
    if fmt.indent is not None:
        result["indent"] = fmt.indent
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

    # Only simplify to border_style/border_color when ALL 4 edges are present and identical
    if len(styles) == 4 and len(set(styles)) == 1:
        result["border_style"] = styles[0]
        # Remove individual entries
        for edge in ["top", "bottom", "left", "right"]:
            result.pop(f"border_{edge}", None)

    if len(colors) == 4 and len(set(colors)) == 1:
        result["border_color"] = colors[0]
        for edge in ["top", "bottom", "left", "right"]:
            result.pop(f"border_{edge}_color", None)

    # Diagonal borders
    if border.diagonal_up:
        result["border_diagonal_up"] = border.diagonal_up.style.value
    if border.diagonal_down:
        result["border_diagonal_down"] = border.diagonal_down.style.value

    return result


def read_dimensions_actual(
    adapter: ExcelAdapter,
    workbook: Any,
    sheet: str,
    cell: str,
    test_case: TestCase,
) -> dict:
    result = {}
    row = test_case.row
    column = _extract_column(cell)
    if "row_height" in test_case.expected:
        result["row_height"] = adapter.read_row_height(workbook, sheet, row)
    if "column_width" in test_case.expected:
        result["column_width"] = adapter.read_column_width(workbook, sheet, column)
    return result


def read_sheet_names_actual(adapter: ExcelAdapter, workbook: Any) -> dict:
    return {"sheet_names": adapter.get_sheet_names(workbook)}


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
    results: list[TestResult] = []

    verifier = get_write_verifier()

    with tempfile.TemporaryDirectory() as tmpdir:
        output_dir = Path(tmpdir) / adapter.name
        output_dir.mkdir(parents=True, exist_ok=True)
        output_path = output_dir / f"{test_file.feature}.xlsx"

        try:
            workbook = adapter.create_workbook()

            sheet_names = _collect_sheet_names(test_file)
            if not sheet_names:
                sheet_names = [test_file.feature]

            for name in sheet_names:
                adapter.add_sheet(workbook, name)

            for tc in test_file.test_cases:
                if "sheet_names" in tc.expected:
                    continue

                target_sheet = tc.sheet or test_file.feature
                target_cell = tc.cell or f"B{tc.row}"

                if test_file.feature == "cell_values":
                    _write_cell_value_case(adapter, workbook, target_sheet, target_cell, tc.expected)
                elif test_file.feature == "formulas":
                    _write_formula_case(adapter, workbook, target_sheet, target_cell, tc.expected)
                elif test_file.feature == "text_formatting":
                    _write_text_format_case(adapter, workbook, target_sheet, target_cell, tc)
                elif test_file.feature == "background_colors":
                    _write_background_color_case(
                        adapter,
                        workbook,
                        target_sheet,
                        target_cell,
                        tc.expected,
                    )
                elif test_file.feature == "number_formats":
                    _write_number_format_case(
                        adapter,
                        workbook,
                        target_sheet,
                        target_cell,
                        tc.expected,
                    )
                elif test_file.feature == "alignment":
                    _write_alignment_case(
                        adapter,
                        workbook,
                        target_sheet,
                        target_cell,
                        tc.expected,
                    )
                elif test_file.feature == "borders":
                    _write_border_case(
                        adapter,
                        workbook,
                        target_sheet,
                        target_cell,
                        tc.expected,
                    )
                elif test_file.feature == "dimensions":
                    _write_dimensions_case(adapter, workbook, target_sheet, target_cell, tc)
                elif test_file.feature == "multiple_sheets":
                    _write_multi_sheet_case(
                        adapter,
                        workbook,
                        target_sheet,
                        target_cell,
                        tc.expected,
                    )

            adapter.save_workbook(workbook, output_path)
        except Exception as e:
            for tc in test_file.test_cases:
                results.append(TestResult(
                    test_case_id=tc.id,
                    operation=OperationType.WRITE,
                    passed=False,
                    expected=tc.expected,
                    actual={"error": str(e)},
                    notes=f"Write failed: {type(e).__name__}",
                ))
            return results

        try:
            verify_wb = verifier.open_workbook(output_path)
        except Exception as e:
            for tc in test_file.test_cases:
                results.append(TestResult(
                    test_case_id=tc.id,
                    operation=OperationType.WRITE,
                    passed=False,
                    expected=tc.expected,
                    actual={"error": str(e)},
                    notes="Failed to open workbook for verification",
                ))
            return results

        try:
            sheet_names = verifier.get_sheet_names(verify_wb)
            default_sheet = sheet_names[0] if sheet_names else test_file.feature
            for tc in test_file.test_cases:
                result = test_read_case(
                    verifier,
                    verify_wb,
                    default_sheet,
                    tc,
                    test_file.feature,
                    OperationType.WRITE,
                )
                results.append(result)
        finally:
            verifier.close_workbook(verify_wb)

    return results


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


def get_write_verifier() -> ExcelAdapter:
    oracle = os.environ.get("EXCELBENCH_WRITE_ORACLE", "auto").lower()
    if oracle == "openpyxl":
        return OpenpyxlAdapter()
    if oracle == "excel":
        return ExcelOracleAdapter()
    if _excel_available():
        return ExcelOracleAdapter()
    return OpenpyxlAdapter()


def _excel_available() -> bool:
    try:
        import xlwings as xw
        app = xw.App(visible=False, add_book=False)
        app.quit()
        return True
    except Exception:
        return False


def _collect_sheet_names(test_file: TestFile) -> list[str]:
    sheet_names: list[str] = []
    explicit = False
    for tc in test_file.test_cases:
        if "sheet_names" in tc.expected:
            sheet_names = list(tc.expected["sheet_names"])
            explicit = True
            break
        if test_file.feature == "formulas":
            sheet_names.extend(_extract_formula_sheet_names(tc.expected.get("formula", "")))
    # Ensure the feature name is included unless sheets were explicitly listed
    if not explicit and test_file.feature not in sheet_names:
        sheet_names.insert(0, test_file.feature)
    for tc in test_file.test_cases:
        name = tc.sheet
        if name and name not in sheet_names:
            sheet_names.append(name)
    return sheet_names


def _extract_column(cell: str) -> str:
    import re
    match = re.match(r"([A-Z]+)", cell.upper())
    if not match:
        return "B"
    return match.group(1)


def _extract_formula_sheet_names(formula: str) -> list[str]:
    import re
    names: list[str] = []
    if not formula:
        return names
    # Sheet names in formulas: 'Sheet Name'!A1 or Sheet1!A1
    for match in re.findall(r"'([^']+)'!", formula):
        names.append(match)
    for match in re.findall(r"\b([A-Za-z0-9_]+)!", formula):
        if match not in names:
            names.append(match)
    return names


def _cell_value_from_expected(expected: dict) -> CellValue:
    type_str = expected.get("type", "string")
    value = expected.get("value")
    if type_str == "blank":
        return CellValue(type=CellType.BLANK)
    if type_str == "boolean":
        return CellValue(type=CellType.BOOLEAN, value=bool(value))
    if type_str == "number":
        return CellValue(type=CellType.NUMBER, value=value)
    if type_str == "date":
        from datetime import date as _date
        if isinstance(value, str):
            value = _date.fromisoformat(value)
        return CellValue(type=CellType.DATE, value=value)
    if type_str == "datetime":
        from datetime import datetime as _datetime
        if isinstance(value, str):
            value = _datetime.fromisoformat(value)
        return CellValue(type=CellType.DATETIME, value=value)
    if type_str == "error":
        return CellValue(type=CellType.ERROR, value=value)
    if type_str == "formula":
        formula = expected.get("formula") or value
        return CellValue(type=CellType.FORMULA, value=value, formula=formula)
    return CellValue(type=CellType.STRING, value=value)


def _cell_format_from_expected(expected: dict) -> CellFormat:
    return CellFormat(
        bold=expected.get("bold"),
        italic=expected.get("italic"),
        underline=expected.get("underline"),
        strikethrough=expected.get("strikethrough"),
        font_name=expected.get("font_name"),
        font_size=expected.get("font_size"),
        font_color=expected.get("font_color"),
        bg_color=expected.get("bg_color"),
        number_format=expected.get("number_format"),
        h_align=expected.get("h_align"),
        v_align=expected.get("v_align"),
        wrap=expected.get("wrap"),
        rotation=expected.get("rotation"),
        indent=expected.get("indent"),
    )


def _border_from_expected(expected: dict) -> BorderInfo:
    default_style = expected.get("border_style")
    default_color = expected.get("border_color")
    if default_color and not default_style:
        default_style = "thin"

    def make_edge(style_key: str, color_key: str) -> BorderEdge | None:
        if style_key in expected:
            style_val = expected[style_key]
        else:
            style_val = default_style

        if style_val is None:
            return None

        if color_key in expected:
            color_val = expected[color_key]
        else:
            color_val = default_color

        if color_val is None:
            color_val = "#000000"

        return BorderEdge(style=BorderStyle(style_val), color=color_val)

    return BorderInfo(
        top=make_edge("border_top", "border_top_color"),
        bottom=make_edge("border_bottom", "border_bottom_color"),
        left=make_edge("border_left", "border_left_color"),
        right=make_edge("border_right", "border_right_color"),
        diagonal_up=make_edge("border_diagonal_up", "border_color")
        if expected.get("border_diagonal_up") is not None
        else None,
        diagonal_down=make_edge("border_diagonal_down", "border_color")
        if expected.get("border_diagonal_down") is not None
        else None,
    )


def _write_cell_value_case(
    adapter: ExcelAdapter,
    workbook: Any,
    sheet: str,
    cell: str,
    expected: dict,
) -> None:
    cell_value = _cell_value_from_expected(expected)
    adapter.write_cell_value(workbook, sheet, cell, cell_value)

    if cell_value.type in (CellType.DATE, CellType.DATETIME):
        number_format = "yyyy-mm-dd" if cell_value.type == CellType.DATE else "yyyy-mm-dd hh:mm:ss"
        adapter.write_cell_format(workbook, sheet, cell, CellFormat(number_format=number_format))


def _write_formula_case(
    adapter: ExcelAdapter,
    workbook: Any,
    sheet: str,
    cell: str,
    expected: dict,
) -> None:
    formula = expected.get("formula")
    cell_value = CellValue(type=CellType.FORMULA, formula=formula, value=formula)
    adapter.write_cell_value(workbook, sheet, cell, cell_value)


def _write_text_format_case(
    adapter: ExcelAdapter,
    workbook: Any,
    sheet: str,
    cell: str,
    test_case: TestCase,
) -> None:
    adapter.write_cell_value(
        workbook,
        sheet,
        cell,
        CellValue(type=CellType.STRING, value=test_case.label),
    )
    adapter.write_cell_format(workbook, sheet, cell, _cell_format_from_expected(test_case.expected))


def _write_background_color_case(
    adapter: ExcelAdapter,
    workbook: Any,
    sheet: str,
    cell: str,
    expected: dict,
) -> None:
    adapter.write_cell_value(workbook, sheet, cell, CellValue(type=CellType.STRING, value="Color"))
    adapter.write_cell_format(workbook, sheet, cell, _cell_format_from_expected(expected))


def _write_number_format_case(
    adapter: ExcelAdapter,
    workbook: Any,
    sheet: str,
    cell: str,
    expected: dict,
) -> None:
    number_format = expected.get("number_format")
    value = 1234.5
    value_type = CellType.NUMBER
    if number_format and any(token in number_format for token in ["y", "m", "d"]):
        from datetime import date
        value = date(2026, 2, 4)
        value_type = CellType.DATE
    adapter.write_cell_value(workbook, sheet, cell, CellValue(type=value_type, value=value))
    adapter.write_cell_format(workbook, sheet, cell, _cell_format_from_expected(expected))


def _write_alignment_case(
    adapter: ExcelAdapter,
    workbook: Any,
    sheet: str,
    cell: str,
    expected: dict,
) -> None:
    adapter.write_cell_value(workbook, sheet, cell, CellValue(type=CellType.STRING, value="Align"))
    adapter.write_cell_format(workbook, sheet, cell, _cell_format_from_expected(expected))


def _write_border_case(
    adapter: ExcelAdapter,
    workbook: Any,
    sheet: str,
    cell: str,
    expected: dict,
) -> None:
    adapter.write_cell_value(workbook, sheet, cell, CellValue(type=CellType.STRING, value="Border"))
    adapter.write_cell_border(workbook, sheet, cell, _border_from_expected(expected))


def _write_dimensions_case(
    adapter: ExcelAdapter,
    workbook: Any,
    sheet: str,
    cell: str,
    test_case: TestCase,
) -> None:
    if "row_height" in test_case.expected:
        adapter.set_row_height(workbook, sheet, test_case.row, test_case.expected["row_height"])
    if "column_width" in test_case.expected:
        column = _extract_column(cell)
        adapter.set_column_width(workbook, sheet, column, test_case.expected["column_width"])


def _write_multi_sheet_case(
    adapter: ExcelAdapter,
    workbook: Any,
    sheet: str,
    cell: str,
    expected: dict,
) -> None:
    if "sheet_names" in expected:
        return
    if "type" in expected:
        _write_cell_value_case(adapter, workbook, sheet, cell, expected)


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
