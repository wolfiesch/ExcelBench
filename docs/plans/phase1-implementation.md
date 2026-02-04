# Phase 1: Foundation (MVP) - Implementation Plan

## Goal
Get a working end-to-end benchmark that tests 3 features across 2 libraries, producing viewable results.

## Scope

**Features to implement:**
1. Cell Values (strings, numbers, dates, booleans, errors)
2. Basic Text Formatting (bold, italic, font size, font color)
3. Borders (styles, colors, positions)

**Libraries to test:**
1. openpyxl (read + write)
2. xlsxwriter (write only)

**Output:**
- `results/results.json` - structured benchmark data
- `results/README.md` - human-readable summary table

---

## Implementation Steps

### Step 1: Generator Framework
**Files:** `src/excelbench/generator/base.py`, `src/excelbench/generator/generate.py`

Create the xlwings-based generator framework:
```python
# base.py - Abstract base for feature generators
class FeatureGenerator(Protocol):
    feature_name: str
    tier: int

    def generate(self, workbook: xw.Book) -> list[TestCase]:
        """Generate test cases in the workbook, return metadata."""
        ...

# generate.py - Main entry point
def generate_all(output_dir: Path) -> Manifest:
    """Generate all test files and return manifest."""
    ...
```

**Test cases structure:**
- Column A: Label describing what's being tested
- Column B: The actual test cell with the feature applied
- Column C: Expected value in parseable format (JSON or key=value)

---

### Step 2: Cell Values Generator
**File:** `src/excelbench/generator/features/cell_values.py`

Test cases to generate:
```
| Label                    | Test Cell          | Expected                          |
|--------------------------|--------------------|------------------------------------|
| String - simple          | "Hello World"      | type=string, value="Hello World"  |
| String - unicode         | "日本語🎉"          | type=string, value="日本語🎉"      |
| String - empty           | ""                 | type=string, value=""             |
| Number - integer         | 42                 | type=number, value=42             |
| Number - float           | 3.14159            | type=number, value=3.14159        |
| Number - negative        | -100.5             | type=number, value=-100.5         |
| Number - large           | 1234567890123      | type=number, value=1234567890123  |
| Date - standard          | 2026-02-04         | type=date, value=2026-02-04       |
| Date - with time         | 2026-02-04 10:30   | type=datetime, value=...          |
| Boolean - true           | TRUE               | type=boolean, value=true          |
| Boolean - false          | FALSE              | type=boolean, value=false         |
| Error - DIV/0            | =1/0               | type=error, value=#DIV/0!         |
| Error - NA               | =NA()              | type=error, value=#N/A            |
| Blank cell               | (empty)            | type=blank                        |
```

---

### Step 3: Basic Formatting Generator
**File:** `src/excelbench/generator/features/text_formatting.py`

Test cases to generate:
```
| Label                    | Test Cell          | Expected                          |
|--------------------------|--------------------|------------------------------------|
| Bold                     | **Text**           | bold=true                         |
| Italic                   | *Text*             | italic=true                       |
| Underline                | _Text_             | underline=single                  |
| Strikethrough            | ~~Text~~           | strikethrough=true                |
| Bold + Italic            | ***Text***         | bold=true, italic=true            |
| Font size 8              | Text               | font_size=8                       |
| Font size 14             | Text               | font_size=14                      |
| Font size 24             | Text               | font_size=24                      |
| Font Arial               | Text               | font_name=Arial                   |
| Font Times New Roman     | Text               | font_name=Times New Roman         |
| Font color red           | Text               | font_color=#FF0000                |
| Font color blue          | Text               | font_color=#0000FF                |
| Font color custom        | Text               | font_color=#8B4513                |
| Combined formatting      | Text               | bold=true, font_size=16, font_color=#FF0000 |
```

---

### Step 4: Borders Generator
**File:** `src/excelbench/generator/features/borders.py`

Test cases to generate:
```
| Label                    | Test Cell          | Expected                          |
|--------------------------|--------------------|------------------------------------|
| Border - thin all        | [box]              | border_style=thin, border_color=#000000 |
| Border - medium all      | [box]              | border_style=medium               |
| Border - thick all       | [box]              | border_style=thick                |
| Border - double          | [box]              | border_style=double               |
| Border - dashed          | [box]              | border_style=dashed               |
| Border - dotted          | [box]              | border_style=dotted               |
| Border - top only        | [top]              | border_top=thin                   |
| Border - bottom only     | [bottom]           | border_bottom=thin                |
| Border - left only       | [left]             | border_left=thin                  |
| Border - right only      | [right]            | border_right=thin                 |
| Border - diagonal up     | [diag]             | border_diagonal_up=thin           |
| Border - diagonal down   | [diag]             | border_diagonal_down=thin         |
| Border - color red       | [box]              | border_color=#FF0000              |
| Border - mixed edges     | [mixed]            | border_top=thick, border_bottom=thin |
```

---

### Step 5: Manifest Generator
**File:** `src/excelbench/generator/manifest.py`

After generating all test files, create `test_files/manifest.json`:
```json
{
  "generated_at": "2026-02-04T10:30:00Z",
  "excel_version": "16.83",
  "generator_version": "0.1.0",
  "files": [
    {
      "path": "tier1/01_cell_values.xlsx",
      "feature": "cell_values",
      "tier": 1,
      "test_cases": [
        {"id": "string_simple", "row": 2, "label": "String - simple"},
        ...
      ]
    }
  ]
}
```

---

### Step 6: Adapter Protocol
**File:** `src/excelbench/harness/adapters/base.py`

Define the interface all library adapters must implement:
```python
from typing import Protocol, Any
from pathlib import Path

class CellValue:
    type: str  # "string", "number", "date", "boolean", "error", "blank"
    value: Any

class CellFormat:
    bold: bool | None
    italic: bool | None
    underline: str | None
    strikethrough: bool | None
    font_name: str | None
    font_size: float | None
    font_color: str | None  # hex color

class BorderInfo:
    top: BorderEdge | None
    bottom: BorderEdge | None
    left: BorderEdge | None
    right: BorderEdge | None
    diagonal_up: BorderEdge | None
    diagonal_down: BorderEdge | None

class BorderEdge:
    style: str  # thin, medium, thick, dashed, dotted, double, etc.
    color: str  # hex color

class ExcelAdapter(Protocol):
    name: str
    version: str
    capabilities: set[str]  # {"read", "write"}

    # Read operations
    def read_cell_value(self, path: Path, sheet: str, cell: str) -> CellValue: ...
    def read_cell_format(self, path: Path, sheet: str, cell: str) -> CellFormat: ...
    def read_cell_border(self, path: Path, sheet: str, cell: str) -> BorderInfo: ...

    # Write operations (for write-capable libraries)
    def create_workbook(self) -> Any: ...
    def write_cell_value(self, wb: Any, sheet: str, cell: str, value: CellValue) -> None: ...
    def write_cell_format(self, wb: Any, sheet: str, cell: str, fmt: CellFormat) -> None: ...
    def write_cell_border(self, wb: Any, sheet: str, cell: str, border: BorderInfo) -> None: ...
    def save_workbook(self, wb: Any, path: Path) -> None: ...
```

---

### Step 7: openpyxl Adapter
**File:** `src/excelbench/harness/adapters/openpyxl_adapter.py`

Implement the adapter for openpyxl (read + write):
- Map openpyxl's cell types to our CellValue
- Map openpyxl's Font object to our CellFormat
- Map openpyxl's Border/Side objects to our BorderInfo
- Implement all write operations

---

### Step 8: xlsxwriter Adapter
**File:** `src/excelbench/harness/adapters/xlsxwriter_adapter.py`

Implement the adapter for xlsxwriter (write only):
- Read operations raise NotImplementedError
- Map our types to xlsxwriter's format system
- Handle xlsxwriter's "format once, apply many" pattern

---

### Step 9: Test Runner
**File:** `src/excelbench/harness/runner.py`

Orchestrate the benchmark:
```python
def run_benchmark(
    manifest_path: Path,
    adapters: list[ExcelAdapter],
    output_dir: Path
) -> BenchmarkResults:
    """
    1. Load manifest
    2. For each test file:
       a. For each adapter:
          - If READ test: read test cells, compare to expected
          - If WRITE test: generate file from spec, verify with xlwings
       b. Score fidelity per rubric
    3. Return structured results
    """
```

---

### Step 10: Scoring Engine
**File:** `src/excelbench/harness/scoring.py`

Apply fidelity rubrics to test results:
```python
def score_feature(
    feature: str,
    test_results: list[TestResult],
    rubric: FeatureRubric
) -> FeatureScore:
    """
    Given test case pass/fail/partial results,
    determine the 0-3 fidelity score based on rubric.
    """
```

For Phase 1, implement simplified scoring:
- 3: All test cases pass
- 2: >80% pass, no critical failures
- 1: >50% pass or basic cases work
- 0: <50% pass or critical failures

---

### Step 11: Results Renderer
**File:** `src/excelbench/results/renderer.py`

Generate output files:
```python
def render_results(results: BenchmarkResults, output_dir: Path) -> None:
    # Write JSON
    write_json(results, output_dir / "results.json")

    # Generate markdown summary
    render_markdown_summary(results, output_dir / "README.md")
```

Markdown format:
```markdown
# ExcelBench Results

*Generated: 2026-02-04*

## Summary

| Feature | openpyxl (R) | openpyxl (W) | xlsxwriter (W) |
|---------|--------------|--------------|----------------|
| Cell Values | 🟢 3 | 🟢 3 | 🟢 3 |
| Text Formatting | 🟡 2 | 🟢 3 | 🟢 3 |
| Borders | 🟡 2 | 🟡 2 | 🟢 3 |

## Detailed Results
...
```

---

### Step 12: CLI Entry Point
**File:** `src/excelbench/cli.py`

```python
import typer
app = typer.Typer()

@app.command()
def generate(output_dir: Path = Path("test_files")):
    """Generate test Excel files using xlwings."""
    ...

@app.command()
def benchmark(
    test_dir: Path = Path("test_files"),
    output_dir: Path = Path("results")
):
    """Run benchmark against all adapters."""
    ...

@app.command()
def report(results_path: Path = Path("results/results.json")):
    """Regenerate reports from existing results."""
    ...
```

---

## Verification

After implementation, verify end-to-end:

1. **Generate test files:**
   ```bash
   uv run excelbench generate
   # Opens Excel, creates test_files/*.xlsx
   # Verify: open test_files/tier1/01_cell_values.xlsx in Excel
   ```

2. **Run benchmark:**
   ```bash
   uv run excelbench benchmark
   # Tests openpyxl and xlsxwriter against generated files
   # Creates results/results.json and results/README.md
   ```

3. **Check results:**
   ```bash
   cat results/README.md
   # Should show comparison table with scores
   ```

4. **Manual verification:**
   - Open a generated test file in Excel - labels should match test cells
   - Check that scores align with rubric definitions
   - Verify JSON structure matches schema

---

## File Checklist

- [ ] `src/excelbench/generator/base.py`
- [ ] `src/excelbench/generator/generate.py`
- [ ] `src/excelbench/generator/manifest.py`
- [ ] `src/excelbench/generator/features/__init__.py`
- [ ] `src/excelbench/generator/features/cell_values.py`
- [ ] `src/excelbench/generator/features/text_formatting.py`
- [ ] `src/excelbench/generator/features/borders.py`
- [ ] `src/excelbench/harness/adapters/base.py`
- [ ] `src/excelbench/harness/adapters/openpyxl_adapter.py`
- [ ] `src/excelbench/harness/adapters/xlsxwriter_adapter.py`
- [ ] `src/excelbench/harness/runner.py`
- [ ] `src/excelbench/harness/scoring.py`
- [ ] `src/excelbench/results/__init__.py`
- [ ] `src/excelbench/results/renderer.py`
- [ ] `src/excelbench/cli.py`
- [ ] `scripts/generate_test_files.py` (thin wrapper)
- [ ] `scripts/run_benchmark.py` (thin wrapper)

---

## Estimated Complexity

| Component | Lines (est.) | Complexity |
|-----------|--------------|------------|
| Generator framework | 150 | Medium |
| 3 feature generators | 300 | Low |
| Adapter protocol | 100 | Low |
| openpyxl adapter | 200 | Medium |
| xlsxwriter adapter | 150 | Medium |
| Test runner | 150 | Medium |
| Scoring engine | 100 | Low |
| Results renderer | 100 | Low |
| CLI | 50 | Low |
| **Total** | **~1300** | |

---

## Dependencies on External Tools

- **Excel** must be installed and have automation permissions
- **xlwings** requires Excel to be launchable via AppleScript
- First run may prompt for accessibility permissions

Test xlwings setup:
```python
import xlwings as xw
wb = xw.Book()  # Should open Excel with new workbook
wb.close()
```
