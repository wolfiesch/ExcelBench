# Library Expansion Tracker

Created: 02/08/2026 02:09 PM PST (via pst-timestamp)

## Overview

Tracks the status of all Excel library adapters in ExcelBench — implemented,
planned, and rejected. Each adapter wraps a specific library behind the
`ExcelAdapter` interface to produce comparable fidelity scores across 17
features (Tier 0/1/2).

Quick links:
- Benchmark results (xlsx): `results/xlsx/README.md`
- Benchmark results (xls): `results/xls/README.md`
- Adapter base class: `src/excelbench/harness/adapters/base.py`
- Adapter registry: `src/excelbench/harness/adapters/__init__.py`

## Adapter Inventory

### Implemented — xlsx profile (8 adapters)

| # | Library | Version | Lang | Caps | Read Score | Write Score | Green (R) | Green (W) | Notes |
|---|---------|---------|------|------|-----------|-------------|-----------|-----------|-------|
| 1 | openpyxl | 3.1.5 | py | R+W | 48/48 | 48/48 | 16/16 | 16/16 | Reference adapter, full fidelity |
| 2 | xlsxwriter | 3.2.9 | py | W | — | 48/48 | — | 16/16 | Write-only, full fidelity |
| 3 | python-calamine | 0.6.1 | py | R | 5/48 | — | 1/16 | — | Read-only, value+sheet only |
| 4 | pylightxl | 1.61 | py | R+W | 9/48 | 9/48 | 2/16 | 2/16 | Lightweight, value-only |
| 5 | pyexcel | 0.7.4 | py | R+W | 10/48 | 12/48 | 2/16 | 3/16 | Meta-library wrapping openpyxl |
| 6 | xlrd | 2.0.2 | py | R | — | — | — | — | .xls-only, not scored on xlsx |
| 7 | xlwt | 1.3.0 | py | W | — | 17/48 | — | 4/16 | .xls writer, limited xlsx compat |
| 8 | **pandas** | 3.0.0 | py | R+W | 5/48 | 12/48 | 1/16 | 3/16 | Abstraction-cost adapter (wraps openpyxl) |

### Implemented — xls profile (2 adapters)

| # | Library | Version | Caps | Green (R) | Notes |
|---|---------|---------|------|-----------|-------|
| 1 | python-calamine | 0.6.1 | R | 2/4 | Cross-format reader |
| 2 | xlrd | 2.0.2 | R | 4/4 | Full .xls read fidelity |

### Implemented — Rust/PyO3 (3 adapters, require compiled extension)

| # | Library | Lang | Caps | Status | Notes |
|---|---------|------|------|--------|-------|
| 1 | calamine (rust) | rust | R | Built, not in CI benchmark | PyO3 binding via excelbench_rust |
| 2 | rust_xlsxwriter | rust | W | Built, not in CI benchmark | PyO3 binding via excelbench_rust |
| 3 | umya-spreadsheet | rust | R+W | Built, not in CI benchmark | PyO3 binding via excelbench_rust |

### Planned / Candidate

| ID | Library | Lang | Caps | Priority | Rationale |
|----|---------|------|------|----------|-----------|
| A2 | odfpy | py | R+W | P2 | ODS format support (not xlsx) |
| A3 | et-xmlfile | py | — | P3 | Low-level XML streaming (used by openpyxl internally) |
| A4 | openpyxl readonly mode | py | R | P1 | Memory-optimized read — test if scores differ from default mode |
| A5 | xlsxwriter constant_memory | py | W | P1 | Memory-optimized write — test if scores differ |
| A6 | polars | py/rust | R | P1 | Rust DataFrame reader, compare to pandas abstraction cost |
| A7 | tablib | py | R+W | P2 | Dataset-oriented library with xlsx support |

### Rejected / Out of Scope

| Library | Reason |
|---------|--------|
| xlwings (as adapter) | Already used as test-file generator / oracle; not a parsing library |
| csv | Not an Excel format |
| openpyxl write_only | Streaming write API — requires different adapter pattern (no random cell access) |

## Score Summary — xlsx profile

Extracted from `results/xlsx/matrix.csv` (02/08/2026 run):

```
Feature              openpyxl  xlsxwriter  calamine  pylightxl  pyexcel  xlwt  pandas
                     R  W      W           R         R  W        R  W     W     R  W
cell_values          3  3      3           1         3  1        3  3     3     1  3
formulas             3  3      3           0         0  3        0  3     0     0  3
text_formatting      3  3      3           0         0  0        0  0     1     0  0
background_colors    3  3      3           0         0  0        0  0     1     0  0
number_formats       3  3      3           0         0  0        0  1     3     0  0
alignment            3  3      3           1         0  1        1  1     3     1  1
borders              3  3      3           0         0  0        0  0     1     0  0
dimensions           3  3      3           0         0  0        0  0     1     0  0
multiple_sheets      3  3      3           3         3  3        3  3     3     3  3
merged_cells         3  3      3           0         0  0        0  0     0     0  0
conditional_format   3  3      3           0         0  0        0  0     0     0  0
data_validation      3  3      3           0         0  0        0  0     0     0  0
hyperlinks           3  3      3           0         0  0        0  0     0     0  0
images               3  3      3           0         0  0        0  0     0     0  0
comments             3  3      3           0         0  0        0  0     0     0  0
freeze_panes         3  3      3           0         0  0        0  0     0     0  0
```

## pandas vs pyexcel — Abstraction Cost Analysis

Both libraries wrap openpyxl internally. Key differences:

| Metric | pandas | pyexcel | Winner |
|--------|--------|---------|--------|
| cell_values read | 1 (errors lost) | 3 (errors preserved) | pyexcel |
| cell_values write | 3 | 3 | tie |
| formulas read | 0 | 0 | tie |
| formulas write | 3 | 3 | tie |
| alignment read | 1 | 1 | tie |
| number_formats write | 0 | 1 | pyexcel |
| Green features (R) | 1/16 | 2/16 | pyexcel |
| Green features (W) | 3/16 | 3/16 | tie |

**Key finding:** pandas loses error values (`#DIV/0!`, `#N/A`, `#VALUE!`) during
`pd.read_excel()` because DataFrames coerce them to `NaN`. pyexcel preserves
errors as strings through its cell iterator. For value-only reads, pyexcel is
the safer abstraction.

## Checklist — Expansion Tasks

- [x] A1: pandas adapter (value-only R+W, abstraction-cost measurement)
- [ ] A4: openpyxl readonly mode adapter
- [ ] A5: xlsxwriter constant_memory mode adapter
- [ ] A6: polars adapter (Rust DataFrame reader)
- [ ] A2: odfpy adapter (ODS format)
- [ ] A7: tablib adapter
- [ ] A3: et-xmlfile investigation

## Session Log (append-only)

### 02/08/2026 02:09 PM PST (via pst-timestamp)
- Worked on: A1 (pandas adapter)
- Committed: `46b8a87 feat(adapter): add pandas adapter measuring abstraction cost vs openpyxl`
- Scores: Read 1/3 cell_values (errors→NaN), Write 3/3 cell_values; 1/16 green R, 3/16 green W
- Benchmark: Regenerated full xlsx+xls profiles with all 8 Python adapters
- Decisions: pandas errors-as-NaN is a genuine abstraction cost, not a bug to fix
- Next: A4 (openpyxl readonly mode) or A6 (polars) — both measure optimization variants
