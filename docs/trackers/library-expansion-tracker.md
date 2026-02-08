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

### Implemented — xlsx profile (12 adapters)

| # | Library | Version | Lang | Caps | Read Score | Write Score | Green (R) | Green (W) | Notes |
|---|---------|---------|------|------|-----------|-------------|-----------|-----------|-------|
| 1 | openpyxl | 3.1.5 | py | R+W | 48/48 | 48/48 | 16/16 | 16/16 | Reference adapter, full fidelity |
| 2 | xlsxwriter | 3.2.9 | py | W | — | 48/48 | — | 16/16 | Write-only, full fidelity |
| 3 | python-calamine | 0.6.1 | py | R | 5/48 | — | 1/16 | — | Read-only, value+sheet only |
| 4 | pylightxl | 1.61 | py | R+W | 9/48 | 9/48 | 2/16 | 2/16 | Lightweight, value-only |
| 5 | pyexcel | 0.7.4 | py | R+W | 10/48 | 12/48 | 2/16 | 3/16 | Meta-library wrapping openpyxl |
| 6 | xlrd | 2.0.2 | py | R | — | — | — | — | .xls-only, not scored on xlsx |
| 7 | xlwt | 1.3.0 | py | W | — | 17/48 | — | 4/16 | .xls writer, limited xlsx compat |
| 8 | pandas | 3.0.0 | py | R+W | 5/48 | 12/48 | 1/16 | 3/16 | Abstraction-cost adapter (wraps openpyxl) |
| 9 | **openpyxl-readonly** | 3.1.5 | py | R | 10/48 | — | 3/16 | — | Streaming read mode, limited formatting |
| 10 | **xlsxwriter-constmem** | 3.2.9 | py | W | — | 43/48 | — | 13/16 | Row-major write, no images/comments |
| 11 | **polars** | 1.38.1 | py/rust | R | 4/48 | — | 0/16 | — | Rust calamine backend, type coercion cost |
| 12 | **tablib** | 3.9.0 | py | R+W | 10/48 | 12/48 | 2/16 | 3/16 | Dataset model wrapping openpyxl |

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

### Rejected / Out of Scope

| Library | Reason |
|---------|--------|
| xlwings (as adapter) | Already used as test-file generator / oracle; not a parsing library |
| csv | Not an Excel format |
| openpyxl write_only | Streaming write API — requires different adapter pattern (no random cell access) |

## Score Summary — xlsx profile

Extracted from `results/xlsx/README.md` (02/08/2026 run, 12 adapters):

```
Feature              openpyxl  xlsxwriter  constmem  calamine  opxl-ro  pylightxl  pyexcel  xlwt  pandas  polars  tablib
                     R  W      W           W         R         R        R  W        R  W     W     R  W    R       R  W
cell_values          3  3      3           3         1         3        3  1        3  3     3     1  3    1       3  3
formulas             3  3      3           3         0         3        0  3        0  3     0     0  3    0       0  3
text_formatting      3  3      3           3         0         0        0  0        0  0     1     0  0    0       0  0
background_colors    3  3      3           3         0         0        0  0        0  0     1     0  0    0       0  0
number_formats       3  3      3           3         0         0        0  0        0  1     3     0  0    0       0  1
alignment            3  3      3           3         1         1        0  1        1  1     3     1  1    1       1  1
borders              3  3      3           3         0         0        0  0        0  0     1     0  0    0       0  0
dimensions           3  3      3           1         0         0        0  0        0  0     1     0  0    0       0  0
multiple_sheets      3  3      3           3         3         3        3  3        3  3     3     3  3    1       3  3
merged_cells         3  3      3           3         0         0        0  0        0  0     0     0  0    0       0  0
conditional_format   3  3      3           3         0         0        0  0        0  0     0     0  0    0       0  0
data_validation      3  3      3           3         0         0        0  0        0  0     0     0  0    0       0  0
hyperlinks           3  3      3           3         0         0        0  0        0  0     0     0  0    0       0  0
images               3  3      3           0         0         0        0  0        0  0     0     0  0    0       0  0
comments             3  3      3           0         0         0        0  0        0  0     0     0  0    0       0  0
freeze_panes         3  3      3           3         0         0        0  0        0  0     0     0  0    0       0  0
```

## Abstraction Cost Analysis

### Value-only wrappers (pandas vs pyexcel vs tablib vs polars)

All four wrap openpyxl or calamine internally. Key differences:

| Metric | pandas | pyexcel | tablib | polars | Winner |
|--------|--------|---------|--------|--------|--------|
| cell_values read | 1 (errors→NaN) | 3 | 3 | 1 (type coercion) | pyexcel/tablib |
| cell_values write | 3 | 3 | 3 | — | tie |
| formulas read | 0 | 0 | 0 | 0 | tie |
| formulas write | 3 | 3 | 3 | — | tie |
| alignment read | 1 | 1 | 1 | 1 | tie |
| number_formats write | 0 | 1 | 1 | — | pyexcel/tablib |
| Green features (R) | 1/16 | 2/16 | 2/16 | 0/16 | pyexcel/tablib |
| Green features (W) | 3/16 | 3/16 | 3/16 | — | tie |

**Key findings:**
- **pandas** loses error values (`#DIV/0!`, `#N/A`) because DataFrames coerce them to `NaN`
- **polars** loses even more due to columnar type coercion — mixed-type columns become strings, and multi-sheet support scores 1 (not 3) due to API limitations
- **tablib** matches pyexcel exactly — both preserve error values through their cell iterators
- **pyexcel** and **tablib** are the safest value-only abstractions for reads

### openpyxl default vs readonly mode

| Metric | openpyxl (default) | openpyxl-readonly | Difference |
|--------|-------------------|-------------------|------------|
| Green features (R) | 16/16 | 3/16 | -13 |
| Read score | 48/48 | 10/48 | -38 |
| Pass rate | 100% | 24% | -76pp |

**Key finding:** Read-only mode loses ALL formatting metadata (text_formatting, borders, background_colors, number_formats, dimensions, comments, images, hyperlinks, merged_cells, conditional_formatting, data_validation, freeze_panes). It preserves only cell_values, formulas, and multiple_sheets.

### xlsxwriter default vs constant_memory mode

| Metric | xlsxwriter (default) | xlsxwriter-constmem | Difference |
|--------|---------------------|---------------------|------------|
| Green features (W) | 16/16 | 13/16 | -3 |
| Write score | 48/48 | 43/48 | -5 |
| Pass rate | 100% | 94% | -6pp |

**Key finding:** constant_memory mode loses images (not supported), comments (not supported), and dimensions (row-major write order limits control). All formatting features (text, borders, colors, alignment, number_formats, conditional_formatting, data_validation, hyperlinks, freeze_panes, merged_cells) are fully preserved.

## Checklist — Expansion Tasks

- [x] A1: pandas adapter (value-only R+W, abstraction-cost measurement)
- [x] A4: openpyxl readonly mode adapter
- [x] A5: xlsxwriter constant_memory mode adapter
- [x] A6: polars adapter (Rust DataFrame reader)
- [x] A7: tablib adapter
- [ ] A2: odfpy adapter (ODS format)
- [ ] A3: et-xmlfile investigation

## Session Log (append-only)

### 02/08/2026 02:09 PM PST (via pst-timestamp)
- Worked on: A1 (pandas adapter)
- Committed: `46b8a87 feat(adapter): add pandas adapter measuring abstraction cost vs openpyxl`
- Scores: Read 1/3 cell_values (errors→NaN), Write 3/3 cell_values; 1/16 green R, 3/16 green W
- Benchmark: Regenerated full xlsx+xls profiles with all 8 Python adapters
- Decisions: pandas errors-as-NaN is a genuine abstraction cost, not a bug to fix
- Next: A4 (openpyxl readonly mode) or A6 (polars) — both measure optimization variants

### 02/08/2026 — Session 2
- Worked on: A4 (openpyxl-readonly), A5 (xlsxwriter-constmem), A6 (polars), A7 (tablib)
- Commits:
  - `89ec792` build: add polars, fastexcel, tablib dependencies
  - `d39a5a0` feat(adapter): add xlsxwriter constant_memory adapter (A5)
  - `56a71ec` feat(adapter): add openpyxl read-only adapter (A4)
  - `fa16ecb` feat(adapter): add polars adapter (A6)
  - `8829ad9` feat(adapter): add tablib adapter (A7)
  - `3e9aa05` feat(adapter): register all 4 in adapter registry
- Test results: 1067 passed, 39 skipped, 6 xfailed (69 new tests, 0 regressions)
- Scores:
  - openpyxl-readonly: 10/48 R, 3/16 green (loses all formatting in streaming mode)
  - xlsxwriter-constmem: 43/48 W, 13/16 green (loses images/comments/dimensions)
  - polars: 4/48 R, 0/16 green (type coercion cost + limited multi-sheet)
  - tablib: 10/48 R + 12/48 W, 2/16 green R + 3/16 green W (matches pyexcel)
- Key findings: See Abstraction Cost Analysis section above
- Total adapters: 12 Python xlsx + 2 xls + 3 Rust/PyO3 = 17
