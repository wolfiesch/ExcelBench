# ExcelBench Results

*Generated: 2026-02-17 03:48 UTC*
*Profile: xls*
*Excel Version: xlwt*
*Platform: Darwin-arm64*

## Overview

> Condensed view — shows the **best score** across read/write for each library. See [Full Results Matrix](#full-results-matrix) for the complete R/W breakdown.

**Tier 0 — Basic Values**

| Feature | calamine | xlrd |
|---------|:-:|:-:|
| Cell Values | 🟢 | 🟢 |
| Sheets | 🟢 | 🟢 |

**Tier 1 — Formatting**

| Feature | calamine | xlrd |
|---------|:-:|:-:|
| Alignment | 🔴 | 🟢 |
| Dimensions | 🔴 | 🟢 |

## Library Tiers

> Libraries ranked by their best capability (max of read/write green features).

| Tier | Library | Caps | Modify | Green Features | Summary |
|:----:|---------|:----:|:------:|:--------------:|---------|
| **S** | xlrd | R | No | 4/4 | Legacy .xls reader — not applicable to .xlsx |
| **C** | python-calamine | R | No | 2/4 | Fast Rust-backed reader — cell values + sheet names only |

## Score Legend

| Score | Meaning |
|-------|---------|
| 🟢 3 | Complete — all basic and edge cases pass |
| 🟡 2 | Functional — all basic pass, one or more edge cases fail |
| 🟠 1 | Minimal — at least one basic case passes, but not all basic cases |
| 🔴 0 | Unsupported — errors or data loss |
| ➖ | Not applicable |

## Full Results Matrix

**Tier 0 — Basic Values**

| Feature | python-calamine (R) | xlrd (R) |
|---------|------------|------------|
| [cell_values](#cell_values-details) | 🟢 3 | 🟢 3 |
| [multiple_sheets](#multiple_sheets-details) | 🟢 3 | 🟢 3 |

**Tier 1 — Formatting**

| Feature | python-calamine (R) | xlrd (R) |
|---------|------------|------------|
| [alignment](#alignment-details) | 🔴 0 | 🟢 3 |
| [dimensions](#dimensions-details) | 🔴 0 | 🟢 3 |

## Notes

- **alignment**: Known limitation: python-calamine alignment read is limited because its API does not expose style/alignment metadata.

## Statistics

| Library | Mode | Tests | Passed | Failed | Pass Rate | Green Features |
|---------|------|-------|--------|--------|-----------|----------------|
| python-calamine | R | 35 | 22 | 13 | 63% | 2/4 |
| xlrd | R | 35 | 35 | 0 | 100% | 4/4 |

## Libraries Tested

- **python-calamine** v0.6.1 (python) - read; modify: No
- **xlrd** v2.0.2 (python) - read; modify: No

## Diagnostics Summary

| Group | Value | Count |
|-------|-------|-------|
| category | data_mismatch | 13 |
| severity | error | 13 |

### Diagnostic Details

| Feature | Library | Test Case | Operation | Category | Severity | Message |
|---------|---------|-----------|-----------|----------|----------|---------|
| alignment | python-calamine | h_left | read | data_mismatch | error | Expected values did not match actual values: expected={'h_align': 'left'}, actual={} |
| alignment | python-calamine | h_center | read | data_mismatch | error | Expected values did not match actual values: expected={'h_align': 'center'}, actual={} |
| alignment | python-calamine | h_right | read | data_mismatch | error | Expected values did not match actual values: expected={'h_align': 'right'}, actual={} |
| alignment | python-calamine | v_top | read | data_mismatch | error | Expected values did not match actual values: expected={'v_align': 'top'}, actual={} |
| alignment | python-calamine | v_center | read | data_mismatch | error | Expected values did not match actual values: expected={'v_align': 'center'}, actual={} |
| alignment | python-calamine | v_bottom | read | data_mismatch | error | Expected values did not match actual values: expected={'v_align': 'bottom'}, actual={} |
| alignment | python-calamine | wrap_text | read | data_mismatch | error | Expected values did not match actual values: expected={'wrap': True}, actual={} |
| alignment | python-calamine | rotation_45 | read | data_mismatch | error | Expected values did not match actual values: expected={'rotation': 45}, actual={} |
| alignment | python-calamine | indent_2 | read | data_mismatch | error | Expected values did not match actual values: expected={'indent': 2}, actual={} |
| dimensions | python-calamine | row_height_30 | read | data_mismatch | error | Expected values did not match actual values: expected={'row_height': 30}, actual={'row_height': None} |
| dimensions | python-calamine | row_height_45 | read | data_mismatch | error | Expected values did not match actual values: expected={'row_height': 45}, actual={'row_height': None} |
| dimensions | python-calamine | col_width_20 | read | data_mismatch | error | Expected values did not match actual values: expected={'column_width': 20}, actual={'column_width': None} |
| dimensions | python-calamine | col_width_8 | read | data_mismatch | error | Expected values did not match actual values: expected={'column_width': 8}, actual={'column_width': None} |

## Detailed Results

<a id="alignment-details"></a>
### alignment

**python-calamine** — Read: 🔴 0
- Notes: Known limitation: python-calamine alignment read is limited because its API does not expose style/alignment metadata.

| Test | Importance | Read |
|------|-----------|------|
| Align - left | basic | ❌ |
| Align - center | basic | ❌ |
| Align - right | basic | ❌ |
| Align - top | basic | ❌ |
| Align - center | basic | ❌ |
| Align - bottom | basic | ❌ |
| Align - wrap text | basic | ❌ |
| Align - rotation 45 | basic | ❌ |
| Align - indent 2 | basic | ❌ |

**xlrd** — Read: 🟢 3

<a id="cell_values-details"></a>
### cell_values

**python-calamine** — Read: 🟢 3

**xlrd** — Read: 🟢 3

<a id="dimensions-details"></a>
### dimensions

**python-calamine** — Read: 🔴 0

| Test | Importance | Read |
|------|-----------|------|
| Row height - 30 | basic | ❌ |
| Row height - 45 | basic | ❌ |
| Column width - D = 20 | basic | ❌ |
| Column width - E = 8 | basic | ❌ |

**xlrd** — Read: 🟢 3

<a id="multiple_sheets-details"></a>
### multiple_sheets

**python-calamine** — Read: 🟢 3

**xlrd** — Read: 🟢 3

---
*Benchmark version: 0.1.0*