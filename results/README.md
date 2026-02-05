# ExcelBench Results

*Generated: 2026-02-05 03:52 UTC*
*Excel Version: openpyxl-generated*
*Platform: Darwin-arm64*

## Score Legend

| Score | Meaning |
|-------|---------|
| 🟢 3 | Complete - full fidelity |
| 🟡 2 | Functional - works for common cases |
| 🟠 1 | Minimal - basic recognition only |
| 🔴 0 | Unsupported - errors or data loss |
| ➖ | Not applicable (library doesn't support this operation) |

## Summary

| Feature | openpyxl (R) | openpyxl (W) | xlsxwriter (W) |
|---------|------------|------------|------------|
| alignment | 🟢 3 | 🟢 3 | 🟡 2 |
| background_colors | 🟢 3 | 🟢 3 | 🟢 3 |
| borders | 🟢 3 | 🟢 3 | 🟡 2 |
| cell_values | 🟢 3 | 🟢 3 | 🟢 3 |
| dimensions | 🟢 3 | 🟢 3 | 🟠 1 |
| formulas | 🟢 3 | 🟢 3 | 🟢 3 |
| multiple_sheets | 🟢 3 | 🟢 3 | 🟢 3 |
| number_formats | 🟢 3 | 🟢 3 | 🟢 3 |
| text_formatting | 🟢 3 | 🟢 3 | 🟢 3 |

## Libraries Tested

- **openpyxl** v3.1.5 (python) - read, write
- **xlsxwriter** v3.2.9 (python) - write

## Detailed Results

### alignment

**openpyxl**
- Read: 🟢 3 (3/3)
- Write: 🟢 3 (3/3)

**xlsxwriter**
- Write: 🟡 2 (2/3)
- Failed tests (1):
  - v_bottom (write)

### background_colors

**openpyxl**
- Read: 🟢 3 (3/3)
- Write: 🟢 3 (3/3)

**xlsxwriter**
- Write: 🟢 3 (3/3)

### borders

**openpyxl**
- Read: 🟢 3 (3/3)
- Write: 🟢 3 (3/3)

**xlsxwriter**
- Write: 🟡 2 (2/3)
- Failed tests (2):
  - diagonal_up (write)
  - diagonal_down (write)

### cell_values

**openpyxl**
- Read: 🟢 3 (3/3)
- Write: 🟢 3 (3/3)

**xlsxwriter**
- Write: 🟢 3 (3/3)

### dimensions

**openpyxl**
- Read: 🟢 3 (3/3)
- Write: 🟢 3 (3/3)

**xlsxwriter**
- Write: 🟠 1 (1/3)
- Failed tests (2):
  - col_width_20 (write)
  - col_width_8 (write)

### formulas

**openpyxl**
- Read: 🟢 3 (3/3)
- Write: 🟢 3 (3/3)

**xlsxwriter**
- Write: 🟢 3 (3/3)

### multiple_sheets

**openpyxl**
- Read: 🟢 3 (3/3)
- Write: 🟢 3 (3/3)

**xlsxwriter**
- Write: 🟢 3 (3/3)

### number_formats

**openpyxl**
- Read: 🟢 3 (3/3)
- Write: 🟢 3 (3/3)

**xlsxwriter**
- Write: 🟢 3 (3/3)

### text_formatting

**openpyxl**
- Read: 🟢 3 (3/3)
- Write: 🟢 3 (3/3)

**xlsxwriter**
- Write: 🟢 3 (3/3)

---
*Benchmark version: 0.1.0*