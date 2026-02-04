# ExcelBench Results

*Generated: 2026-02-04 14:47 UTC*
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
| borders | 🟠 1 | ➖ | ➖ |
| cell_values | 🟠 1 | ➖ | ➖ |
| text_formatting | 🟢 3 | ➖ | ➖ |

## Libraries Tested

- **openpyxl** v3.1.5 (python) - read, write
- **xlsxwriter** v3.2.9 (python) - write

## Detailed Results

### borders

**openpyxl**
- Read: 🟠 1 (1/3)
- Failed tests (4):
  - top_only
  - bottom_only
  - left_only
  - right_only

**xlsxwriter**

### cell_values

**openpyxl**
- Read: 🟠 1 (1/3)
- Failed tests (5):
  - string_empty
  - date_standard
  - error_div0
  - error_na
  - error_value

**xlsxwriter**

### text_formatting

**openpyxl**
- Read: 🟢 3 (3/3)

**xlsxwriter**

---
*Benchmark version: 0.1.0*