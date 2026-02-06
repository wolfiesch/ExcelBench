# ExcelBench Results

*Generated: 2026-02-06 10:16 UTC*
*Profile: xls*
*Excel Version: xlwt*
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

| Feature | python-calamine (R) | xlrd (R) |
|---------|------------|------------|
| alignment | 🟠 1 | 🟢 3 |
| cell_values | 🟢 3 | 🟢 3 |
| dimensions | 🔴 0 | 🟢 3 |
| multiple_sheets | 🟢 3 | 🟢 3 |

Notes:
- alignment: Known limitation: python-calamine alignment read is limited because its API does not expose style/alignment metadata.

## Libraries Tested

- **python-calamine** v0.6.1 (python) - read
- **xlrd** v2.0.2 (python) - read

## Detailed Results

### alignment

**python-calamine**
- Read: 🟠 1 (1/3)
- Notes: Known limitation: python-calamine alignment read is limited because its API does not expose style/alignment metadata.
- Failed tests (8):
  - h_left (read)
  - h_center (read)
  - h_right (read)
  - v_top (read)
  - v_center (read)
  - ... and 3 more

**xlrd**
- Read: 🟢 3 (3/3)

### cell_values

**python-calamine**
- Read: 🟢 3 (3/3)

**xlrd**
- Read: 🟢 3 (3/3)

### dimensions

**python-calamine**
- Read: 🔴 0 (0/3)
- Failed tests (4):
  - row_height_30 (read)
  - row_height_45 (read)
  - col_width_20 (read)
  - col_width_8 (read)

**xlrd**
- Read: 🟢 3 (3/3)

### multiple_sheets

**python-calamine**
- Read: 🟢 3 (3/3)

**xlrd**
- Read: 🟢 3 (3/3)

---
*Benchmark version: 0.1.0*