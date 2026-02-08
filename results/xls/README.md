# ExcelBench Results

*Generated: 2026-02-08 22:09 UTC*
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

**Tier 0 — Basic Values**

| Feature | python-calamine (R) | xlrd (R) |
|---------|------------|------------|
| cell_values | 🟢 3 | 🟢 3 |
| multiple_sheets | 🟢 3 | 🟢 3 |

**Tier 1 — Formatting**

| Feature | python-calamine (R) | xlrd (R) |
|---------|------------|------------|
| alignment | 🟠 1 | 🟢 3 |
| dimensions | 🔴 0 | 🟢 3 |

Notes:
- alignment: Known limitation: python-calamine alignment read is limited because its API does not expose style/alignment metadata.

## Statistics

| Library | Mode | Tests | Passed | Failed | Pass Rate | Green Features |
|---------|------|-------|--------|--------|-----------|----------------|
| python-calamine | R | 35 | 23 | 12 | 66% | 2/4 |
| xlrd | R | 35 | 35 | 0 | 100% | 4/4 |

## Libraries Tested

- **python-calamine** v0.6.1 (python) - read
- **xlrd** v2.0.2 (python) - read

## Detailed Results

### alignment

**python-calamine** — Read: 🟠 1
- Notes: Known limitation: python-calamine alignment read is limited because its API does not expose style/alignment metadata.

| Test | Importance | Read |
|------|-----------|------|
| Align - left | basic | ❌ |
| Align - center | basic | ❌ |
| Align - right | basic | ❌ |
| Align - top | basic | ❌ |
| Align - center | basic | ❌ |
| Align - bottom | basic | ✅ |
| Align - wrap text | basic | ❌ |
| Align - rotation 45 | basic | ❌ |
| Align - indent 2 | basic | ❌ |

**xlrd** — Read: 🟢 3

### cell_values

**python-calamine** — Read: 🟢 3

**xlrd** — Read: 🟢 3

### dimensions

**python-calamine** — Read: 🔴 0

| Test | Importance | Read |
|------|-----------|------|
| Row height - 30 | basic | ❌ |
| Row height - 45 | basic | ❌ |
| Column width - D = 20 | basic | ❌ |
| Column width - E = 8 | basic | ❌ |

**xlrd** — Read: 🟢 3

### multiple_sheets

**python-calamine** — Read: 🟢 3

**xlrd** — Read: 🟢 3

---
*Benchmark version: 0.1.0*