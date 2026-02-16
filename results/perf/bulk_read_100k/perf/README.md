# ExcelBench Performance Results

*Generated: 2026-02-16T02:49:09.451830+00:00*
*Profile: xlsx*
*Platform: Darwin-arm64*
*Python: 3.12.3*
*Commit: 1e90e78*
*Config: warmup=1 iters=5 iteration_policy=fixed breakdown=False*

## Notes

These numbers measure only the library under test. Write timings do NOT include oracle verification.

## Throughput (derived from p50)

Computed as: op_count * 1000 / p50_wall_ms

**Bulk Read**

| Scenario | op_count | op_unit | calamine-styled (R units/s) | openpyxl (R units/s) | openpyxl (W units/s) | openpyxl-readonly (R units/s) | pandas (R units/s) | pandas (W units/s) | polars (R units/s) | python-calamine (R units/s) | tablib (R units/s) | tablib (W units/s) | wolfxl (R units/s) | wolfxl (W units/s) |
|----------|----------|---------|----------------|----------------|----------------|----------------|----------------|----------------|----------------|----------------|----------------|----------------|----------------|----------------|
| cell_values_100k_bulk_read | 99856 | cells | 120.53K | 249.42K | — | 397.21K | 414.29K | — | 2.00M | 1.14M | 444.92K | — | 120.76K | — |
