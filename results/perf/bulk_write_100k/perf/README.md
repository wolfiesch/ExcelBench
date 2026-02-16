# ExcelBench Performance Results

*Generated: 2026-02-16T02:49:28.197836+00:00*
*Profile: xlsx*
*Platform: Darwin-arm64*
*Python: 3.12.3*
*Commit: 1e90e78*
*Config: warmup=1 iters=5 iteration_policy=fixed breakdown=False*

## Notes

These numbers measure only the library under test. Write timings do NOT include oracle verification.

## Throughput (derived from p50)

Computed as: op_count * 1000 / p50_wall_ms

**Bulk Write**

| Scenario | op_count | op_unit | openpyxl (R units/s) | openpyxl (W units/s) | pandas (R units/s) | pandas (W units/s) | rust_xlsxwriter (W units/s) | tablib (R units/s) | tablib (W units/s) | wolfxl (R units/s) | wolfxl (W units/s) | xlsxwriter (W units/s) |
|----------|----------|---------|----------------|----------------|----------------|----------------|----------------|----------------|----------------|----------------|----------------|----------------|
| cell_values_100k_bulk_write | 99856 | cells | — | 319.83K | — | 238.32K | 138.89K | — | 255.74K | — | 137.63K | 556.55K |
