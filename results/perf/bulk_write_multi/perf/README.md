# ExcelBench Performance Results

*Generated: 2026-02-16T02:48:44.102949+00:00*
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
| cell_values_10k_1000x10_bulk_write | 10000 | cells | — | 333.46K | — | 246.40K | 111.35K | — | 258.25K | — | 110.38K | 487.29K |
| cell_values_10k_10x1000_bulk_write | 10000 | cells | — | 337.88K | — | 186.85K | 107.42K | — | 271.47K | — | 107.94K | 519.80K |
| cell_values_10k_bulk_write | 10000 | cells | — | 327.77K | — | 243.81K | 108.69K | — | 267.69K | — | 108.94K | 500.44K |
| cell_values_10k_sparse_1pct_bulk_write | 100 | cells | — | 37.49K | — | 2.69K | 19.49K | — | 4.31K | — | 20.25K | 36.50K |
| cell_values_1k_bulk_write | 1000 | cells | — | 213.97K | — | 161.79K | 98.28K | — | 193.35K | — | 97.44K | 258.55K |
| strings_repeated_10k_bulk_write | 10000 | cells | — | 259.05K | — | 191.42K | 144.76K | — | 230.97K | — | 144.07K | 525.82K |
| strings_repeated_1k_len256_bulk_write | 1000 | cells | — | 130.11K | — | 105.83K | 94.91K | — | 118.20K | — | 93.34K | 244.13K |
| strings_unique_10k_bulk_write | 10000 | cells | — | 271.29K | — | 191.73K | 75.32K | — | 215.90K | — | 77.09K | 327.25K |
| strings_unique_1k_bulk_write | 1000 | cells | — | 142.20K | — | 128.25K | 71.64K | — | 159.54K | — | 70.88K | 184.05K |
| strings_unique_1k_len256_bulk_write | 1000 | cells | — | 129.06K | — | 95.53K | 34.09K | — | 112.61K | — | 34.40K | 111.90K |
| strings_unique_1k_len64_bulk_write | 1000 | cells | — | 155.21K | — | 117.56K | 51.90K | — | 139.68K | — | 49.93K | 179.57K |
