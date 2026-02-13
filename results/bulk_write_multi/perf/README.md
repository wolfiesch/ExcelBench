# ExcelBench Performance Results

*Generated: 2026-02-12T12:42:33.289056+00:00*
*Profile: xlsx*
*Platform: Darwin-arm64*
*Python: 3.12.3*
*Commit: bbe3330*
*Config: warmup=1 iters=5 breakdown=False*

## Notes

These numbers measure only the library under test. Write timings do NOT include oracle verification.

## Throughput (derived from p50)

Computed as: op_count * 1000 / p50_wall_ms

**Bulk Write**

| Scenario | op_count | op_unit | openpyxl (R units/s) | openpyxl (W units/s) | pandas (R units/s) | pandas (W units/s) | tablib (R units/s) | tablib (W units/s) | xlsxwriter (W units/s) |
|----------|----------|---------|----------------|----------------|----------------|----------------|----------------|----------------|----------------|
| cell_values_10k_1000x10_bulk_write | 10000 | cells | — | 307.36K | — | 227.89K | — | 243.13K | 462.73K |
| cell_values_10k_10x1000_bulk_write | 10000 | cells | — | 327.21K | — | 177.73K | — | 261.54K | 471.64K |
| cell_values_10k_bulk_write | 10000 | cells | — | 301.89K | — | 217.37K | — | 250.81K | 465.95K |
| cell_values_10k_sparse_1pct_bulk_write | 100 | cells | — | 43.73K | — | 2.22K | — | 4.12K | 30.70K |
| cell_values_1k_bulk_write | 1000 | cells | — | 182.00K | — | 144.70K | — | 176.18K | 223.22K |
| strings_repeated_10k_bulk_write | 10000 | cells | — | 283.76K | — | 201.75K | — | 231.62K | 515.01K |
| strings_repeated_1k_len256_bulk_write | 1000 | cells | — | 87.78K | — | 79.72K | — | 101.27K | 241.09K |
| strings_unique_10k_bulk_write | 10000 | cells | — | 236.48K | — | 183.32K | — | 213.22K | 308.53K |
| strings_unique_1k_bulk_write | 1000 | cells | — | 175.37K | — | 123.59K | — | 145.52K | 176.95K |
| strings_unique_1k_len256_bulk_write | 1000 | cells | — | 102.79K | — | 93.39K | — | 105.73K | 74.67K |
| strings_unique_1k_len64_bulk_write | 1000 | cells | — | 142.00K | — | 112.11K | — | 129.96K | 143.39K |
