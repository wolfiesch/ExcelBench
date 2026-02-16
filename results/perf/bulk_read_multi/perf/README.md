# ExcelBench Performance Results

*Generated: 2026-02-16T02:48:27.823228+00:00*
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
| cell_values_10k_1000x10_bulk_read | 10000 | cells | 119.37K | 319.35K | — | 371.81K | 384.09K | — | 2.01M | 1.46M | 420.95K | — | 120.48K | — |
| cell_values_10k_10x1000_bulk_read | 10000 | cells | 128.04K | 347.12K | — | 404.14K | 275.55K | — | 727.25K | 1.61M | 452.29K | — | 127.55K | — |
| cell_values_10k_bulk_read | 10000 | cells | 124.34K | 333.27K | — | 400.49K | 394.01K | — | 1.22M | 1.57M | 462.86K | — | 126.16K | — |
| cell_values_1k_bulk_read | 1000 | cells | 116.80K | 269.59K | — | 250.28K | 218.82K | — | 573.79K | 1.48M | 266.78K | — | 115.30K | — |
| formulas_10k_bulk_read | 10000 | cells | 90.77K | 270.11K | — | 313.71K | 343.52K | — | 1.44M | 1.39M | 372.82K | — | 90.98K | — |
| formulas_1k_bulk_read | 1000 | cells | 83.77K | 209.23K | — | 206.93K | 185.97K | — | 681.82K | 1.31M | 232.56K | — | 85.92K | — |
| strings_repeated_10k_bulk_read | 10000 | cells | 106.88K | 296.77K | — | 358.31K | 374.16K | — | 1.57M | 1.19M | 436.47K | — | 107.79K | — |
| strings_repeated_1k_len256_bulk_read | 1000 | cells | 98.39K | 240.46K | — | 241.76K | 189.25K | — | 613.91K | 1.08M | 234.66K | — | 100.05K | — |
| strings_unique_10k_bulk_read | 10000 | cells | 83.11K | 158.66K | — | 172.83K | 171.74K | — | 1.15M | 1.00M | 185.27K | — | 83.40K | — |
| strings_unique_1k_bulk_read | 1000 | cells | 82.13K | 138.75K | — | 135.73K | 128.39K | — | 684.44K | 992.35K | 151.08K | — | 81.07K | — |
| strings_unique_1k_len256_bulk_read | 1000 | cells | 64.33K | 120.49K | — | 121.92K | 112.41K | — | 511.90K | 661.16K | 128.12K | — | 63.25K | — |
| strings_unique_1k_len64_bulk_read | 1000 | cells | 75.31K | 136.29K | — | 134.37K | 114.13K | — | 575.91K | 909.12K | 140.79K | — | 75.84K | — |
