# ExcelBench Performance Results

*Generated: 2026-02-12T12:42:24.989131+00:00*
*Profile: xlsx*
*Platform: Darwin-arm64*
*Python: 3.12.3*
*Commit: bbe3330*
*Config: warmup=1 iters=5 breakdown=False*

## Notes

These numbers measure only the library under test. Write timings do NOT include oracle verification.

## Throughput (derived from p50)

Computed as: op_count * 1000 / p50_wall_ms

**Bulk Read**

| Scenario | op_count | op_unit | openpyxl (R units/s) | openpyxl (W units/s) | openpyxl-readonly (R units/s) | pandas (R units/s) | pandas (W units/s) | polars (R units/s) | python-calamine (R units/s) | tablib (R units/s) | tablib (W units/s) |
|----------|----------|---------|----------------|----------------|----------------|----------------|----------------|----------------|----------------|----------------|----------------|
| cell_values_10k_1000x10_bulk_read | 10000 | cells | 324.94K | — | 373.69K | 372.94K | — | 1.92M | 1.47M | 406.17K | — |
| cell_values_10k_10x1000_bulk_read | 10000 | cells | 195.69K | — | 392.67K | 183.66K | — | 662.89K | 1.64M | 427.90K | — |
| cell_values_10k_bulk_read | 10000 | cells | 336.79K | — | 402.07K | 393.16K | — | 1.27M | 1.62M | 456.85K | — |
| cell_values_1k_bulk_read | 1000 | cells | 261.57K | — | 247.61K | 215.04K | — | 654.34K | 1.50M | 269.45K | — |
| formulas_10k_bulk_read | 10000 | cells | 266.11K | — | 307.59K | 335.13K | — | 1.39M | 1.41M | 374.19K | — |
| formulas_1k_bulk_read | 1000 | cells | 194.65K | — | 217.90K | 195.69K | — | 658.27K | 1.35M | 228.09K | — |
| strings_repeated_10k_bulk_read | 10000 | cells | 305.20K | — | 361.03K | 371.57K | — | 1.51M | 1.21M | 431.04K | — |
| strings_repeated_1k_len256_bulk_read | 1000 | cells | 234.36K | — | 240.71K | 194.68K | — | 702.56K | 1.14M | 254.91K | — |
| strings_unique_10k_bulk_read | 10000 | cells | 159.92K | — | 171.22K | 171.48K | — | 1.14M | 996.93K | 188.77K | — |
| strings_unique_1k_bulk_read | 1000 | cells | 139.78K | — | 139.32K | 122.48K | — | 520.61K | 869.12K | 142.25K | — |
| strings_unique_1k_len256_bulk_read | 1000 | cells | 125.35K | — | 119.53K | 114.81K | — | 511.26K | 702.56K | 128.00K | — |
| strings_unique_1k_len64_bulk_read | 1000 | cells | 125.96K | — | 127.36K | 119.95K | — | 581.68K | 889.28K | 139.16K | — |
