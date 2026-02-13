# ExcelBench Performance Results

*Generated: 2026-02-12T12:42:40.696360+00:00*
*Profile: xlsx*
*Platform: Darwin-arm64*
*Python: 3.12.3*
*Commit: bbe3330*
*Config: warmup=1 iters=5 breakdown=False*

## Notes

These numbers measure only the library under test. Write timings do NOT include oracle verification.

## Throughput (derived from p50)

Computed as: op_count * 1000 / p50_wall_ms

**Per-Cell**

| Scenario | op_count | op_unit | openpyxl (R units/s) | openpyxl (W units/s) | pyexcel (R units/s) | pyexcel (W units/s) | pylightxl (R units/s) | pylightxl (W units/s) | xlsxwriter (W units/s) |
|----------|----------|---------|----------------|----------------|----------------|----------------|----------------|----------------|----------------|
| alignment_1k | 1000 | cells | 119.87K | 54.43K | 204.56K | 172.62K | 190.10K | 273.70K | 76.00K |
| background_colors_1k | 1000 | cells | 120.85K | 45.76K | 197.42K | 137.29K | 189.25K | 256.63K | 72.94K |
| borders_200 | 200 | cells | 72.44K | 17.66K | 104.96K | 79.78K | 117.14K | 179.25K | 33.42K |
| cell_values_10k | 10000 | cells | 269.01K | 252.22K | 61.43K | 305.34K | 255.60K | 299.94K | 303.23K |
| cell_values_1k | 1000 | cells | 218.59K | 182.76K | 116.16K | 221.08K | 197.06K | 280.27K | 188.98K |
| formulas_10k | 10000 | cells | 219.23K | 249.66K | 61.92K | 296.06K | 226.01K | 321.89K | 119.04K |
| formulas_1k | 1000 | cells | 173.69K | 152.17K | 107.90K | 185.56K | 170.66K | 267.75K | 80.38K |
| number_formats_1k | 1000 | cells | 123.72K | 122.19K | 217.18K | 187.87K | 192.04K | 265.82K | 83.79K |

## Run Issues

- alignment_1k / xlsxwriter: Read unsupported
- background_colors_1k / xlsxwriter: Read unsupported
- borders_200 / xlsxwriter: Read unsupported
- cell_values_10k / xlsxwriter: Read unsupported
- cell_values_1k / xlsxwriter: Read unsupported
- formulas_10k / xlsxwriter: Read unsupported
- formulas_1k / xlsxwriter: Read unsupported
- number_formats_1k / xlsxwriter: Read unsupported
