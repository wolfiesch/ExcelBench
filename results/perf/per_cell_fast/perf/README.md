# ExcelBench Performance Results

*Generated: 2026-02-16T02:48:58.315420+00:00*
*Profile: xlsx*
*Platform: Darwin-arm64*
*Python: 3.12.3*
*Commit: 1e90e78*
*Config: warmup=1 iters=5 iteration_policy=fixed breakdown=False*

## Notes

These numbers measure only the library under test. Write timings do NOT include oracle verification.

## Throughput (derived from p50)

Computed as: op_count * 1000 / p50_wall_ms

**Per-Cell**

| Scenario | op_count | op_unit | openpyxl (R units/s) | openpyxl (W units/s) | pyexcel (R units/s) | pyexcel (W units/s) | pylightxl (R units/s) | pylightxl (W units/s) | wolfxl (R units/s) | wolfxl (W units/s) | xlsxwriter (W units/s) |
|----------|----------|---------|----------------|----------------|----------------|----------------|----------------|----------------|----------------|----------------|----------------|
| alignment_1k | 1000 | cells | 120.33K | 55.85K | 207.80K | 179.08K | 195.57K | 276.04K | 99.43K | 67.96K | 81.46K |
| background_colors_1k | 1000 | cells | 120.82K | 46.54K | 190.89K | 163.01K | 186.42K | 264.30K | 86.37K | 67.09K | 76.63K |
| borders_200 | 200 | cells | 77.77K | 17.54K | 109.48K | 82.75K | 113.90K | 207.50K | 45.89K | 33.82K | 34.74K |
| cell_values_10k | 10000 | cells | 267.12K | 257.98K | 62.05K | 322.68K | 257.06K | 306.42K | 108.08K | 95.02K | 273.14K |
| cell_values_1k | 1000 | cells | 226.10K | 193.96K | 114.39K | 217.06K | 188.65K | 295.97K | 100.62K | 85.19K | 218.56K |
| formulas_10k | 10000 | cells | 212.07K | 241.57K | 63.38K | 322.05K | 224.15K | 317.32K | 80.83K | 128.56K | 120.75K |
| formulas_1k | 1000 | cells | 187.84K | 184.34K | 112.97K | 216.83K | 182.39K | 305.71K | 78.64K | 98.53K | 99.67K |
| number_formats_1k | 1000 | cells | 129.78K | 130.14K | 214.96K | 221.22K | 193.62K | 280.16K | 109.36K | 65.18K | 84.20K |

## Run Issues

- alignment_1k / xlsxwriter: Read unsupported
- background_colors_1k / xlsxwriter: Read unsupported
- borders_200 / xlsxwriter: Read unsupported
- cell_values_10k / xlsxwriter: Read unsupported
- cell_values_1k / xlsxwriter: Read unsupported
- formulas_10k / xlsxwriter: Read unsupported
- formulas_1k / xlsxwriter: Read unsupported
- number_formats_1k / xlsxwriter: Read unsupported
