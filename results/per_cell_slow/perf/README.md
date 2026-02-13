# ExcelBench Performance Results

*Generated: 2026-02-12T12:42:48.810135+00:00*
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

| Scenario | op_count | op_unit | python-calamine (R units/s) |
|----------|----------|---------|----------------|
| cell_values_1k | 1000 | cells | 2.70K |
| formulas_1k | 1000 | cells | 2.16K |

## Run Issues

- cell_values_1k / python-calamine: Write unsupported
- formulas_1k / python-calamine: Write unsupported
