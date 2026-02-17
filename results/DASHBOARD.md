# ExcelBench Dashboard

*Profile: xlsx | Generated: 2026-02-17T03:47:59.231174+00:00*

> Combined fidelity and performance view. Fidelity shows correctness;
> throughput shows speed. Use this to find the right library for your needs.

## Library Comparison

| Library | Caps | Green Features | Pass Rate | Best For |
|---------|:----:|:--------------:|:---------:|----------|
| openpyxl | R+W | 16/16 | 100% | Full-fidelity read + write |
| xlsxwriter | W | 15/16 | 99% | High-fidelity write-only workflows |
| rust_xlsxwriter | W | 14/16 | 97% | General use |
| wolfxl | R+W | 14/16 | 97% | General use |
| xlsxwriter-constmem | W | 12/16 | 93% | Large file writes with memory limits |
| xlwt | W | 4/16 | 64% | Legacy .xls file writes |
| openpyxl-readonly | R | 3/16 | 23% | Streaming reads when formatting isn't needed |
| pandas | R+W | 3/16 | 20% | Data analysis pipelines (accept NaN coercion) |
| pyexcel | R+W | 3/16 | 21% | Multi-format compatibility layer |
| tablib | R+W | 3/16 | 21% | Dataset export/import workflows |
| pylightxl | R+W | 2/16 | 19% | Lightweight value extraction |
| python-calamine | R | 1/16 | 17% | Fast bulk value reads |
| polars | R | 0/16 | 15% | High-performance DataFrames (values only) |

## Key Insights

- **Fidelity leaders**: openpyxl (16/16 green features)
- **Abstraction cost**: pandas wraps openpyxl but drops from 16 to 3 green features due to DataFrame coercion
- **Optimization cost**: xlsxwriter constant_memory mode loses 3 green features for lower memory usage

## Best Adapter by Use Case

| Use Case | Best Read Adapter | Best Write Adapter |
|----------|-------------------|--------------------|
| Full fidelity (all features) | openpyxl | openpyxl / xlsxwriter |
| High throughput (values only) | python-calamine | xlsxwriter-constmem |
| Rust-backed (full fidelity) | wolfxl | wolfxl / rust_xlsxwriter |
| Data analysis pipeline | pandas | pandas |
| Memory-constrained writes | — | xlsxwriter-constmem |
| Legacy .xls format | xlrd | xlwt |
