# ExcelBench Dashboard

*Profile: xlsx | Generated: 2026-02-16T02:40:41.595881+00:00*

> Combined fidelity and performance view. Fidelity shows correctness;
> throughput shows speed. Use this to find the right library for your needs.

## Library Comparison

| Library | Caps | Green Features | Pass Rate | Best For |
|---------|:----:|:--------------:|:---------:|----------|
| openpyxl | R+W | 18/18 | 100% | Full-fidelity read + write |
| calamine-styled | R | 17/18 | 98% | General use |
| rust_xlsxwriter | W | 17/18 | 98% | General use |
| wolfxl | R+W | 17/18 | 98% | General use |
| xlsxwriter | W | 16/18 | 90% | High-fidelity write-only workflows |
| xlsxwriter-constmem | W | 13/18 | 85% | Large file writes with memory limits |
| xlwt | W | 4/18 | 58% | Legacy .xls file writes |
| openpyxl-readonly | R | 3/18 | 22% | Streaming reads when formatting isn't needed |
| pandas | R+W | 3/18 | 19% | Data analysis pipelines (accept NaN coercion) |
| pyexcel | R+W | 3/18 | 20% | Multi-format compatibility layer |
| tablib | R+W | 3/18 | 20% | Dataset export/import workflows |
| calamine | R | 2/18 | 18% | General use |
| pylightxl | R+W | 2/18 | 18% | Lightweight value extraction |
| python-calamine | R | 1/18 | 16% | Fast bulk value reads |
| polars | R | 0/18 | 14% | High-performance DataFrames (values only) |

## Key Insights

- **Fidelity leaders**: openpyxl (18/18 green features)
- **Abstraction cost**: pandas wraps openpyxl but drops from 18 to 3 green features due to DataFrame coercion
- **Optimization cost**: xlsxwriter constant_memory mode loses 3 green features for lower memory usage

## Best Adapter by Workload Profile

| Workload Size | Best Read Adapter | Best Write Adapter |
|---------------|-------------------|--------------------|
| small | — | — |
| medium | — | — |
| large | — | — |
| small | — | — |
| medium | — | — |
| large | — | — |
