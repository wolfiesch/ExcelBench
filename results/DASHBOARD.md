# ExcelBench Dashboard

*Profile: xlsx | Generated: 2026-02-17T03:47:59.231174+00:00*

> Combined fidelity and performance view. Fidelity shows correctness;
> throughput shows speed. Use this to find the right library for your needs.

## Library Comparison

| Library | Caps | Modify | Green Features | Pass Rate | Best For |
|---------|:----:|:------:|:--------------:|:---------:|----------|
| openpyxl | R+W | Rewrite | 16/16 | 100% | Full-fidelity read + write |
| xlsxwriter | W | No | 15/16 | 99% | High-fidelity write-only workflows |
| rust_xlsxwriter | W | No | 14/16 | 97% | General use |
| wolfxl | R+W | Patch | 14/16 | 97% | General use |
| xlsxwriter-constmem | W | No | 12/16 | 93% | Large file writes with memory limits |
| xlwt | W | No | 4/16 | 64% | Legacy .xls file writes |
| openpyxl-readonly | R | No | 3/16 | 23% | Streaming reads when formatting isn't needed |
| pandas | R+W | Rebuild | 3/16 | 20% | Data analysis pipelines (accept NaN coercion) |
| pyexcel | R+W | Rebuild | 3/16 | 21% | Multi-format compatibility layer |
| tablib | R+W | Rebuild | 3/16 | 21% | Dataset export/import workflows |
| pylightxl | R+W | Rebuild | 2/16 | 19% | Lightweight value extraction |
| python-calamine | R | No | 1/16 | 17% | Fast bulk value reads |
| polars | R | No | 0/16 | 15% | High-performance DataFrames (values only) |

## Key Insights

- **Fidelity leaders**: openpyxl (16/16 green features)
- **Abstraction cost**: pandas wraps openpyxl but drops from 16 to 3 green features due to DataFrame coercion
- **Optimization cost**: xlsxwriter constant_memory mode loses 3 green features for lower memory usage

## Best Adapter by Workload Profile

| Workload Size | Best Read Adapter | Best Write Adapter |
|---------------|-------------------|--------------------|
| small | — | — |
| medium | — | — |
| large | — | — |
