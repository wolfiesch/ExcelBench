# Known Limitations

WolfXL is optimized for high-impact openpyxl-style workflows, not complete openpyxl API parity.

## Scope notes

- Compatibility-focused API subset is prioritized.
- Some advanced or niche spreadsheet behaviors may not be implemented yet.

## Performance claim guardrails

- No claim of universal speedups.
- Benchmark outcomes depend on workload shape and environment.
- Always validate on your own files.

## Integrity guidance

- Use reproducible fixtures/benchmarks for acceptance testing.
- Review output workbooks in Excel for business-critical templates.
