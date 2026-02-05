# Excel-Generated Fixtures

This folder is intended to contain **canonical Excel-generated fixtures** for CI and benchmarking.

To generate:
```bash
uv run excelbench generate --output fixtures/excel
```

Notes:
- Requires Excel installed (xlwings automation).
- These fixtures are the source of truth for CI runs.
- `test_files/` remains local scratch and is gitignored.
