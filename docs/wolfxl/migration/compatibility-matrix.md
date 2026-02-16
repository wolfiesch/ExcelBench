# Compatibility Matrix

Status legend:

- `Supported` - implemented and covered by tests/fixtures
- `Partial` - implemented with caveats
- `Not Yet` - not implemented in WolfXL API surface

## WolfXL API Surface (current)

| Area | Status | Notes |
|---|---|---|
| `load_workbook(path)` | Supported | Read mode |
| `load_workbook(path, modify=True)` | Supported | Modify mode via patcher |
| `Workbook()` | Supported | Write mode |
| `wb.sheetnames` | Supported | List of sheet names |
| `wb.active` | Supported | Returns first sheet |
| `wb["Sheet"]` | Supported | Sheet by name |
| `wb.create_sheet(title)` | Supported | Write mode only |
| `wb.save(path)` | Supported | Write/modify mode |
| `ws["A1"]`, assignment | Supported | Cell access/updates |
| `ws.cell(row, column, value)` | Supported | openpyxl-like API |
| `ws.iter_rows(...)` | Supported | values and cell objects |
| `ws.merge_cells(range)` | Supported | Write mode |
| Font/fill/border/alignment styles | Supported | Via style dataclasses |
| Number format | Supported | `cell.number_format` |
| Full openpyxl API parity | Partial | Focused subset |

## Ecosystem Comparison (high-level)

| Capability | openpyxl | XlsxWriter | python-calamine / fastexcel | WolfXL |
|---|---|---|---|---|
| Read `.xlsx` | Yes | No (write-only) | Yes (reader-focused) | Yes |
| Write `.xlsx` | Yes | Yes | Reader-focused | Yes |
| Modify existing workbook | Yes | No | Not primary scope | Yes |
| openpyxl-style workflow | Native | Different API | Different API | Targeted compatibility |

Notes:

- XlsxWriter explicitly documents write-only scope.
- Read-focused libraries are excellent for ingestion workloads.
- WolfXL emphasizes compatibility + performance + fidelity checks.
