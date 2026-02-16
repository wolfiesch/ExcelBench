# WolfXL

**The fastest openpyxl-compatible Excel library for Python.** Drop-in replacement backed by Rust — up to 5x faster reads and writes with zero code changes.

## The 1-Line Migration

```diff
- from openpyxl import load_workbook, Workbook
- from openpyxl.styles import Font, PatternFill, Alignment, Border
+ from wolfxl import load_workbook, Workbook, Font, PatternFill, Alignment, Border
```

That's it. Your existing code works as-is.

## Install

```bash
pip install wolfxl
```

## Why WolfXL?

openpyxl is the most popular Python Excel library, but it's slow — pure Python XML parsing can't keep up with large files. WolfXL uses the same familiar API but routes everything through Rust:

| Operation | openpyxl | WolfXL | Speedup |
|-----------|----------|--------|---------|
| Read 10M cells (45 MB) | 47.8s | 13.0s | **3.7x** |
| Write 10M cells (51 MB) | 31.8s | 6.7s | **4.8x** |
| Read 1M cells (3 MB) | 4.0s | 1.1s | **3.6x** |
| Write 1M cells (3 MB) | 2.9s | 0.9s | **3.4x** |
| Read 100K cells | 0.42s | 0.11s | **3.8x** |
| Write 100K cells | 0.28s | 0.06s | **4.6x** |

*Benchmarked on Apple M1 Pro, Python 3.12. Full methodology at [ExcelBench](https://excelbench.vercel.app).*

## Quick Start

### Write a styled spreadsheet

```python
from wolfxl import Workbook, Font, PatternFill

wb = Workbook()
ws = wb.active

# Styled header
ws["A1"].value = "Product"
ws["A1"].font = Font(bold=True, size=14, color="FFFFFF")
ws["A1"].fill = PatternFill(fill_type="solid", fgColor="336699")

# Data
ws["A2"].value = "Widget"
ws["B2"].value = 9.99

wb.save("report.xlsx")
```

### Read an existing file

```python
from wolfxl import load_workbook

wb = load_workbook("report.xlsx")
ws = wb[wb.sheetnames[0]]

for row in ws.iter_rows(min_row=1, max_row=10, values_only=False):
    for cell in row:
        print(f"{cell.coordinate}: {cell.value} (bold={cell.font.bold})")

wb.close()
```

### Modify in place (preserves charts, macros, images)

```python
from wolfxl import load_workbook

wb = load_workbook("existing.xlsx", modify=True)
ws = wb[wb.sheetnames[0]]
ws["A1"].value = "Updated!"
wb.save("existing.xlsx")  # Surgical ZIP patch — only changed cells are rewritten
```

## Supported Features

- Cell values (strings, numbers, dates, booleans)
- Formulas (read and write)
- Font styling (bold, italic, underline, color, size)
- Fill colors (solid pattern fills)
- Borders (all styles and colors)
- Number formats
- Alignment (horizontal, vertical, wrap, rotation)
- Multiple sheets
- Merged cells
- Named ranges
- Freeze panes
- Data validation
- Conditional formatting
- Hyperlinks
- Comments
- Tables

## How It Works

WolfXL is a thin Python wrapper over three Rust engines:

- **Read mode** (`load_workbook`): Uses [calamine](https://github.com/tafia/calamine) for fast XLSX parsing with full style extraction
- **Write mode** (`Workbook()`): Uses [rust_xlsxwriter](https://github.com/jmcnamara/rust_xlsxwriter) for fast XLSX generation
- **Modify mode** (`load_workbook(path, modify=True)`): Surgical ZIP patching that rewrites only changed cells, preserving everything else (charts, macros, images, pivot tables)

All three modes expose the same openpyxl-compatible API. Cells use lazy proxies — opening a 10M-cell file is instant; values are fetched from Rust only when accessed.

## License

MIT
