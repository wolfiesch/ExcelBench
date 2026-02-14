# API Reference

## `UmyaBook`

The main workbook class. Wraps an umya-spreadsheet `Spreadsheet` object.

!!! note "Thread Safety"
    `UmyaBook` is **not thread-safe** (`unsendable` in PyO3). Do not share instances
    across threads. Create separate instances per thread if needed.

---

### File I/O

#### `UmyaBook()`

Create a new, empty workbook with no sheets.

```python
book = UmyaBook()
```

#### `UmyaBook.open(path: str) -> UmyaBook`

Open an existing `.xlsx` file for reading and modification.

```python
book = UmyaBook.open("report.xlsx")
```

**Raises**: `IOError` if the file cannot be read or is not a valid `.xlsx`.

#### `save(path: str) -> None`

Save the workbook to a file. Can only be called **once** per instance.

```python
book.save("output.xlsx")
```

**Raises**: `ValueError` if `save()` has already been called on this instance.

!!! warning "Single-save limitation"
    The current implementation consumes the internal spreadsheet on save.
    To write multiple copies, create separate `UmyaBook` instances.

---

### Sheet Management

#### `sheet_names() -> list[str]`

Return the names of all sheets in the workbook.

```python
names = book.sheet_names()  # ["Sheet1", "Summary"]
```

#### `add_sheet(name: str) -> None`

Add a new empty sheet with the given name.

```python
book.add_sheet("Data")
```

**Raises**: `ValueError` if a sheet with that name already exists.

---

### Cell Values

#### `read_cell_value(sheet: str, cell: str) -> dict`

Read the value of a cell. Returns a dict with `type` and `value` keys.

```python
book.read_cell_value("Sheet1", "A1")
# {"type": "string", "value": "Hello"}

book.read_cell_value("Sheet1", "B2")
# {"type": "number", "value": 42.0}

book.read_cell_value("Sheet1", "C3")
# {"type": "blank"}
```

**Supported types**: `string`, `number`, `boolean`, `formula`, `error`, `date`, `datetime`, `blank`.

#### `write_cell_value(sheet: str, cell: str, payload: dict) -> None`

Write a value to a cell. The payload dict must include a `type` key.

```python
# String
book.write_cell_value("Sheet1", "A1", {"type": "string", "value": "Hello"})

# Number
book.write_cell_value("Sheet1", "B1", {"type": "number", "value": 3.14})

# Boolean
book.write_cell_value("Sheet1", "C1", {"type": "boolean", "value": True})

# Formula
book.write_cell_value("Sheet1", "D1", {"type": "formula", "formula": "=SUM(B1:B10)"})

# Date (ISO format, stored as Excel serial number)
book.write_cell_value("Sheet1", "E1", {"type": "date", "value": "2026-01-15"})

# Datetime
book.write_cell_value("Sheet1", "F1", {"type": "datetime", "value": "2026-01-15T14:30:00"})

# Error
book.write_cell_value("Sheet1", "G1", {"type": "error", "value": "#DIV/0!"})
```

---

### Cell Formatting

#### `read_cell_format(sheet: str, cell: str) -> dict`

Read formatting properties of a cell. Returns a dict with only non-default properties.

```python
book.read_cell_format("Sheet1", "A1")
# {"bold": True, "font_size": 14.0, "font_name": "Arial", "bg_color": "#FFFF00"}
```

**Possible keys**: `bold`, `italic`, `underline`, `strikethrough`, `font_name`, `font_size`,
`font_color`, `bg_color`, `number_format`, `h_align`, `v_align`, `wrap`, `rotation`.

#### `write_cell_format(sheet: str, cell: str, format_dict: dict) -> None`

Apply formatting to a cell. Only specified keys are changed; others are left unchanged.

```python
book.write_cell_format("Sheet1", "A1", {
    "bold": True,
    "italic": False,
    "font_name": "Calibri",
    "font_size": 12.0,
    "font_color": "#FF0000",
    "bg_color": "#FFFF00",
    "number_format": "#,##0.00",
    "h_align": "center",
    "v_align": "top",
    "wrap": True,
    "rotation": 45,
})
```

---

### Borders

#### `read_cell_border(sheet: str, cell: str) -> dict`

Read border information. Returns a dict with edge keys.

```python
book.read_cell_border("Sheet1", "A1")
# {"top": {"style": "thin", "color": "#000000"}, "bottom": {"style": "double", "color": "#0000FF"}}
```

**Edge keys**: `top`, `bottom`, `left`, `right`, `diagonal_up`, `diagonal_down`.

**Border styles**: `thin`, `medium`, `thick`, `double`, `dashed`, `dotted`, `hair`,
`mediumDashed`, `dashDot`, `mediumDashDot`, `dashDotDot`, `mediumDashDotDot`, `slantDashDot`.

#### `write_cell_border(sheet: str, cell: str, border_dict: dict) -> None`

Apply borders to a cell.

```python
book.write_cell_border("Sheet1", "A1", {
    "top": {"style": "thin", "color": "#000000"},
    "bottom": {"style": "medium", "color": "#FF0000"},
    "left": {"style": "dashed", "color": "#00FF00"},
    "right": {"style": "double", "color": "#0000FF"},
})
```

---

### Dimensions

#### `read_row_height(sheet: str, row: int) -> float | None`

Read the height of a row (0-indexed). Returns `None` if not explicitly set.

```python
book.read_row_height("Sheet1", 0)  # 30.0 or None
```

#### `read_column_width(sheet: str, col: str) -> float | None`

Read the width of a column by letter. Returns `None` if not explicitly set.

```python
book.read_column_width("Sheet1", "A")  # 15.0 or None
```

#### `set_row_height(sheet: str, row: int, height: float) -> None`

Set the height of a row (0-indexed).

```python
book.set_row_height("Sheet1", 0, 30.0)
```

#### `set_column_width(sheet: str, col: str, width: float) -> None`

Set the width of a column by letter.

```python
book.set_column_width("Sheet1", "A", 25.0)
```

---

### Merged Cells

#### `read_merged_ranges(sheet: str) -> list[str]`

Read merged ranges as A1 range strings.

```python
book.read_merged_ranges("Sheet1")
# ["A1:D1", "B3:B5"]
```

#### `merge_cells(sheet: str, range: str) -> None`

Merge a range (e.g. `"A1:D1"`).

```python
book.merge_cells("Sheet1", "A1:D1")
```

---

### Comments

#### `read_comments(sheet: str) -> list[dict]`

```python
book.read_comments("Sheet1")
# [{"cell": "A1", "text": "...", "author": "...", "threaded": False}, ...]
```

#### `add_comment(sheet: str, comment_dict: dict) -> None`

Required keys: `cell`. Optional: `text`, `author`.

```python
book.add_comment("Sheet1", {"cell": "A1", "text": "Note", "author": "Me"})
```

---

### Hyperlinks

#### `read_hyperlinks(sheet: str) -> list[dict]`

```python
book.read_hyperlinks("Sheet1")
# [{"cell": "A1", "target": "...", "display": "...", "tooltip": None, "internal": False}, ...]
```

#### `add_hyperlink(sheet: str, link_dict: dict) -> None`

Required keys: `cell`, `target`. Optional: `display`, `tooltip`, `internal`.

```python
book.add_hyperlink("Sheet1", {"cell": "A1", "target": "https://example.com", "display": "Example"})
```

---

### Freeze Panes

#### `read_freeze_panes(sheet: str) -> dict`

Returns an empty dict if no freeze/split is set.

```python
book.read_freeze_panes("Sheet1")
# {"mode": "freeze", "top_left_cell": "B2"}
```

#### `set_freeze_panes(sheet: str, settings: dict) -> None`

Canonical keys:
- `mode`: `"freeze"` or `"split"`
- `top_left_cell`: A1 cell coordinate, e.g. `"B2"`

Convenience aliases (freeze only):
- `row`, `column`: split position as counts of frozen rows/cols.

```python
book.set_freeze_panes("Sheet1", {"mode": "freeze", "top_left_cell": "B2"})
book.set_freeze_panes("Sheet1", {"row": 1, "column": 1})  # alias for B2
```

---

### Images

#### `read_images(sheet: str) -> list[dict]`

Currently exposes anchor metadata only (not image bytes).

```python
book.read_images("Sheet1")
# [{"cell": "C3", "anchor": "oneCell", "offset": [0, 0], "path": None, "alt_text": None}, ...]
```

#### `add_image(sheet: str, image_dict: dict) -> None`

Required keys: `path`, `cell`.

```python
book.add_image("Sheet1", {"path": "logo.png", "cell": "C3"})
```

---

### Data Validation

#### `read_data_validations(sheet: str) -> list[dict]`

```python
book.read_data_validations("Sheet1")
# [{"range": "B2:B10", "validation_type": "list", "formula1": "...", ...}, ...]
```

#### `add_data_validation(sheet: str, validation_dict: dict) -> None`

Key fields:
- `range` (A1 range string)
- `validation_type` (e.g. `list`, `whole`, `date`, ...)
- optional: `operator`, `formula1`, `formula2`, `allow_blank`, `show_input`, `show_error`, etc.

Aliases:
- `type` -> `validation_type`
- `ranges` (list[str]) -> `range` (only when length is 1)
- `error_message` -> `error`

```python
book.add_data_validation("Sheet1", {"range": "B2:B10", "validation_type": "list", "formula1": "A,B,C"})
```

---

### Conditional Formatting

#### `read_conditional_formats(sheet: str) -> list[dict]`

```python
book.read_conditional_formats("Sheet1")
# [{"range": "A1:A10", "rule_type": "cellIs", "operator": "greaterThan", "formula": "5", ...}, ...]
```

#### `add_conditional_format(sheet: str, rule_dict: dict) -> None`

Key fields:
- `range` (A1 range string)
- `rule_type` (e.g. `cellIs`, `expression`, ...)
- optional: `operator`, `formula`, `priority`, `stop_if_true`
- optional: `format`: `{"bg_color": "#RRGGBB", "font_color": "#RRGGBB"}`

Aliases:
- `type` -> `rule_type`
- `ranges` (list[str]) -> `range` (only when length is 1)

```python
book.add_conditional_format("Sheet1", {
    "range": "A1:A10",
    "rule_type": "cellIs",
    "operator": "greaterThan",
    "formula": "5",
    "format": {"bg_color": "#C6EFCE"},
})
```

---

### Named Ranges

#### `read_named_ranges(sheet: str) -> list[dict]`

```python
book.read_named_ranges("Sheet1")
# [{"name": "MyName", "scope": "workbook", "refers_to": "Sheet1!A1:B2"}, ...]
```

#### `add_named_range(sheet: str, nr_dict: dict) -> None`

Required keys: `name`, `refers_to`. Optional: `scope` (`workbook` or `sheet`).

```python
book.add_named_range("Sheet1", {"name": "MyName", "refers_to": "Sheet1!A1:B2", "scope": "workbook"})
```

---

### Tables

#### `read_tables(sheet: str) -> list[dict]`

```python
book.read_tables("Sheet1")
# [{"name": "...", "ref": "A1:D10", "columns": [...], "style": None, "autofilter": False, ...}, ...]
```

#### `add_table(sheet: str, table_dict: dict) -> None`

Required keys: `name`, `ref` (e.g. `"A1:D10"`). Optional: `columns`, `style`, `totals_row`, `autofilter`.

```python
book.add_table("Sheet1", {"name": "Table1", "ref": "A1:D10", "columns": ["A", "B", "C", "D"]})
```

---

### Auto Filter

#### `get_auto_filter(sheet: str) -> str | None`

```python
book.get_auto_filter("Sheet1")
# "A1:D10" or None
```

#### `set_auto_filter(sheet: str, range: str) -> None`

```python
book.set_auto_filter("Sheet1", "A1:D10")
```

#### `remove_auto_filter(sheet: str) -> None`

```python
book.remove_auto_filter("Sheet1")
```

#### `has_auto_filter(sheet: str) -> bool`

```python
book.has_auto_filter("Sheet1")
# True / False
```
