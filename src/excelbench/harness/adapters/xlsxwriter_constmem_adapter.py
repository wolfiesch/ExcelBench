"""Adapter for xlsxwriter library with constant_memory mode (write-only).

Identical to XlsxwriterAdapter except the workbook is created with
``{'constant_memory': True}``.  This forces row-by-row streaming writes
(no random cell access) — measuring whether fidelity scores change.

Key constraint: cells must be written top→bottom, left→right.
The deferred buffer is sorted by (row, col) before replay.
"""

from pathlib import Path
from typing import Any

import xlsxwriter

from excelbench.harness.adapters.xlsxwriter_adapter import XlsxwriterAdapter
from excelbench.models import CellFormat, CellType, CellValue, LibraryInfo

WorkbookData = dict[str, Any]


def _get_version() -> str:
    return str(xlsxwriter.__version__)


class XlsxwriterConstmemAdapter(XlsxwriterAdapter):
    """xlsxwriter with ``constant_memory=True`` (streaming writes).

    Inherits all buffering/replay logic from :class:`XlsxwriterAdapter`.
    Only ``info`` and ``save_workbook`` are overridden.
    """

    @property
    def info(self) -> LibraryInfo:
        return LibraryInfo(
            name="xlsxwriter-constmem",
            version=_get_version(),
            language="python",
            capabilities={"write"},
        )

    def save_workbook(self, workbook: WorkbookData, path: Path) -> None:
        """Save with constant_memory=True, replaying ops in row-major order."""
        wb = xlsxwriter.Workbook(str(path), {"constant_memory": True})

        try:
            for sheet_name, operations in workbook["sheets"].items():
                ws = wb.add_worksheet(sheet_name)

                # Row heights / column widths (must be set before writing rows)
                for row_index, height in workbook["row_heights"].get(sheet_name, {}).items():
                    ws.set_row(row_index, height)
                for col_index, width in workbook["col_widths"].get(sheet_name, {}).items():
                    ws.set_column(col_index, col_index, width)

                # Freeze/split panes
                freeze = workbook["freeze"].get(sheet_name)
                if freeze:
                    cfg = freeze.get("freeze", freeze)
                    mode = cfg.get("mode")
                    if mode == "freeze" and cfg.get("top_left_cell"):
                        r, c = self._parse_cell(cfg["top_left_cell"])
                        ws.freeze_panes(r, c)
                    elif mode == "split":
                        ws.split_panes(cfg.get("y_split", 0), cfg.get("x_split", 0))

                # Merged ranges
                for cell_range in workbook["merges"].get(sheet_name, []):
                    ws.merge_range(cell_range, "")

                # Group operations by cell, then sort by (row, col)
                cell_ops: dict[tuple[int, int], dict[str, Any]] = {}
                for op in operations:
                    key = (op["row"], op["col"])
                    if key not in cell_ops:
                        cell_ops[key] = {"value": None, "format": None, "border": None}
                    if op["type"] == "value":
                        cell_ops[key]["value"] = op["value"]
                    elif op["type"] == "format":
                        cell_ops[key]["format"] = op["format"]
                    elif op["type"] == "border":
                        cell_ops[key]["border"] = op["border"]

                # Sort by row then col — required for constant_memory mode
                for (row, col), data in sorted(cell_ops.items()):
                    cell_value: CellValue | None = data["value"]
                    cell_format: CellFormat | None = data["format"]
                    cell_border = data["border"]

                    fmt = None
                    if cell_format or cell_border:
                        fmt = self._create_format(wb, cell_format, cell_border)

                    if cell_value:
                        if cell_value.type in (CellType.DATE, CellType.DATETIME) and fmt is None:
                            default_format = (
                                "yyyy-mm-dd"
                                if cell_value.type == CellType.DATE
                                else "yyyy-mm-dd hh:mm:ss"
                            )
                            fmt = self._create_format(
                                wb, CellFormat(number_format=default_format), None
                            )
                        self._write_typed_cell(ws, wb, row, col, cell_value, fmt)
                    elif fmt:
                        ws.write_blank(row, col, None, fmt)

                # Conditional formats (constant_memory still supports these)
                for rule in workbook["conditional_formats"].get(sheet_name, []):
                    self._apply_conditional_format(ws, wb, rule)

                # Data validations
                for validation in workbook["data_validations"].get(sheet_name, []):
                    self._apply_data_validation(ws, validation)

                # Hyperlinks
                for link in workbook["hyperlinks"].get(sheet_name, []):
                    self._apply_hyperlink(ws, link)

                # NOTE: Images and comments are NOT supported in constant_memory mode.
                # xlsxwriter docs: "insert_image() and write_comment() are not
                # supported with constant_memory mode."

        finally:
            wb.close()

    # -- Helper methods extracted for reuse --

    @staticmethod
    def _write_typed_cell(
        ws: Any,
        wb: Any,
        row: int,
        col: int,
        cell_value: CellValue,
        fmt: Any,
    ) -> None:
        from datetime import date as _date
        from datetime import datetime as _datetime

        if cell_value.type == CellType.BLANK:
            ws.write_blank(row, col, None, fmt)
        elif cell_value.type == CellType.FORMULA:
            ws.write_formula(row, col, cell_value.formula or cell_value.value, fmt)
        elif cell_value.type == CellType.BOOLEAN:
            ws.write_boolean(row, col, cell_value.value, fmt)
        elif cell_value.type == CellType.NUMBER:
            ws.write_number(row, col, cell_value.value, fmt)
        elif cell_value.type == CellType.DATE:
            dt_value = cell_value.value
            if isinstance(dt_value, _date) and not isinstance(dt_value, _datetime):
                dt_value = _datetime.combine(dt_value, _datetime.min.time())
            ws.write_datetime(row, col, dt_value, fmt)
        elif cell_value.type == CellType.DATETIME:
            ws.write_datetime(row, col, cell_value.value, fmt)
        elif cell_value.type == CellType.ERROR:
            error_formulas = {
                "#DIV/0!": "=1/0",
                "#N/A": "=NA()",
                "#VALUE!": '="text"+1',
            }
            fallback = f'=ERROR("{cell_value.value}")'
            formula = error_formulas.get(cell_value.value, fallback)
            ws.write_formula(row, col, formula, fmt)
        else:
            ws.write_string(row, col, str(cell_value.value), fmt)

    def _apply_conditional_format(self, ws: Any, wb: Any, rule: dict[str, Any]) -> None:
        cf = rule.get("cf_rule", rule)
        rng = cf.get("range")
        rule_type = cf.get("rule_type")
        operator = cf.get("operator")
        formula = cf.get("formula")
        fmt = cf.get("format") or {}
        stop_if_true = cf.get("stop_if_true")

        options: dict[str, Any] = {}
        if rule_type in ("cellIs", "cellIsRule"):
            op_map = {
                "greaterThan": ">",
                "lessThan": "<",
                "between": "between",
                "equal": "==",
                "notEqual": "!=",
                "greaterThanOrEqual": ">=",
                "lessThanOrEqual": "<=",
            }
            options["type"] = "cell"
            options["criteria"] = op_map.get(operator, operator)
            options["value"] = formula
        elif rule_type in ("expression", "formula"):
            options["type"] = "formula"
            criteria = formula.lstrip("=") if formula else formula
            options["criteria"] = criteria
        elif rule_type == "colorScale":
            options["type"] = "3_color_scale"
        elif rule_type == "dataBar":
            options["type"] = "data_bar"

        if stop_if_true:
            options["stop_if_true"] = True

        if fmt.get("bg_color"):
            options["format"] = wb.add_format({"fg_color": fmt["bg_color"], "pattern": 1})
        if options and rng:
            ws.conditional_format(rng, options)

    @staticmethod
    def _apply_data_validation(ws: Any, validation: dict[str, Any]) -> None:
        v = validation.get("validation", validation)
        cell_range = v.get("range")
        vtype = v.get("validation_type")
        vop = v.get("operator")
        dv_options: dict[str, Any] = {}
        type_map = {
            "list": "list",
            "whole": "integer",
            "custom": "custom",
            "decimal": "decimal",
            "date": "date",
            "time": "time",
            "textLength": "length",
        }
        dv_options["validate"] = type_map.get(vtype, vtype)
        if vop:
            dv_options["criteria"] = vop
        if v.get("formula1"):
            if dv_options["validate"] == "list":
                source = v.get("formula1")
                if isinstance(source, str) and source.startswith('"') and source.endswith('"'):
                    source = source[1:-1]
                dv_options["source"] = source
            else:
                dv_options["value"] = v.get("formula1")
        if v.get("formula2"):
            dv_options["maximum"] = v.get("formula2")
        if v.get("allow_blank") is not None:
            dv_options["ignore_blank"] = bool(v.get("allow_blank"))
        if v.get("prompt_title"):
            dv_options["input_title"] = v.get("prompt_title")
        if v.get("prompt"):
            dv_options["input_message"] = v.get("prompt")
        if v.get("error_title"):
            dv_options["error_title"] = v.get("error_title")
        if v.get("error"):
            dv_options["error_message"] = v.get("error")
        if cell_range and dv_options:
            ws.data_validation(cell_range, dv_options)

    def _apply_hyperlink(self, ws: Any, link: dict[str, Any]) -> None:
        data = link.get("hyperlink", link)
        cell = data.get("cell")
        target = data.get("target")
        display = data.get("display")
        tooltip = data.get("tooltip")
        internal = data.get("internal")
        if not cell or not target:
            return
        url = target
        if internal:
            url = f"internal:{str(target).lstrip('#')}"
        r, c = self._parse_cell(cell)
        url_opts: dict[str, Any] = {}
        if tooltip:
            url_opts["tip"] = tooltip
        ws.write_url(r, c, url, string=display, **url_opts)
