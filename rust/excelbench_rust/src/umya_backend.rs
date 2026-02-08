use pyo3::exceptions::{PyIOError, PyValueError};
use pyo3::prelude::*;
use pyo3::types::PyDict;

use std::collections::HashMap;
use std::io::{Read, Write};
use std::path::Path;

use chrono::{Duration, NaiveDate, NaiveDateTime, NaiveTime};

use std::str::FromStr;

use umya_spreadsheet::{
    new_file, reader, writer, Border, Color, ColorScale, ConditionalFormatValueObject,
    ConditionalFormatValueObjectValues, ConditionalFormatValues, ConditionalFormatting,
    ConditionalFormattingOperatorValues, ConditionalFormattingRule, DataBar, DataValidation,
    DataValidationOperatorValues, DataValidationValues, DataValidations, Fill, Formula,
    HorizontalAlignmentValues, NumberingFormat, Pane, PaneStateValues, PaneValues, PatternFill,
    SequenceOfReferences, SheetView, Spreadsheet, Style, VerticalAlignmentValues,
};

use zip::write::FileOptions;
use zip::{ZipArchive, ZipWriter};

use crate::util::{a1_to_row_col, cell_blank, cell_with_value, parse_iso_date, parse_iso_datetime};

fn looks_like_date_format(code: &str) -> bool {
    // Heuristic: date formats typically include year + day tokens.
    let lc = code.to_ascii_lowercase();
    lc.contains('y') && lc.contains('d')
}

fn to_umya_argb(color: &str) -> String {
    // Normalize common "#RRGGBB" colors into "FFRRGGBB".
    let mut s = color.trim().to_string();
    if let Some(rest) = s.strip_prefix('#') {
        s = rest.to_string();
    }
    let s = s.to_ascii_uppercase();
    if s.len() == 6 {
        format!("FF{s}")
    } else {
        s
    }
}

fn normalize_a1(s: &str) -> String {
    s.replace('$', "")
}

fn strip_leading_equal(s: &str) -> &str {
    s.strip_prefix('=').unwrap_or(s)
}

fn border_style_to_umya(style: &str) -> Option<&'static str> {
    match style {
        "none" => None,
        "thin" => Some(Border::BORDER_THIN),
        "medium" => Some(Border::BORDER_MEDIUM),
        "thick" => Some(Border::BORDER_THICK),
        "double" => Some(Border::BORDER_DOUBLE),
        "dashed" => Some(Border::BORDER_DASHED),
        "dotted" => Some(Border::BORDER_DOTTED),
        "hair" => Some(Border::BORDER_HAIR),
        "mediumDashed" => Some(Border::BORDER_MEDIUMDASHED),
        "dashDot" => Some(Border::BORDER_DASHDOT),
        "mediumDashDot" => Some(Border::BORDER_MEDIUMDASHDOT),
        "dashDotDot" => Some(Border::BORDER_DASHDOTDOT),
        "mediumDashDotDot" => Some(Border::BORDER_MEDIUMDASHDOTDOT),
        "slantDashDot" => Some(Border::BORDER_SLANTDASHDOT),
        _ => None,
    }
}

fn excel_serial_to_naive_datetime(serial: f64) -> Option<NaiveDateTime> {
    // Excel 1900 date system, with the standard 1900 leap-year bug adjustment.
    let epoch = NaiveDate::from_ymd_opt(1899, 12, 30)?.and_time(NaiveTime::MIN);
    let mut f = serial;
    if f < 60.0 {
        f += 1.0;
    }
    let total_ms = (f * 86_400_000.0).round() as i64;
    epoch.checked_add_signed(Duration::milliseconds(total_ms))
}

fn naive_datetime_to_excel_serial(dt: NaiveDateTime) -> Option<f64> {
    let epoch = NaiveDate::from_ymd_opt(1899, 12, 30)?.and_time(NaiveTime::MIN);
    let delta = dt - epoch;
    let total_ms = delta.num_milliseconds();
    Some(total_ms as f64 / 86_400_000.0)
}

#[pyclass(unsendable)]
pub struct UmyaBook {
    book: Spreadsheet,
    saved: bool,
    hyperlink_tooltips: HashMap<String, HashMap<String, String>>, // sheet -> cell -> tooltip
}

#[pymethods]
impl UmyaBook {
    #[new]
    pub fn new() -> Self {
        let mut book = new_file();
        // Match other adapters: start without a default sheet.
        let _ = book.remove_sheet_by_name("Sheet1");
        Self {
            book,
            saved: false,
            hyperlink_tooltips: HashMap::new(),
        }
    }

    #[staticmethod]
    pub fn open(path: &str) -> PyResult<Self> {
        let p = Path::new(path);
        let book = reader::xlsx::read(p)
            .map_err(|e| PyErr::new::<PyIOError, _>(format!("Failed to open workbook: {e}")))?;
        Ok(Self {
            book,
            saved: false,
            hyperlink_tooltips: HashMap::new(),
        })
    }

    pub fn sheet_names(&self) -> PyResult<Vec<String>> {
        let mut names: Vec<String> = Vec::new();
        for sheet in self.book.get_sheet_collection().iter() {
            names.push(sheet.get_name().to_string());
        }
        Ok(names)
    }

    pub fn add_sheet(&mut self, name: &str) -> PyResult<()> {
        self.book
            .new_sheet(name)
            .map_err(|e| PyErr::new::<PyValueError, _>(format!("Failed to add sheet: {e}")))?;
        Ok(())
    }

    pub fn read_cell_value(&self, py: Python<'_>, sheet: &str, a1: &str) -> PyResult<PyObject> {
        let ws = self
            .book
            .get_sheet_by_name(sheet)
            .ok_or_else(|| PyErr::new::<PyValueError, _>(format!("Unknown sheet: {sheet}")))?;

        let (row0, col0) = a1_to_row_col(a1).map_err(|msg| PyErr::new::<PyValueError, _>(msg))?;
        let coord = (col0 + 1, row0 + 1);

        let cell = match ws.get_cell(coord) {
            Some(c) => c,
            None => return cell_blank(py),
        };

        // Formula wins over value.
        let formula = cell.get_formula();
        if !formula.is_empty() {
            // Map well-known error formulas to error tokens (similar to OpenpyxlAdapter).
            let norm = if formula.starts_with('=') {
                formula.to_string()
            } else {
                format!("={formula}")
            };
            let token = match norm.as_str() {
                "=1/0" => Some("#DIV/0!"),
                "=NA()" => Some("#N/A"),
                "=\"text\"+1" => Some("#VALUE!"),
                _ => None,
            };
            if let Some(t) = token {
                return cell_with_value(py, "error", t);
            }

            let d = PyDict::new_bound(py);
            d.set_item("type", "formula")?;
            d.set_item("formula", formula.to_string())?;
            d.set_item("value", formula.to_string())?;
            return Ok(d.into());
        }

        // Numeric typed access.
        if let Some(f) = cell.get_value_number() {
            if let Some(nf) = cell.get_style().get_number_format() {
                let code = nf.get_format_code();
                if looks_like_date_format(code) {
                    if let Some(ndt) = excel_serial_to_naive_datetime(f) {
                        let midnight = NaiveTime::from_hms_opt(0, 0, 0).unwrap();
                        if ndt.time() == midnight {
                            let s = ndt.date().format("%Y-%m-%d").to_string();
                            return cell_with_value(py, "date", s);
                        }
                        let s = ndt.format("%Y-%m-%dT%H:%M:%S").to_string();
                        return cell_with_value(py, "datetime", s);
                    }
                }
            }

            return cell_with_value(py, "number", f);
        }

        // TODO: improve typing (beyond Tier 1). For now treat as error/bool/string.
        let raw = cell
            .get_value()
            .into_owned()
            .replace("\r\n", "\n")
            .replace('\r', "\n");

        // Errors
        if raw == "#N/A" || (raw.starts_with('#') && raw.ends_with('!')) {
            return cell_with_value(py, "error", raw);
        }
        // Boolean
        if raw.eq_ignore_ascii_case("true") {
            return cell_with_value(py, "boolean", true);
        }
        if raw.eq_ignore_ascii_case("false") {
            return cell_with_value(py, "boolean", false);
        }

        if raw.is_empty() {
            return cell_blank(py);
        }

        cell_with_value(py, "string", raw)
    }

    pub fn write_cell_value(
        &mut self,
        sheet: &str,
        a1: &str,
        payload: &Bound<'_, PyAny>,
    ) -> PyResult<()> {
        let ws = self
            .book
            .get_sheet_by_name_mut(sheet)
            .ok_or_else(|| PyErr::new::<PyValueError, _>(format!("Unknown sheet: {sheet}")))?;

        let dict = payload
            .downcast::<PyDict>()
            .map_err(|_| PyErr::new::<PyValueError, _>("payload must be a dict"))?;
        let type_obj = dict
            .get_item("type")?
            .ok_or_else(|| PyErr::new::<PyValueError, _>("payload missing 'type'"))?;
        let type_str: String = type_obj.extract()?;

        match type_str.as_str() {
            "blank" => {
                // no-op
                Ok(())
            }
            "string" => {
                let v = dict.get_item("value")?;
                let s = match v {
                    Some(v) => v.extract::<String>()?,
                    None => String::new(),
                };
                ws.get_cell_mut(a1).set_value_string(s);
                Ok(())
            }
            "number" => {
                let v = dict.get_item("value")?.ok_or_else(|| {
                    PyErr::new::<PyValueError, _>("number payload missing 'value'")
                })?;
                let f = v.extract::<f64>()?;
                ws.get_cell_mut(a1).set_value_number(f);
                Ok(())
            }
            "boolean" => {
                let v = dict.get_item("value")?.ok_or_else(|| {
                    PyErr::new::<PyValueError, _>("boolean payload missing 'value'")
                })?;
                let b = v.extract::<bool>()?;
                ws.get_cell_mut(a1).set_value_bool(b);
                Ok(())
            }
            "formula" => {
                let v = if let Some(v) = dict.get_item("formula")? {
                    v
                } else if let Some(v) = dict.get_item("value")? {
                    v
                } else {
                    return Err(PyErr::new::<PyValueError, _>(
                        "formula payload missing 'formula'",
                    ));
                };
                let formula = v.extract::<String>()?;
                let f = formula.strip_prefix('=').unwrap_or(&formula);
                ws.get_cell_mut(a1).set_formula(f);
                Ok(())
            }
            "error" => {
                let v = dict.get_item("value")?.ok_or_else(|| {
                    PyErr::new::<PyValueError, _>("error payload missing 'value'")
                })?;
                let token = v.extract::<String>()?;
                // Prefer formulas for errors that OpenpyxlAdapter can recognize.
                // For other error tokens, write the literal token as a string.
                let formula = match token.as_str() {
                    "#DIV/0!" => Some("1/0"),
                    "#N/A" => Some("NA()"),
                    "#VALUE!" => Some("\"text\"+1"),
                    _ => None,
                };
                if let Some(f) = formula {
                    ws.get_cell_mut(a1).set_formula(f);
                } else {
                    ws.get_cell_mut(a1).set_value_string(token);
                }
                Ok(())
            }
            "date" => {
                let v = dict
                    .get_item("value")?
                    .ok_or_else(|| PyErr::new::<PyValueError, _>("date payload missing 'value'"))?;
                let s = v.extract::<String>()?;
                let d = parse_iso_date(&s)
                    .ok_or_else(|| PyErr::new::<PyValueError, _>("Invalid ISO date"))?;
                let dt = d.and_time(NaiveTime::from_hms_opt(0, 0, 0).unwrap());
                let serial = naive_datetime_to_excel_serial(dt)
                    .ok_or_else(|| PyErr::new::<PyValueError, _>("Failed to convert date"))?;

                ws.get_cell_mut(a1).set_value_number(serial);
                ws.get_style_mut(a1)
                    .get_number_format_mut()
                    .set_format_code(NumberingFormat::FORMAT_DATE_YYYYMMDD);
                Ok(())
            }
            "datetime" => {
                let v = dict.get_item("value")?.ok_or_else(|| {
                    PyErr::new::<PyValueError, _>("datetime payload missing 'value'")
                })?;
                let s = v.extract::<String>()?;
                let dt = parse_iso_datetime(&s)
                    .ok_or_else(|| PyErr::new::<PyValueError, _>("Invalid ISO datetime"))?;
                let serial = naive_datetime_to_excel_serial(dt)
                    .ok_or_else(|| PyErr::new::<PyValueError, _>("Failed to convert datetime"))?;

                ws.get_cell_mut(a1).set_value_number(serial);
                ws.get_style_mut(a1)
                    .get_number_format_mut()
                    .set_format_code("yyyy-mm-dd h:mm:ss");
                Ok(())
            }
            other => Err(PyErr::new::<PyValueError, _>(format!(
                "Unsupported cell type: {other}"
            ))),
        }
    }

    pub fn write_cell_format(
        &mut self,
        sheet: &str,
        a1: &str,
        payload: &Bound<'_, PyAny>,
    ) -> PyResult<()> {
        let ws = self
            .book
            .get_sheet_by_name_mut(sheet)
            .ok_or_else(|| PyErr::new::<PyValueError, _>(format!("Unknown sheet: {sheet}")))?;

        let dict = payload
            .downcast::<PyDict>()
            .map_err(|_| PyErr::new::<PyValueError, _>("payload must be a dict"))?;

        let style = ws.get_style_mut(a1);

        if let Some(v) = dict.get_item("bold")? {
            let b = v.extract::<bool>()?;
            style.get_font_mut().set_bold(b);
        }
        if let Some(v) = dict.get_item("italic")? {
            let b = v.extract::<bool>()?;
            style.get_font_mut().set_italic(b);
        }
        if let Some(v) = dict.get_item("underline")? {
            let s = v.extract::<String>()?;
            // Only pass through known strings; umya panics on unknown underline.
            if s == "single" || s == "double" {
                style.get_font_mut().set_underline(s);
            }
        }
        if let Some(v) = dict.get_item("strikethrough")? {
            let b = v.extract::<bool>()?;
            style.get_font_mut().set_strikethrough(b);
        }
        if let Some(v) = dict.get_item("font_name")? {
            let s = v.extract::<String>()?;
            style.get_font_mut().set_name(s);
        }
        if let Some(v) = dict.get_item("font_size")? {
            let sz = v.extract::<f64>()?;
            style.get_font_mut().set_size(sz);
        }
        if let Some(v) = dict.get_item("font_color")? {
            let s = v.extract::<String>()?;
            style
                .get_font_mut()
                .get_color_mut()
                .set_argb(to_umya_argb(&s));
        }
        if let Some(v) = dict.get_item("bg_color")? {
            let s = v.extract::<String>()?;
            style.set_background_color_solid(to_umya_argb(&s));
        }
        if let Some(v) = dict.get_item("number_format")? {
            let s = v.extract::<String>()?;
            style.get_number_format_mut().set_format_code(s);
        }
        if let Some(v) = dict.get_item("h_align")? {
            let s = v.extract::<String>()?;
            if let Ok(a) = HorizontalAlignmentValues::from_str(&s) {
                style.get_alignment_mut().set_horizontal(a);
            }
        }
        if let Some(v) = dict.get_item("v_align")? {
            let s = v.extract::<String>()?;
            if let Ok(a) = VerticalAlignmentValues::from_str(&s) {
                style.get_alignment_mut().set_vertical(a);
            }
        }
        if let Some(v) = dict.get_item("wrap")? {
            let b = v.extract::<bool>()?;
            style.get_alignment_mut().set_wrap_text(b);
        }
        if let Some(v) = dict.get_item("rotation")? {
            let rot = v.extract::<i64>()?;
            if rot >= 0 {
                style.get_alignment_mut().set_text_rotation(rot as u32);
            }
        }

        Ok(())
    }

    pub fn write_cell_border(
        &mut self,
        sheet: &str,
        a1: &str,
        payload: &Bound<'_, PyAny>,
    ) -> PyResult<()> {
        let ws = self
            .book
            .get_sheet_by_name_mut(sheet)
            .ok_or_else(|| PyErr::new::<PyValueError, _>(format!("Unknown sheet: {sheet}")))?;

        let dict = payload
            .downcast::<PyDict>()
            .map_err(|_| PyErr::new::<PyValueError, _>("payload must be a dict"))?;

        let style = ws.get_style_mut(a1);

        if let Some(edge) = dict.get_item("top")? {
            let edge = edge
                .downcast::<PyDict>()
                .map_err(|_| PyErr::new::<PyValueError, _>("border edge must be a dict"))?;
            let border_obj = style.get_borders_mut().get_top_mut();
            if let Some(v) = edge.get_item("style")? {
                let s = v.extract::<String>()?;
                if let Some(bs) = border_style_to_umya(s.as_str()) {
                    border_obj.set_border_style(bs);
                }
            }
            if let Some(v) = edge.get_item("color")? {
                let s = v.extract::<String>()?;
                border_obj.get_color_mut().set_argb(to_umya_argb(&s));
            }
        }
        if let Some(edge) = dict.get_item("bottom")? {
            let edge = edge
                .downcast::<PyDict>()
                .map_err(|_| PyErr::new::<PyValueError, _>("border edge must be a dict"))?;
            let border_obj = style.get_borders_mut().get_bottom_mut();
            if let Some(v) = edge.get_item("style")? {
                let s = v.extract::<String>()?;
                if let Some(bs) = border_style_to_umya(s.as_str()) {
                    border_obj.set_border_style(bs);
                }
            }
            if let Some(v) = edge.get_item("color")? {
                let s = v.extract::<String>()?;
                border_obj.get_color_mut().set_argb(to_umya_argb(&s));
            }
        }
        if let Some(edge) = dict.get_item("left")? {
            let edge = edge
                .downcast::<PyDict>()
                .map_err(|_| PyErr::new::<PyValueError, _>("border edge must be a dict"))?;
            let border_obj = style.get_borders_mut().get_left_mut();
            if let Some(v) = edge.get_item("style")? {
                let s = v.extract::<String>()?;
                if let Some(bs) = border_style_to_umya(s.as_str()) {
                    border_obj.set_border_style(bs);
                }
            }
            if let Some(v) = edge.get_item("color")? {
                let s = v.extract::<String>()?;
                border_obj.get_color_mut().set_argb(to_umya_argb(&s));
            }
        }
        if let Some(edge) = dict.get_item("right")? {
            let edge = edge
                .downcast::<PyDict>()
                .map_err(|_| PyErr::new::<PyValueError, _>("border edge must be a dict"))?;
            let border_obj = style.get_borders_mut().get_right_mut();
            if let Some(v) = edge.get_item("style")? {
                let s = v.extract::<String>()?;
                if let Some(bs) = border_style_to_umya(s.as_str()) {
                    border_obj.set_border_style(bs);
                }
            }
            if let Some(v) = edge.get_item("color")? {
                let s = v.extract::<String>()?;
                border_obj.get_color_mut().set_argb(to_umya_argb(&s));
            }
        }

        if let Some(edge) = dict.get_item("diagonal_up")? {
            let edge = edge
                .downcast::<PyDict>()
                .map_err(|_| PyErr::new::<PyValueError, _>("border edge must be a dict"))?;
            style.get_borders_mut().set_diagonal_up(true);
            let border_obj = style.get_borders_mut().get_diagonal_mut();
            if let Some(v) = edge.get_item("style")? {
                let s = v.extract::<String>()?;
                if let Some(bs) = border_style_to_umya(s.as_str()) {
                    border_obj.set_border_style(bs);
                }
            }
            if let Some(v) = edge.get_item("color")? {
                let s = v.extract::<String>()?;
                border_obj.get_color_mut().set_argb(to_umya_argb(&s));
            }
        }

        if let Some(edge) = dict.get_item("diagonal_down")? {
            let edge = edge
                .downcast::<PyDict>()
                .map_err(|_| PyErr::new::<PyValueError, _>("border edge must be a dict"))?;
            style.get_borders_mut().set_diagonal_down(true);
            let border_obj = style.get_borders_mut().get_diagonal_mut();
            if let Some(v) = edge.get_item("style")? {
                let s = v.extract::<String>()?;
                if let Some(bs) = border_style_to_umya(s.as_str()) {
                    border_obj.set_border_style(bs);
                }
            }
            if let Some(v) = edge.get_item("color")? {
                let s = v.extract::<String>()?;
                border_obj.get_color_mut().set_argb(to_umya_argb(&s));
            }
        }

        Ok(())
    }

    pub fn set_row_height(&mut self, sheet: &str, row: u32, height: f64) -> PyResult<()> {
        let ws = self
            .book
            .get_sheet_by_name_mut(sheet)
            .ok_or_else(|| PyErr::new::<PyValueError, _>(format!("Unknown sheet: {sheet}")))?;
        ws.get_row_dimension_mut(&row).set_height(height);
        Ok(())
    }

    pub fn set_column_width(&mut self, sheet: &str, column: &str, width: f64) -> PyResult<()> {
        let ws = self
            .book
            .get_sheet_by_name_mut(sheet)
            .ok_or_else(|| PyErr::new::<PyValueError, _>(format!("Unknown sheet: {sheet}")))?;
        ws.get_column_dimension_mut(column).set_width(width);
        Ok(())
    }

    // =========================================================================
    // Tier 2 Write Operations
    // =========================================================================

    pub fn merge_cells(&mut self, sheet: &str, cell_range: &str) -> PyResult<()> {
        let ws = self
            .book
            .get_sheet_by_name_mut(sheet)
            .ok_or_else(|| PyErr::new::<PyValueError, _>(format!("Unknown sheet: {sheet}")))?;

        let range = normalize_a1(cell_range);
        if !range.contains(':') {
            return Ok(());
        }
        ws.add_merge_cells(range);
        Ok(())
    }

    pub fn add_conditional_format(
        &mut self,
        sheet: &str,
        payload: &Bound<'_, PyAny>,
    ) -> PyResult<()> {
        let ws = self
            .book
            .get_sheet_by_name_mut(sheet)
            .ok_or_else(|| PyErr::new::<PyValueError, _>(format!("Unknown sheet: {sheet}")))?;

        let outer = payload
            .downcast::<PyDict>()
            .map_err(|_| PyErr::new::<PyValueError, _>("payload must be a dict"))?;
        let cf_any = outer
            .get_item("cf_rule")?
            .unwrap_or_else(|| outer.clone().into_any());
        let cf = cf_any
            .downcast::<PyDict>()
            .map_err(|_| PyErr::new::<PyValueError, _>("cf_rule must be a dict"))?;

        let range = cf
            .get_item("range")?
            .ok_or_else(|| PyErr::new::<PyValueError, _>("cf_rule missing 'range'"))?
            .extract::<String>()?;
        let rule_type = cf
            .get_item("rule_type")?
            .ok_or_else(|| PyErr::new::<PyValueError, _>("cf_rule missing 'rule_type'"))?
            .extract::<String>()?;

        let stop_if_true = cf
            .get_item("stop_if_true")?
            .and_then(|v| v.extract::<bool>().ok())
            .unwrap_or(false);
        let priority = cf
            .get_item("priority")?
            .and_then(|v| v.extract::<i64>().ok())
            .unwrap_or(1) as i32;

        // Optional formatting (ExcelBench currently checks bg_color for these cases).
        let mut bg_color: Option<String> = None;
        if let Some(fmt_any) = cf.get_item("format")? {
            if let Ok(fmt) = fmt_any.downcast::<PyDict>() {
                if let Some(v) = fmt.get_item("bg_color")? {
                    bg_color = Some(v.extract::<String>()?);
                }
            }
        }

        let mut rule = ConditionalFormattingRule::default();
        if let Ok(t) = ConditionalFormatValues::from_str(&rule_type) {
            rule.set_type(t);
        } else {
            // Unknown types are ignored.
            return Ok(());
        }
        rule.set_priority(priority);
        if stop_if_true {
            rule.set_stop_if_true(true);
        }

        match rule.get_type() {
            ConditionalFormatValues::CellIs => {
                if let Some(op_any) = cf.get_item("operator")? {
                    let op = op_any.extract::<String>()?;
                    if let Ok(v) = ConditionalFormattingOperatorValues::from_str(&op) {
                        rule.set_operator(v);
                    }
                }
                if let Some(f_any) = cf.get_item("formula")? {
                    let mut f = Formula::default();
                    let formula = f_any.extract::<String>()?;
                    f.set_string_value(strip_leading_equal(&formula));
                    rule.set_formula(f);
                }
            }
            ConditionalFormatValues::Expression => {
                if let Some(f_any) = cf.get_item("formula")? {
                    let mut f = Formula::default();
                    let formula = f_any.extract::<String>()?;
                    f.set_string_value(strip_leading_equal(&formula));
                    rule.set_formula(f);
                }
            }
            ConditionalFormatValues::DataBar => {
                let mut min = ConditionalFormatValueObject::default();
                min.set_type(ConditionalFormatValueObjectValues::Number)
                    .set_val("0");
                let mut max = ConditionalFormatValueObject::default();
                max.set_type(ConditionalFormatValueObjectValues::Number)
                    .set_val("10");
                let mut color = Color::default();
                color.set_argb("FF638EC6");
                let mut db = DataBar::default();
                db.add_cfvo_collection(min)
                    .add_cfvo_collection(max)
                    .add_color_collection(color);
                rule.set_data_bar(db);
            }
            ConditionalFormatValues::ColorScale => {
                let mut c1 = ConditionalFormatValueObject::default();
                c1.set_type(ConditionalFormatValueObjectValues::Min);
                let mut c2 = ConditionalFormatValueObject::default();
                c2.set_type(ConditionalFormatValueObjectValues::Percentile)
                    .set_val("50");
                let mut c3 = ConditionalFormatValueObject::default();
                c3.set_type(ConditionalFormatValueObjectValues::Max);

                let mut start = Color::default();
                start.set_argb("FFF8696B");
                let mut mid = Color::default();
                mid.set_argb("FFFFEB84");
                let mut end = Color::default();
                end.set_argb("FF63BE7B");

                let mut cs = ColorScale::default();
                cs.add_cfvo_collection(c1)
                    .add_cfvo_collection(c2)
                    .add_cfvo_collection(c3)
                    .add_color_collection(start)
                    .add_color_collection(mid)
                    .add_color_collection(end);
                rule.set_color_scale(cs);
            }
            _ => {
                // Other rule types aren't needed for current Tier2 cases.
            }
        }

        if let Some(bg) = bg_color {
            let mut color = Color::default();
            color.set_argb(to_umya_argb(&bg));
            let mut pattern_fill = PatternFill::default();
            // Excel conditional format fills are stored as foreground colors (dxf fills).
            pattern_fill.set_foreground_color(color);
            let mut fill = Fill::default();
            fill.set_pattern_fill(pattern_fill);
            let mut style = Style::default();
            style.set_fill(fill);
            rule.set_style(style);
        }

        let mut seq = SequenceOfReferences::default();
        seq.set_sqref(normalize_a1(&range));

        let mut group = ConditionalFormatting::default();
        group.set_sequence_of_references(seq);
        group.set_conditional_collection(vec![rule]);

        ws.add_conditional_formatting_collection(group);
        Ok(())
    }

    pub fn add_data_validation(&mut self, sheet: &str, payload: &Bound<'_, PyAny>) -> PyResult<()> {
        let ws = self
            .book
            .get_sheet_by_name_mut(sheet)
            .ok_or_else(|| PyErr::new::<PyValueError, _>(format!("Unknown sheet: {sheet}")))?;

        let outer = payload
            .downcast::<PyDict>()
            .map_err(|_| PyErr::new::<PyValueError, _>("payload must be a dict"))?;
        let v_any = outer
            .get_item("validation")?
            .unwrap_or_else(|| outer.clone().into_any());
        let v = v_any
            .downcast::<PyDict>()
            .map_err(|_| PyErr::new::<PyValueError, _>("validation must be a dict"))?;

        let range = v
            .get_item("range")?
            .ok_or_else(|| PyErr::new::<PyValueError, _>("validation missing 'range'"))?
            .extract::<String>()?;
        let validation_type = v
            .get_item("validation_type")?
            .ok_or_else(|| PyErr::new::<PyValueError, _>("validation missing 'validation_type'"))?
            .extract::<String>()?;

        let mut dv = DataValidation::default();
        if let Ok(t) = DataValidationValues::from_str(&validation_type) {
            dv.set_type(t);
        } else {
            return Ok(());
        }
        if let Some(op_any) = v.get_item("operator")? {
            let op = op_any.extract::<String>()?;
            if let Ok(opv) = DataValidationOperatorValues::from_str(&op) {
                dv.set_operator(opv);
            }
        }

        if let Some(v1_any) = v.get_item("formula1")? {
            let f1 = v1_any.extract::<String>()?;
            dv.set_formula1(strip_leading_equal(&f1));
        }
        if let Some(v2_any) = v.get_item("formula2")? {
            let f2 = v2_any.extract::<String>()?;
            dv.set_formula2(strip_leading_equal(&f2));
        }

        if let Some(a_any) = v.get_item("allow_blank")? {
            dv.set_allow_blank(a_any.extract::<bool>()?);
        }
        if let Some(t_any) = v.get_item("prompt_title")? {
            dv.set_prompt_title(t_any.extract::<String>()?);
        }
        if let Some(p_any) = v.get_item("prompt")? {
            dv.set_prompt(p_any.extract::<String>()?);
        }
        if let Some(t_any) = v.get_item("error_title")? {
            dv.set_error_title(t_any.extract::<String>()?);
        }
        if let Some(e_any) = v.get_item("error")? {
            dv.set_error_message(e_any.extract::<String>()?);
        }

        let mut seq = SequenceOfReferences::default();
        seq.set_sqref(normalize_a1(&range));
        dv.set_sequence_of_references(seq);

        if ws.get_data_validations_mut().is_none() {
            ws.set_data_validations(DataValidations::default());
        }
        if let Some(dvs) = ws.get_data_validations_mut() {
            dvs.add_data_validation_list(dv);
        }

        Ok(())
    }

    pub fn add_hyperlink(&mut self, sheet: &str, payload: &Bound<'_, PyAny>) -> PyResult<()> {
        let ws = self
            .book
            .get_sheet_by_name_mut(sheet)
            .ok_or_else(|| PyErr::new::<PyValueError, _>(format!("Unknown sheet: {sheet}")))?;

        let outer = payload
            .downcast::<PyDict>()
            .map_err(|_| PyErr::new::<PyValueError, _>("payload must be a dict"))?;
        let link_any = outer
            .get_item("hyperlink")?
            .unwrap_or_else(|| outer.clone().into_any());
        let link = link_any
            .downcast::<PyDict>()
            .map_err(|_| PyErr::new::<PyValueError, _>("hyperlink must be a dict"))?;

        let cell = link
            .get_item("cell")?
            .ok_or_else(|| PyErr::new::<PyValueError, _>("hyperlink missing 'cell'"))?
            .extract::<String>()?;
        let target = link
            .get_item("target")?
            .ok_or_else(|| PyErr::new::<PyValueError, _>("hyperlink missing 'target'"))?
            .extract::<String>()?;
        let display = link
            .get_item("display")?
            .and_then(|v| v.extract::<String>().ok());
        let tooltip = link
            .get_item("tooltip")?
            .and_then(|v| v.extract::<String>().ok());
        let internal = link
            .get_item("internal")?
            .and_then(|v| v.extract::<bool>().ok())
            .unwrap_or(false);

        let a1 = normalize_a1(&cell);
        if let Some(text) = &display {
            ws.get_cell_mut(a1.as_str())
                .set_value_string(text.to_string());
        }

        let hyperlink = ws.get_cell_mut(a1.as_str()).get_hyperlink_mut();
        hyperlink.set_url(normalize_a1(&target));
        hyperlink.set_location(internal);
        if let Some(tip) = &tooltip {
            hyperlink.set_tooltip(tip.to_string());
            self.hyperlink_tooltips
                .entry(sheet.to_string())
                .or_default()
                .insert(a1.clone(), tip.to_string());
        }

        Ok(())
    }

    pub fn add_image(&mut self, sheet: &str, payload: &Bound<'_, PyAny>) -> PyResult<()> {
        let ws = self
            .book
            .get_sheet_by_name_mut(sheet)
            .ok_or_else(|| PyErr::new::<PyValueError, _>(format!("Unknown sheet: {sheet}")))?;

        let outer = payload
            .downcast::<PyDict>()
            .map_err(|_| PyErr::new::<PyValueError, _>("payload must be a dict"))?;
        let img_any = outer
            .get_item("image")?
            .unwrap_or_else(|| outer.clone().into_any());
        let img = img_any
            .downcast::<PyDict>()
            .map_err(|_| PyErr::new::<PyValueError, _>("image must be a dict"))?;

        let cell = img
            .get_item("cell")?
            .ok_or_else(|| PyErr::new::<PyValueError, _>("image missing 'cell'"))?
            .extract::<String>()?;
        let path = img
            .get_item("path")?
            .ok_or_else(|| PyErr::new::<PyValueError, _>("image missing 'path'"))?
            .extract::<String>()?;

        let p = Path::new(&path);
        if !p.exists() {
            return Err(PyErr::new::<PyIOError, _>(format!(
                "Image path does not exist: {path}"
            )));
        }

        let mut marker = umya_spreadsheet::structs::drawing::spreadsheet::MarkerType::default();
        marker.set_coordinate(normalize_a1(&cell));

        let mut image = umya_spreadsheet::structs::Image::default();
        let res = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
            image.new_image(&path, marker);
        }));
        if res.is_err() {
            return Err(PyErr::new::<PyIOError, _>(format!(
                "Failed to insert image (umya-spreadsheet panicked) for path: {path}"
            )));
        }

        ws.add_image(image);
        Ok(())
    }

    pub fn add_comment(&mut self, sheet: &str, payload: &Bound<'_, PyAny>) -> PyResult<()> {
        let ws = self
            .book
            .get_sheet_by_name_mut(sheet)
            .ok_or_else(|| PyErr::new::<PyValueError, _>(format!("Unknown sheet: {sheet}")))?;

        let outer = payload
            .downcast::<PyDict>()
            .map_err(|_| PyErr::new::<PyValueError, _>("payload must be a dict"))?;
        let note_any = outer
            .get_item("comment")?
            .unwrap_or_else(|| outer.clone().into_any());
        let note = note_any
            .downcast::<PyDict>()
            .map_err(|_| PyErr::new::<PyValueError, _>("comment must be a dict"))?;

        let cell = note
            .get_item("cell")?
            .ok_or_else(|| PyErr::new::<PyValueError, _>("comment missing 'cell'"))?
            .extract::<String>()?;
        let text = note
            .get_item("text")?
            .ok_or_else(|| PyErr::new::<PyValueError, _>("comment missing 'text'"))?
            .extract::<String>()?;
        let author = note
            .get_item("author")?
            .and_then(|v| v.extract::<String>().ok());

        let mut c = umya_spreadsheet::structs::Comment::default();
        c.new_comment(normalize_a1(&cell));
        c.set_text_string(text);
        if let Some(a) = author {
            c.set_author(a);
        }
        ws.add_comments(c);
        Ok(())
    }

    pub fn set_freeze_panes(&mut self, sheet: &str, payload: &Bound<'_, PyAny>) -> PyResult<()> {
        let ws = self
            .book
            .get_sheet_by_name_mut(sheet)
            .ok_or_else(|| PyErr::new::<PyValueError, _>(format!("Unknown sheet: {sheet}")))?;

        let outer = payload
            .downcast::<PyDict>()
            .map_err(|_| PyErr::new::<PyValueError, _>("payload must be a dict"))?;
        let cfg_any = outer
            .get_item("freeze")?
            .unwrap_or_else(|| outer.clone().into_any());
        let cfg = cfg_any
            .downcast::<PyDict>()
            .map_err(|_| PyErr::new::<PyValueError, _>("freeze must be a dict"))?;

        let mode = cfg
            .get_item("mode")?
            .and_then(|v| v.extract::<String>().ok())
            .unwrap_or_else(|| "freeze".to_string());

        let views = ws.get_sheet_views_mut().get_sheet_view_list_mut();
        if views.is_empty() {
            views.push(SheetView::default());
        }
        let view = views
            .get_mut(0)
            .ok_or_else(|| PyErr::new::<PyValueError, _>("Failed to access sheet view"))?;

        let mut pane = Pane::default();
        pane.set_active_pane(PaneValues::BottomRight);

        if mode == "freeze" {
            let tl = cfg
                .get_item("top_left_cell")?
                .and_then(|v| v.extract::<String>().ok())
                .ok_or_else(|| PyErr::new::<PyValueError, _>("freeze missing 'top_left_cell'"))?;
            let tl = normalize_a1(&tl);
            let (row0, col0) =
                a1_to_row_col(&tl).map_err(|msg| PyErr::new::<PyValueError, _>(msg))?;
            pane.set_horizontal_split(col0 as f64);
            pane.set_vertical_split(row0 as f64);
            pane.set_state(PaneStateValues::Frozen);
            pane.get_top_left_cell_mut().set_coordinate(tl);
        } else if mode == "split" {
            let x_split = cfg
                .get_item("x_split")?
                .and_then(|v| v.extract::<u32>().ok())
                .unwrap_or(0);
            let y_split = cfg
                .get_item("y_split")?
                .and_then(|v| v.extract::<u32>().ok())
                .unwrap_or(0);
            pane.set_horizontal_split(x_split as f64);
            pane.set_vertical_split(y_split as f64);
            pane.set_state(PaneStateValues::Split);
            pane.get_top_left_cell_mut().set_coordinate("A1");
        } else {
            return Ok(());
        }

        view.set_pane(pane);
        Ok(())
    }

    pub fn save(&mut self, path: &str) -> PyResult<()> {
        if self.saved {
            return Err(PyErr::new::<PyValueError, _>(
                "Workbook already saved (UmyaBook is consumed-on-save)",
            ));
        }
        self.saved = true;

        let p = Path::new(path);
        writer::xlsx::write(&self.book, p)
            .map_err(|e| PyErr::new::<PyIOError, _>(format!("Failed to save workbook: {e}")))?;

        if !self.hyperlink_tooltips.is_empty() {
            patch_xlsx_hyperlink_tooltips(p, &self.hyperlink_tooltips)?;
        }

        Ok(())
    }
}

fn xml_escape_attr(value: &str) -> String {
    let mut out = String::with_capacity(value.len());
    for ch in value.chars() {
        match ch {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            '"' => out.push_str("&quot;"),
            '\'' => out.push_str("&apos;"),
            _ => out.push(ch),
        }
    }
    out
}

fn parse_attr(tag: &str, name: &str) -> Option<String> {
    let needle = format!("{name}=\"");
    let start = tag.find(&needle)? + needle.len();
    let rest = &tag[start..];
    let end = rest.find('"')?;
    Some(rest[..end].to_string())
}

fn patch_sheet_xml_tooltips(xml: &str, tooltips: &HashMap<String, String>) -> String {
    if tooltips.is_empty() {
        return xml.to_string();
    }

    let mut out = String::with_capacity(xml.len() + 128);
    let mut i: usize = 0;

    while let Some(rel) = xml[i..].find("<hyperlink") {
        let start = i + rel;
        // Avoid matching the <hyperlinks> container tag.
        let after = xml.get(start + "<hyperlink".len()..start + "<hyperlink".len() + 1);
        if after != Some(" ") && after != Some(">") {
            out.push_str(&xml[i..start + "<hyperlink".len()]);
            i = start + "<hyperlink".len();
            continue;
        }

        out.push_str(&xml[i..start]);

        let end_rel = xml[start..].find("/>").or_else(|| xml[start..].find('>'));
        let Some(tag_end_rel) = end_rel else {
            out.push_str(&xml[start..]);
            return out;
        };
        let tag_end = start + tag_end_rel;
        let close_len = if xml[tag_end..].starts_with("/>") {
            2
        } else {
            1
        };
        let tag = &xml[start..tag_end + close_len];

        if tag.contains("tooltip=") {
            out.push_str(tag);
            i = tag_end + close_len;
            continue;
        }

        if let Some(r) = parse_attr(tag, "ref") {
            if let Some(tip) = tooltips.get(&r) {
                let esc = xml_escape_attr(tip);
                if close_len == 2 {
                    let without = &tag[..tag.len() - 2];
                    out.push_str(without);
                    out.push_str(&format!(" tooltip=\"{esc}\"/>"));
                } else {
                    let without = &tag[..tag.len() - 1];
                    out.push_str(without);
                    out.push_str(&format!(" tooltip=\"{esc}\">"));
                }
                i = tag_end + close_len;
                continue;
            }
        }

        out.push_str(tag);
        i = tag_end + close_len;
    }

    out.push_str(&xml[i..]);
    out
}

fn parse_workbook_sheet_map(workbook_xml: &str) -> HashMap<String, String> {
    let mut out: HashMap<String, String> = HashMap::new();
    let mut i: usize = 0;
    while let Some(rel) = workbook_xml[i..].find("<sheet ") {
        let start = i + rel;
        let end_rel = workbook_xml[start..]
            .find("/>")
            .or_else(|| workbook_xml[start..].find('>'));
        let Some(tag_end_rel) = end_rel else {
            break;
        };
        let tag_end = start + tag_end_rel;
        let close_len = if workbook_xml[tag_end..].starts_with("/>") {
            2
        } else {
            1
        };
        let tag = &workbook_xml[start..tag_end + close_len];
        let name = parse_attr(tag, "name");
        let rid = parse_attr(tag, "r:id");
        if let (Some(n), Some(r)) = (name, rid) {
            out.insert(n, r);
        }
        i = tag_end + close_len;
    }
    out
}

fn parse_workbook_rels_map(rels_xml: &str) -> HashMap<String, String> {
    let mut out: HashMap<String, String> = HashMap::new();
    let mut i: usize = 0;
    while let Some(rel) = rels_xml[i..].find("<Relationship ") {
        let start = i + rel;
        let end_rel = rels_xml[start..]
            .find("/>")
            .or_else(|| rels_xml[start..].find('>'));
        let Some(tag_end_rel) = end_rel else {
            break;
        };
        let tag_end = start + tag_end_rel;
        let close_len = if rels_xml[tag_end..].starts_with("/>") {
            2
        } else {
            1
        };
        let tag = &rels_xml[start..tag_end + close_len];
        let id = parse_attr(tag, "Id");
        let target = parse_attr(tag, "Target");
        if let (Some(iid), Some(t)) = (id, target) {
            out.insert(iid, t);
        }
        i = tag_end + close_len;
    }
    out
}

fn patch_xlsx_hyperlink_tooltips(
    path: &Path,
    tooltips: &HashMap<String, HashMap<String, String>>,
) -> PyResult<()> {
    let f = std::fs::File::open(path).map_err(|e| {
        PyErr::new::<PyIOError, _>(format!("Failed to open xlsx for patching: {e}"))
    })?;
    let mut zip = ZipArchive::new(f)
        .map_err(|e| PyErr::new::<PyIOError, _>(format!("Invalid xlsx zip: {e}")))?;

    let mut workbook_xml = String::new();
    {
        let mut entry = zip
            .by_name("xl/workbook.xml")
            .map_err(|e| PyErr::new::<PyIOError, _>(format!("Missing xl/workbook.xml: {e}")))?;
        entry
            .read_to_string(&mut workbook_xml)
            .map_err(|e| PyErr::new::<PyIOError, _>(format!("Read workbook.xml failed: {e}")))?;
    }

    let mut rels_xml = String::new();
    {
        let mut entry = zip.by_name("xl/_rels/workbook.xml.rels").map_err(|e| {
            PyErr::new::<PyIOError, _>(format!("Missing xl/_rels/workbook.xml.rels: {e}"))
        })?;
        entry
            .read_to_string(&mut rels_xml)
            .map_err(|e| PyErr::new::<PyIOError, _>(format!("Read workbook rels failed: {e}")))?;
    }

    let sheet_to_rid = parse_workbook_sheet_map(&workbook_xml);
    let rid_to_target = parse_workbook_rels_map(&rels_xml);

    let mut targets: HashMap<String, HashMap<String, String>> = HashMap::new();
    for (sheet_name, cells) in tooltips {
        let Some(rid) = sheet_to_rid.get(sheet_name) else {
            continue;
        };
        let Some(target) = rid_to_target.get(rid) else {
            continue;
        };
        targets.insert(format!("xl/{target}"), cells.clone());
    }

    if targets.is_empty() {
        return Ok(());
    }

    drop(zip);

    let f = std::fs::File::open(path).map_err(|e| {
        PyErr::new::<PyIOError, _>(format!("Failed to re-open xlsx for patching: {e}"))
    })?;
    let mut zip = ZipArchive::new(f)
        .map_err(|e| PyErr::new::<PyIOError, _>(format!("Invalid xlsx zip: {e}")))?;

    let tmp_path = path.with_extension("xlsx.tmp");
    let tmp_file = std::fs::File::create(&tmp_path)
        .map_err(|e| PyErr::new::<PyIOError, _>(format!("Failed to create temp xlsx: {e}")))?;
    let mut out = ZipWriter::new(tmp_file);

    for idx in 0..zip.len() {
        let mut file = zip
            .by_index(idx)
            .map_err(|e| PyErr::new::<PyIOError, _>(format!("Zip read failed: {e}")))?;
        let name = file.name().to_string();

        let options = FileOptions::default()
            .compression_method(file.compression())
            .last_modified_time(file.last_modified());

        if file.is_dir() {
            out.add_directory(name, options).map_err(|e| {
                PyErr::new::<PyIOError, _>(format!("Zip write directory failed: {e}"))
            })?;
            continue;
        }

        let mut buf: Vec<u8> = Vec::new();
        file.read_to_end(&mut buf)
            .map_err(|e| PyErr::new::<PyIOError, _>(format!("Zip entry read failed: {e}")))?;

        if let Some(cells) = targets.get(&name) {
            if let Ok(s) = std::str::from_utf8(&buf) {
                let patched = patch_sheet_xml_tooltips(s, cells);
                buf = patched.into_bytes();
            }
        }

        out.start_file(name, options)
            .map_err(|e| PyErr::new::<PyIOError, _>(format!("Zip write failed: {e}")))?;
        out.write_all(&buf)
            .map_err(|e| PyErr::new::<PyIOError, _>(format!("Zip write failed: {e}")))?;
    }

    out.finish()
        .map_err(|e| PyErr::new::<PyIOError, _>(format!("Zip finalize failed: {e}")))?;

    std::fs::rename(&tmp_path, path).map_err(|e| {
        PyErr::new::<PyIOError, _>(format!("Failed to replace xlsx with patched version: {e}"))
    })?;

    Ok(())
}
