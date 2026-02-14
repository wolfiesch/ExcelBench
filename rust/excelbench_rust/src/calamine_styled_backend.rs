use pyo3::exceptions::{PyIOError, PyValueError};
use pyo3::prelude::*;
use pyo3::types::PyDict;

use std::collections::HashMap;
use std::fs::File;
use std::io::BufReader;

use calamine::{Data, Reader, Xlsx};
use calamine::{
    Alignment, BorderStyle as CalBorderStyle, Color, Fill, FillPattern, Font, FontStyle,
    FontWeight, HorizontalAlignment, Style, StyleRange, TextRotation,
    UnderlineStyle, VerticalAlignment, WorksheetLayout,
};
use chrono::NaiveTime;

use crate::util::{a1_to_row_col, cell_blank, cell_with_value, parse_iso_date, parse_iso_datetime};

fn map_error_value(err_str: &str) -> &'static str {
    let e = err_str.to_ascii_uppercase();
    match e.as_str() {
        "DIV0" | "DIV/0" | "#DIV/0!" => "#DIV/0!",
        "NA" | "#N/A" => "#N/A",
        "VALUE" | "#VALUE!" => "#VALUE!",
        "REF" | "#REF!" => "#REF!",
        "NAME" | "#NAME?" => "#NAME?",
        "NUM" | "#NUM!" => "#NUM!",
        "NULL" | "#NULL!" => "#NULL!",
        _ => "#ERROR!",
    }
}

/// Convert a calamine Color to a "#RRGGBB" hex string.
fn color_to_hex(c: &Color) -> String {
    format!("#{:02X}{:02X}{:02X}", c.red, c.green, c.blue)
}

/// Convert a calamine BorderStyle to the ExcelBench string token.
fn border_style_str(s: &CalBorderStyle) -> &'static str {
    match s {
        CalBorderStyle::None => "none",
        CalBorderStyle::Thin => "thin",
        CalBorderStyle::Medium => "medium",
        CalBorderStyle::Thick => "thick",
        CalBorderStyle::Double => "double",
        CalBorderStyle::Hair => "hair",
        CalBorderStyle::Dashed => "dashed",
        CalBorderStyle::Dotted => "dotted",
        CalBorderStyle::MediumDashed => "mediumDashed",
        CalBorderStyle::DashDot => "dashDot",
        CalBorderStyle::DashDotDot => "dashDotDot",
        CalBorderStyle::SlantDashDot => "slantDashDot",
    }
}

/// Convert a calamine HorizontalAlignment to the ExcelBench string.
fn h_align_str(a: &HorizontalAlignment) -> Option<&'static str> {
    match a {
        HorizontalAlignment::General => None, // default — omit
        HorizontalAlignment::Left => Some("left"),
        HorizontalAlignment::Center => Some("center"),
        HorizontalAlignment::Right => Some("right"),
        HorizontalAlignment::Justify => Some("justify"),
        HorizontalAlignment::Distributed => Some("distributed"),
        HorizontalAlignment::Fill => Some("fill"),
    }
}

/// Convert a calamine VerticalAlignment to the ExcelBench string.
fn v_align_str(a: &VerticalAlignment) -> Option<&'static str> {
    match a {
        VerticalAlignment::Bottom => None, // default — omit
        VerticalAlignment::Top => Some("top"),
        VerticalAlignment::Center => Some("center"),
        VerticalAlignment::Justify => Some("justify"),
        VerticalAlignment::Distributed => Some("distributed"),
    }
}

/// Convert a calamine UnderlineStyle to the ExcelBench string.
fn underline_str(u: &UnderlineStyle) -> Option<&'static str> {
    match u {
        UnderlineStyle::None => None,
        UnderlineStyle::Single => Some("single"),
        UnderlineStyle::Double => Some("double"),
        UnderlineStyle::SingleAccounting => Some("singleAccounting"),
        UnderlineStyle::DoubleAccounting => Some("doubleAccounting"),
    }
}

type XlsxReader = Xlsx<BufReader<File>>;

/// Per-sheet cached data: style grid + layout dimensions.
struct SheetCache {
    styles: StyleRange,
    layout: WorksheetLayout,
    /// Offset from StyleRange.start() so we can look up absolute (row,col).
    style_origin: (u32, u32),
}

#[pyclass(unsendable)]
pub struct CalamineStyledBook {
    workbook: XlsxReader,
    sheet_names: Vec<String>,
    /// Cache of StyleRange per sheet name, populated lazily on first format/border read.
    style_cache: HashMap<String, SheetCache>,
}

#[pymethods]
impl CalamineStyledBook {
    #[staticmethod]
    pub fn open(path: &str) -> PyResult<Self> {
        let file = File::open(path)
            .map_err(|e| PyErr::new::<PyIOError, _>(format!("Failed to open file: {e}")))?;
        let reader = BufReader::new(file);
        let wb: XlsxReader = Xlsx::new(reader)
            .map_err(|e| PyErr::new::<PyIOError, _>(format!("Failed to parse xlsx: {e}")))?;
        let names = wb.sheet_names().to_vec();
        Ok(Self {
            workbook: wb,
            sheet_names: names,
            style_cache: HashMap::new(),
        })
    }

    pub fn sheet_names(&self) -> Vec<String> {
        self.sheet_names.clone()
    }

    pub fn read_cell_value(&mut self, py: Python<'_>, sheet: &str, a1: &str) -> PyResult<PyObject> {
        let (row, col) = a1_to_row_col(a1).map_err(|msg| PyErr::new::<PyValueError, _>(msg))?;

        if !self.sheet_names.iter().any(|name| name == sheet) {
            return Err(PyErr::new::<PyValueError, _>(format!(
                "Unknown sheet: {sheet}"
            )));
        }

        let range = self.workbook.worksheet_range(sheet).map_err(|e| {
            PyErr::new::<PyIOError, _>(format!("Failed to read sheet {sheet}: {e}"))
        })?;

        let value = match range.get_value((row, col)) {
            None => return cell_blank(py),
            Some(v) => v,
        };

        let out = match value {
            Data::Empty => cell_blank(py)?,
            Data::String(s) => cell_with_value(py, "string", s.clone())?,
            Data::Float(f) => cell_with_value(py, "number", *f)?,
            Data::Int(i) => cell_with_value(py, "number", *i as f64)?,
            Data::Bool(b) => cell_with_value(py, "boolean", *b)?,
            Data::DateTime(dt) => {
                if let Some(ndt) = dt.as_datetime() {
                    let midnight = NaiveTime::from_hms_opt(0, 0, 0).unwrap();
                    if ndt.time() == midnight {
                        let s = ndt.date().format("%Y-%m-%d").to_string();
                        cell_with_value(py, "date", s)?
                    } else {
                        let s = ndt.format("%Y-%m-%dT%H:%M:%S").to_string();
                        cell_with_value(py, "datetime", s)?
                    }
                } else {
                    cell_with_value(py, "number", dt.as_f64())?
                }
            }
            Data::DateTimeIso(s) => {
                let raw = s.trim_end_matches('Z');
                if let Some(d) = parse_iso_date(raw) {
                    cell_with_value(py, "date", d.format("%Y-%m-%d").to_string())?
                } else if let Some(ndt) = parse_iso_datetime(raw) {
                    let midnight = NaiveTime::from_hms_opt(0, 0, 0).unwrap();
                    if ndt.time() == midnight {
                        cell_with_value(py, "date", ndt.date().format("%Y-%m-%d").to_string())?
                    } else {
                        cell_with_value(
                            py,
                            "datetime",
                            ndt.format("%Y-%m-%dT%H:%M:%S").to_string(),
                        )?
                    }
                } else {
                    cell_with_value(py, "datetime", s.clone())?
                }
            }
            Data::DurationIso(s) => cell_with_value(py, "string", s.clone())?,
            Data::RichText(rt) => cell_with_value(py, "string", rt.plain_text())?,
            Data::Error(e) => {
                let normalized = map_error_value(&format!("{e:?}"));
                let d = PyDict::new(py);
                d.set_item("type", "error")?;
                d.set_item("value", normalized)?;
                d.into()
            }
        };

        Ok(out)
    }

    pub fn read_cell_format(
        &mut self,
        py: Python<'_>,
        sheet: &str,
        a1: &str,
    ) -> PyResult<PyObject> {
        let (row, col) = a1_to_row_col(a1).map_err(|msg| PyErr::new::<PyValueError, _>(msg))?;
        let style = self.get_style(sheet, row, col)?;
        let d = PyDict::new(py);

        if let Some(style) = style {
            // Font
            if let Some(font) = &style.font {
                Self::populate_font(py, &d, font)?;
            }
            // Fill
            if let Some(fill) = &style.fill {
                Self::populate_fill(py, &d, fill)?;
            }
            // NumberFormat
            if let Some(nf) = &style.number_format {
                if nf.format_code != "General" {
                    d.set_item("number_format", &nf.format_code)?;
                }
            }
            // Alignment
            if let Some(align) = &style.alignment {
                Self::populate_alignment(py, &d, align)?;
            }
        }

        Ok(d.into())
    }

    pub fn read_cell_border(
        &mut self,
        py: Python<'_>,
        sheet: &str,
        a1: &str,
    ) -> PyResult<PyObject> {
        let (row, col) = a1_to_row_col(a1).map_err(|msg| PyErr::new::<PyValueError, _>(msg))?;
        let style = self.get_style(sheet, row, col)?;
        let d = PyDict::new(py);

        if let Some(style) = style {
            if let Some(borders) = &style.borders {
                Self::maybe_set_edge(py, &d, "top", &borders.top)?;
                Self::maybe_set_edge(py, &d, "bottom", &borders.bottom)?;
                Self::maybe_set_edge(py, &d, "left", &borders.left)?;
                Self::maybe_set_edge(py, &d, "right", &borders.right)?;
                Self::maybe_set_edge(py, &d, "diagonal_up", &borders.diagonal_up)?;
                Self::maybe_set_edge(py, &d, "diagonal_down", &borders.diagonal_down)?;
            }
        }

        Ok(d.into())
    }

    pub fn read_row_height(&mut self, sheet: &str, row: i64) -> PyResult<Option<f64>> {
        // ExcelBench uses 1-indexed rows.
        let row_0 = (row - 1) as u32;
        self.ensure_cache(sheet)?;
        let cache = self.style_cache.get(sheet).unwrap();
        Ok(cache
            .layout
            .get_row_height(row_0)
            .filter(|rh| rh.custom_height)
            .map(|rh| rh.height))
    }

    pub fn read_column_width(&mut self, sheet: &str, col_letter: &str) -> PyResult<Option<f64>> {
        let col_0 = Self::col_letter_to_index(col_letter)?;
        self.ensure_cache(sheet)?;
        let cache = self.style_cache.get(sheet).unwrap();
        Ok(cache
            .layout
            .get_column_width(col_0)
            .filter(|cw| cw.custom_width)
            .map(|cw| cw.width))
    }
}

// Non-Python helper methods.
impl CalamineStyledBook {
    /// Ensure the StyleRange + WorksheetLayout are cached for this sheet.
    fn ensure_cache(&mut self, sheet: &str) -> PyResult<()> {
        if self.style_cache.contains_key(sheet) {
            return Ok(());
        }
        let styles = self
            .workbook
            .worksheet_style(sheet)
            .map_err(|e| PyErr::new::<PyIOError, _>(format!("Style error for {sheet}: {e}")))?;
        let layout = self
            .workbook
            .worksheet_layout(sheet)
            .map_err(|e| PyErr::new::<PyIOError, _>(format!("Layout error for {sheet}: {e}")))?;
        let origin = styles.start().unwrap_or((0, 0));
        self.style_cache.insert(
            sheet.to_string(),
            SheetCache {
                styles,
                layout,
                style_origin: origin,
            },
        );
        Ok(())
    }

    /// Get the Style for an absolute (row, col) position, or None if no style applied.
    fn get_style(&mut self, sheet: &str, row: u32, col: u32) -> PyResult<Option<Style>> {
        self.ensure_cache(sheet)?;
        let cache = self.style_cache.get(sheet).unwrap();
        let (or, oc) = cache.style_origin;
        if row < or || col < oc {
            return Ok(None);
        }
        let pos = ((row - or) as usize, (col - oc) as usize);
        Ok(cache.styles.get(pos).cloned())
    }

    fn populate_font(
        _py: Python<'_>,
        d: &Bound<'_, PyDict>,
        font: &Font,
    ) -> PyResult<()> {
        if font.weight == FontWeight::Bold {
            d.set_item("bold", true)?;
        }
        if font.style == FontStyle::Italic {
            d.set_item("italic", true)?;
        }
        if let Some(u) = underline_str(&font.underline) {
            d.set_item("underline", u)?;
        }
        if font.strikethrough {
            d.set_item("strikethrough", true)?;
        }
        if let Some(name) = &font.name {
            d.set_item("font_name", name.as_str())?;
        }
        if let Some(size) = font.size {
            d.set_item("font_size", size)?;
        }
        if let Some(color) = &font.color {
            d.set_item("font_color", color_to_hex(color))?;
        }
        Ok(())
    }

    fn populate_fill(
        _py: Python<'_>,
        d: &Bound<'_, PyDict>,
        fill: &Fill,
    ) -> PyResult<()> {
        if fill.pattern != FillPattern::None {
            if let Some(color) = fill.get_color() {
                d.set_item("bg_color", color_to_hex(&color))?;
            }
        }
        Ok(())
    }

    fn populate_alignment(
        _py: Python<'_>,
        d: &Bound<'_, PyDict>,
        align: &Alignment,
    ) -> PyResult<()> {
        if let Some(h) = h_align_str(&align.horizontal) {
            d.set_item("h_align", h)?;
        }
        if let Some(v) = v_align_str(&align.vertical) {
            d.set_item("v_align", v)?;
        }
        if align.wrap_text {
            d.set_item("wrap", true)?;
        }
        match align.text_rotation {
            TextRotation::None => {}
            TextRotation::Degrees(deg) => {
                if deg != 0 {
                    d.set_item("rotation", deg)?;
                }
            }
            TextRotation::Stacked => {
                d.set_item("rotation", 255)?;
            }
        }
        if let Some(indent) = align.indent {
            if indent > 0 {
                d.set_item("indent", indent)?;
            }
        }
        Ok(())
    }

    fn maybe_set_edge(
        py: Python<'_>,
        d: &Bound<'_, PyDict>,
        key: &str,
        border: &calamine::Border,
    ) -> PyResult<()> {
        if border.style == CalBorderStyle::None {
            return Ok(());
        }
        let edge = PyDict::new(py);
        edge.set_item("style", border_style_str(&border.style))?;
        let color_str = border
            .color
            .as_ref()
            .map(|c| color_to_hex(c))
            .unwrap_or_else(|| "#000000".to_string());
        edge.set_item("color", color_str)?;
        d.set_item(key, edge)?;
        Ok(())
    }

    fn col_letter_to_index(col: &str) -> PyResult<u32> {
        let mut idx: u32 = 0;
        for ch in col.chars() {
            if !ch.is_ascii_alphabetic() {
                return Err(PyErr::new::<PyValueError, _>(format!(
                    "Invalid column letter: {col}"
                )));
            }
            idx = idx * 26 + (ch.to_ascii_uppercase() as u8 - b'A' + 1) as u32;
        }
        Ok(idx - 1)
    }
}
