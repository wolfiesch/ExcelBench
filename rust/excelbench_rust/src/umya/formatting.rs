use pyo3::exceptions::PyValueError;
use pyo3::prelude::*;
use pyo3::types::PyDict;

use std::str::FromStr;

use umya_spreadsheet::structs::{
    EnumTrait, HorizontalAlignmentValues, PatternValues, VerticalAlignmentValues,
};

use crate::util::a1_to_row_col;

use super::util::{argb_to_hex, hex_to_argb};
use super::UmyaBook;

#[pymethods]
impl UmyaBook {
    pub fn read_cell_format(&self, py: Python<'_>, sheet: &str, a1: &str) -> PyResult<PyObject> {
        let ws = self
            .book
            .get_sheet_by_name(sheet)
            .ok_or_else(|| PyErr::new::<PyValueError, _>(format!("Unknown sheet: {sheet}")))?;

        let (row0, col0) = a1_to_row_col(a1).map_err(|msg| PyErr::new::<PyValueError, _>(msg))?;
        let coord = (col0 + 1, row0 + 1);

        let d = PyDict::new(py);

        let cell = match ws.get_cell(coord) {
            Some(c) => c,
            None => return Ok(d.into()),
        };

        let style = cell.get_style();

        // Font properties
        if let Some(font) = style.get_font() {
            if *font.get_bold() {
                d.set_item("bold", true)?;
            }
            if *font.get_italic() {
                d.set_item("italic", true)?;
            }
            {
                let ul = font.get_underline();
                if !ul.is_empty() && ul != "none" {
                    d.set_item("underline", ul.to_string())?;
                }
            }
            if *font.get_strikethrough() {
                d.set_item("strikethrough", true)?;
            }
            {
                let name = font.get_name();
                if !name.is_empty() {
                    d.set_item("font_name", name.to_string())?;
                }
            }
            {
                let size = *font.get_size();
                if size > 0.0 {
                    d.set_item("font_size", size)?;
                }
            }
            {
                let argb = font.get_color().get_argb();
                if !argb.is_empty() {
                    let hex = argb_to_hex(argb);
                    if hex != "#000000" {
                        d.set_item("font_color", hex)?;
                    }
                }
            }
        }

        // Fill / background color
        if let Some(fill) = style.get_fill() {
            if let Some(pf) = fill.get_pattern_fill() {
                if let Some(fg) = pf.get_foreground_color() {
                    let argb = fg.get_argb();
                    if !argb.is_empty() {
                        let hex = argb_to_hex(argb);
                        d.set_item("bg_color", hex)?;
                    }
                }
            }
        }

        // Number format
        if let Some(nf) = style.get_number_format() {
            let code = nf.get_format_code();
            if !code.is_empty() && code != "General" {
                d.set_item("number_format", code.to_string())?;
            }
        }

        // Alignment
        if let Some(align) = style.get_alignment() {
            let h = align.get_horizontal().get_value_string();
            if !h.is_empty() && h != "general" {
                d.set_item("h_align", h.to_string())?;
            }
            let v = align.get_vertical().get_value_string();
            if !v.is_empty() && v != "bottom" {
                d.set_item("v_align", v.to_string())?;
            }
            if *align.get_wrap_text() {
                d.set_item("wrap", true)?;
            }
            let rot = *align.get_text_rotation();
            if rot != 0 {
                d.set_item("rotation", rot)?;
            }
        }

        Ok(d.into())
    }

    pub fn write_cell_format(
        &mut self,
        sheet: &str,
        a1: &str,
        format_dict: &Bound<'_, PyAny>,
    ) -> PyResult<()> {
        let ws = self
            .book
            .get_sheet_by_name_mut(sheet)
            .ok_or_else(|| PyErr::new::<PyValueError, _>(format!("Unknown sheet: {sheet}")))?;

        let dict = format_dict
            .downcast::<PyDict>()
            .map_err(|_| PyErr::new::<PyValueError, _>("format_dict must be a dict"))?;

        let style = ws.get_style_mut(a1);

        // Font properties
        {
            let font = style.get_font_mut();

            if let Some(bold) = dict
                .get_item("bold")?
                .and_then(|v| v.extract::<bool>().ok())
            {
                font.set_bold(bold);
            }
            if let Some(italic) = dict
                .get_item("italic")?
                .and_then(|v| v.extract::<bool>().ok())
            {
                font.set_italic(italic);
            }
            if let Some(ul) = dict
                .get_item("underline")?
                .and_then(|v| v.extract::<String>().ok())
            {
                font.set_underline(ul);
            }
            if let Some(st) = dict
                .get_item("strikethrough")?
                .and_then(|v| v.extract::<bool>().ok())
            {
                font.set_strikethrough(st);
            }
            if let Some(name) = dict
                .get_item("font_name")?
                .and_then(|v| v.extract::<String>().ok())
            {
                font.set_name(name);
            }
            if let Some(size) = dict
                .get_item("font_size")?
                .and_then(|v| v.extract::<f64>().ok())
            {
                font.set_size(size);
            }
            if let Some(color) = dict
                .get_item("font_color")?
                .and_then(|v| v.extract::<String>().ok())
            {
                font.get_color_mut().set_argb(hex_to_argb(&color));
            }
        }

        // Background color via pattern fill
        if let Some(bg) = dict
            .get_item("bg_color")?
            .and_then(|v| v.extract::<String>().ok())
        {
            let fill = style.get_fill_mut();
            let pf = fill.get_pattern_fill_mut();
            pf.set_pattern_type(PatternValues::Solid);
            pf.get_foreground_color_mut().set_argb(hex_to_argb(&bg));
        }

        // Number format
        if let Some(nf) = dict
            .get_item("number_format")?
            .and_then(|v| v.extract::<String>().ok())
        {
            style.get_number_format_mut().set_format_code(nf);
        }

        // Alignment
        if let Some(h) = dict
            .get_item("h_align")?
            .and_then(|v| v.extract::<String>().ok())
        {
            if let Ok(ha) = HorizontalAlignmentValues::from_str(&h) {
                style.get_alignment_mut().set_horizontal(ha);
            }
        }
        if let Some(v) = dict
            .get_item("v_align")?
            .and_then(|v| v.extract::<String>().ok())
        {
            if let Ok(va) = VerticalAlignmentValues::from_str(&v) {
                style.get_alignment_mut().set_vertical(va);
            }
        }
        if let Some(wrap) = dict
            .get_item("wrap")?
            .and_then(|v| v.extract::<bool>().ok())
        {
            style.get_alignment_mut().set_wrap_text(wrap);
        }
        if let Some(rot) = dict
            .get_item("rotation")?
            .and_then(|v| v.extract::<i64>().ok())
        {
            if let Ok(r) = u32::try_from(rot) {
                style.get_alignment_mut().set_text_rotation(r);
            }
        }

        Ok(())
    }
}
