use pyo3::exceptions::PyValueError;
use pyo3::prelude::*;
use pyo3::types::PyDict;

use crate::util::a1_to_row_col;

use super::util::{argb_to_hex, hex_to_argb, umya_border_style_to_str};
use super::UmyaBook;

#[pymethods]
impl UmyaBook {
    pub fn read_cell_border(
        &self,
        py: Python<'_>,
        sheet: &str,
        a1: &str,
    ) -> PyResult<PyObject> {
        let ws = self
            .book
            .get_sheet_by_name(sheet)
            .ok_or_else(|| PyErr::new::<PyValueError, _>(format!("Unknown sheet: {sheet}")))?;

        let (row0, col0) = a1_to_row_col(a1).map_err(|msg| PyErr::new::<PyValueError, _>(msg))?;
        let coord = (col0 + 1, row0 + 1);

        let d = PyDict::new_bound(py);

        let cell = match ws.get_cell(coord) {
            Some(c) => c,
            None => return Ok(d.into()),
        };

        let style = cell.get_style();
        if let Some(borders) = style.get_borders() {
            let read_edge = |e: &umya_spreadsheet::structs::Border| -> Option<(String, String)> {
                let style_str = e.get_border_style();
                if style_str.is_empty() || style_str == "none" {
                    return None;
                }
                let argb = e.get_color().get_argb();
                let color_str = if argb.is_empty() {
                    "#000000".to_string()
                } else {
                    argb_to_hex(argb)
                };
                Some((
                    umya_border_style_to_str(style_str).to_string(),
                    color_str,
                ))
            };

            if let Some((s, c)) = read_edge(borders.get_top()) {
                let edge = PyDict::new_bound(py);
                edge.set_item("style", s)?;
                edge.set_item("color", c)?;
                d.set_item("top", edge)?;
            }
            if let Some((s, c)) = read_edge(borders.get_bottom()) {
                let edge = PyDict::new_bound(py);
                edge.set_item("style", s)?;
                edge.set_item("color", c)?;
                d.set_item("bottom", edge)?;
            }
            if let Some((s, c)) = read_edge(borders.get_left()) {
                let edge = PyDict::new_bound(py);
                edge.set_item("style", s)?;
                edge.set_item("color", c)?;
                d.set_item("left", edge)?;
            }
            if let Some((s, c)) = read_edge(borders.get_right()) {
                let edge = PyDict::new_bound(py);
                edge.set_item("style", s)?;
                edge.set_item("color", c)?;
                d.set_item("right", edge)?;
            }
            if let Some((s, c)) = read_edge(borders.get_diagonal()) {
                let edge = PyDict::new_bound(py);
                edge.set_item("style", s)?;
                edge.set_item("color", c)?;
                d.set_item("diagonal_up", edge)?;
            }
        }

        Ok(d.into())
    }

    pub fn write_cell_border(
        &mut self,
        sheet: &str,
        a1: &str,
        border_dict: &Bound<'_, PyAny>,
    ) -> PyResult<()> {
        let ws = self
            .book
            .get_sheet_by_name_mut(sheet)
            .ok_or_else(|| PyErr::new::<PyValueError, _>(format!("Unknown sheet: {sheet}")))?;

        let dict = border_dict
            .downcast::<PyDict>()
            .map_err(|_| PyErr::new::<PyValueError, _>("border_dict must be a dict"))?;

        let style = ws.get_style_mut(a1);
        let borders = style.get_borders_mut();

        fn apply_edge(
            edge: &mut umya_spreadsheet::structs::Border,
            sub: &Bound<'_, PyDict>,
        ) -> PyResult<()> {
            if let Some(s) = sub
                .get_item("style")?
                .and_then(|v| v.extract::<String>().ok())
            {
                edge.set_border_style(s);
            }
            if let Some(c) = sub
                .get_item("color")?
                .and_then(|v| v.extract::<String>().ok())
            {
                edge.get_color_mut().set_argb(hex_to_argb(&c));
            }
            Ok(())
        }

        if let Some(sub) = dict.get_item("top")? {
            if let Ok(d) = sub.downcast::<PyDict>() {
                apply_edge(borders.get_top_mut(), d)?;
            }
        }
        if let Some(sub) = dict.get_item("bottom")? {
            if let Ok(d) = sub.downcast::<PyDict>() {
                apply_edge(borders.get_bottom_mut(), d)?;
            }
        }
        if let Some(sub) = dict.get_item("left")? {
            if let Ok(d) = sub.downcast::<PyDict>() {
                apply_edge(borders.get_left_mut(), d)?;
            }
        }
        if let Some(sub) = dict.get_item("right")? {
            if let Ok(d) = sub.downcast::<PyDict>() {
                apply_edge(borders.get_right_mut(), d)?;
            }
        }
        if let Some(sub) = dict.get_item("diagonal_up")? {
            if let Ok(d) = sub.downcast::<PyDict>() {
                apply_edge(borders.get_diagonal_mut(), d)?;
            }
        }
        if let Some(sub) = dict.get_item("diagonal_down")? {
            if let Ok(d) = sub.downcast::<PyDict>() {
                apply_edge(borders.get_diagonal_mut(), d)?;
            }
        }

        Ok(())
    }
}
