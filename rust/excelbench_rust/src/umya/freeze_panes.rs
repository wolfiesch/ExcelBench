use pyo3::exceptions::PyValueError;
use pyo3::prelude::*;
use pyo3::types::PyDict;

use umya_spreadsheet::structs::{EnumTrait, Pane, PaneStateValues, PaneValues, SheetView};

use super::UmyaBook;

/// Extract a string value from a PyDict, looking in an optional inner dict first.
fn get_str(dict: &Bound<'_, PyDict>, key: &str) -> PyResult<Option<String>> {
    dict.get_item(key)?
        .map(|v| v.extract::<String>())
        .transpose()
}

fn get_f64(dict: &Bound<'_, PyDict>, key: &str) -> PyResult<Option<f64>> {
    Ok(dict.get_item(key)?.and_then(|v| v.extract::<f64>().ok()))
}

#[pymethods]
impl UmyaBook {
    pub fn read_freeze_panes(&self, py: Python<'_>, sheet: &str) -> PyResult<PyObject> {
        let ws = self
            .book
            .get_sheet_by_name(sheet)
            .ok_or_else(|| PyErr::new::<PyValueError, _>(format!("Unknown sheet: {sheet}")))?;

        let d = PyDict::new(py);

        let views = ws.get_sheets_views();
        let view_list = views.get_sheet_view_list();
        let sv = match view_list.first() {
            Some(v) => v,
            None => return Ok(d.into()),
        };

        let pane = match sv.get_pane() {
            Some(p) => p,
            None => return Ok(d.into()),
        };

        match pane.get_state() {
            PaneStateValues::Frozen | PaneStateValues::FrozenSplit => {
                d.set_item("mode", "freeze")?;
                let tlc = pane.get_top_left_cell().to_string();
                if !tlc.is_empty() {
                    d.set_item("top_left_cell", tlc)?;
                }
            }
            PaneStateValues::Split => {
                let x = *pane.get_horizontal_split();
                let y = *pane.get_vertical_split();
                if x == 0.0 && y == 0.0 {
                    return Ok(d.into());
                }
                d.set_item("mode", "split")?;
                d.set_item("x_split", x as i64)?;
                d.set_item("y_split", y as i64)?;
                let tlc = pane.get_top_left_cell().to_string();
                if !tlc.is_empty() {
                    d.set_item("top_left_cell", tlc)?;
                }
                d.set_item("active_pane", pane.get_active_pane().get_value_string())?;
            }
        }

        Ok(d.into())
    }

    pub fn set_freeze_panes(
        &mut self,
        sheet: &str,
        settings: &Bound<'_, PyAny>,
    ) -> PyResult<()> {
        let ws = self
            .book
            .get_sheet_by_name_mut(sheet)
            .ok_or_else(|| PyErr::new::<PyValueError, _>(format!("Unknown sheet: {sheet}")))?;

        let dict = settings
            .downcast::<PyDict>()
            .map_err(|_| PyErr::new::<PyValueError, _>("settings must be a dict"))?;

        // Support optional wrapper key "freeze" — unwrap into owned values
        let inner: Option<Bound<'_, PyAny>> = dict.get_item("freeze")?;
        let cfg: &Bound<'_, PyDict> = match &inner {
            Some(v) => v.downcast::<PyDict>().unwrap_or(dict),
            None => dict,
        };

        let mode = get_str(cfg, "mode")?.unwrap_or_default();

        let views = ws.get_sheet_views_mut();
        let view_list = views.get_sheet_view_list_mut();
        if view_list.is_empty() {
            view_list.push(SheetView::default());
        }
        let sv = &mut view_list[0];

        let mut pane = Pane::default();

        if mode == "freeze" {
            pane.set_state(PaneStateValues::Frozen);
            if let Some(s) = get_str(cfg, "top_left_cell")? {
                pane.get_top_left_cell_mut().set_coordinate(&s);
                let (row, col) = parse_top_left(&s);
                if col > 0 {
                    pane.set_horizontal_split(col as f64);
                }
                if row > 0 {
                    pane.set_vertical_split(row as f64);
                }
            }
            pane.set_active_pane(PaneValues::BottomRight);
        } else if mode == "split" {
            pane.set_state(PaneStateValues::Split);
            if let Some(x) = get_f64(cfg, "x_split")? {
                pane.set_horizontal_split(x);
            }
            if let Some(y) = get_f64(cfg, "y_split")? {
                pane.set_vertical_split(y);
            }
            if let Some(s) = get_str(cfg, "top_left_cell")? {
                pane.get_top_left_cell_mut().set_coordinate(&s);
            }
            if let Some(ap) = get_str(cfg, "active_pane")? {
                let pv = match ap.as_str() {
                    "bottomLeft" => PaneValues::BottomLeft,
                    "topRight" => PaneValues::TopRight,
                    "topLeft" => PaneValues::TopLeft,
                    _ => PaneValues::BottomRight,
                };
                pane.set_active_pane(pv);
            }
        }

        sv.set_pane(pane);
        Ok(())
    }
}

/// Parse "B2" → (row_offset=1, col_offset=1) for freeze splits.
fn parse_top_left(a1: &str) -> (u32, u32) {
    let mut col: u32 = 0;
    let mut row: u32 = 0;
    let mut in_digits = false;
    for ch in a1.chars() {
        if ch.is_ascii_alphabetic() && !in_digits {
            col = col * 26 + (ch.to_ascii_uppercase() as u32 - b'A' as u32 + 1);
        } else if ch.is_ascii_digit() {
            in_digits = true;
            row = row * 10 + ch.to_digit(10).unwrap();
        }
    }
    (row.saturating_sub(1), col.saturating_sub(1))
}
