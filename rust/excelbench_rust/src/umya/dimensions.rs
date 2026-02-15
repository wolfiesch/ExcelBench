use pyo3::exceptions::PyValueError;
use pyo3::prelude::*;

use super::util::col_letter_to_u32;
use super::UmyaBook;

#[pymethods]
impl UmyaBook {
    pub fn read_row_height(&self, sheet: &str, row: u32) -> PyResult<Option<f64>> {
        let ws = self
            .book
            .get_sheet_by_name(sheet)
            .ok_or_else(|| PyErr::new::<PyValueError, _>(format!("Unknown sheet: {sheet}")))?;

        // umya uses 1-based row index.
        if let Some(rd) = ws.get_row_dimension(&(row + 1)) {
            let h = rd.get_height();
            if h > &0.0 {
                return Ok(Some(*h));
            }
        }
        Ok(None)
    }

    pub fn read_column_width(&self, sheet: &str, col_str: &str) -> PyResult<Option<f64>> {
        let ws = self
            .book
            .get_sheet_by_name(sheet)
            .ok_or_else(|| PyErr::new::<PyValueError, _>(format!("Unknown sheet: {sheet}")))?;

        let col_idx = col_letter_to_u32(col_str).map_err(|e| PyErr::new::<PyValueError, _>(e))?;

        if let Some(cd) = ws.get_column_dimension_by_number(&col_idx) {
            let w = cd.get_width();
            if w > &0.0 {
                return Ok(Some(*w));
            }
        }
        Ok(None)
    }

    pub fn set_row_height(&mut self, sheet: &str, row: u32, height: f64) -> PyResult<()> {
        let ws = self
            .book
            .get_sheet_by_name_mut(sheet)
            .ok_or_else(|| PyErr::new::<PyValueError, _>(format!("Unknown sheet: {sheet}")))?;

        // umya uses 1-based row index.
        ws.get_row_dimension_mut(&(row + 1)).set_height(height);
        Ok(())
    }

    pub fn set_column_width(&mut self, sheet: &str, col_str: &str, width: f64) -> PyResult<()> {
        let ws = self
            .book
            .get_sheet_by_name_mut(sheet)
            .ok_or_else(|| PyErr::new::<PyValueError, _>(format!("Unknown sheet: {sheet}")))?;

        let col_idx = col_letter_to_u32(col_str).map_err(|e| PyErr::new::<PyValueError, _>(e))?;

        ws.get_column_dimension_by_number_mut(&col_idx)
            .set_width(width);
        Ok(())
    }
}
