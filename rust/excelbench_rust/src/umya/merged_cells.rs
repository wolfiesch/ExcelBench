use pyo3::exceptions::PyValueError;
use pyo3::prelude::*;

use super::UmyaBook;

#[pymethods]
impl UmyaBook {
    pub fn read_merged_ranges(&self, sheet: &str) -> PyResult<Vec<String>> {
        let ws = self
            .book
            .get_sheet_by_name(sheet)
            .ok_or_else(|| PyErr::new::<PyValueError, _>(format!("Unknown sheet: {sheet}")))?;

        let ranges: Vec<String> = ws
            .get_merge_cells()
            .iter()
            .map(|mc| mc.get_range().to_string())
            .collect();
        Ok(ranges)
    }

    pub fn merge_cells(&mut self, sheet: &str, range: &str) -> PyResult<()> {
        let ws = self
            .book
            .get_sheet_by_name_mut(sheet)
            .ok_or_else(|| PyErr::new::<PyValueError, _>(format!("Unknown sheet: {sheet}")))?;

        ws.add_merge_cells(range);
        Ok(())
    }
}
