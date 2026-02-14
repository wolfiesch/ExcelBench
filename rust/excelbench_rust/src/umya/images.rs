use pyo3::exceptions::PyValueError;
use pyo3::prelude::*;
use pyo3::types::{PyDict, PyList};

use umya_spreadsheet::structs::drawing::spreadsheet::MarkerType;
use umya_spreadsheet::structs::Image;

use super::UmyaBook;

#[pymethods]
impl UmyaBook {
    pub fn read_images(&self, py: Python<'_>, sheet: &str) -> PyResult<PyObject> {
        let ws = self
            .book
            .get_sheet_by_name(sheet)
            .ok_or_else(|| PyErr::new::<PyValueError, _>(format!("Unknown sheet: {sheet}")))?;

        let result = PyList::empty(py);

        for img in ws.get_image_collection() {
            let d = PyDict::new(py);

            // Determine anchor type and cell from the image's anchor
            if let Some(two_cell) = img.get_two_cell_anchor() {
                let from = two_cell.get_from_marker();
                d.set_item("cell", from.get_coordinate())?;
                d.set_item("anchor", "twoCell")?;
                let offsets = PyList::new(py, [*from.get_col_off(), *from.get_row_off()])?;
                d.set_item("offset", offsets)?;
            } else if let Some(one_cell) = img.get_one_cell_anchor() {
                let from = one_cell.get_from_marker();
                d.set_item("cell", from.get_coordinate())?;
                d.set_item("anchor", "oneCell")?;
                let offsets = PyList::new(py, [*from.get_col_off(), *from.get_row_off()])?;
                d.set_item("offset", offsets)?;
            } else {
                d.set_item("cell", py.None())?;
                d.set_item("anchor", py.None())?;
                d.set_item("offset", py.None())?;
            }

            // Path/media reference not directly exposed in umya — set to None
            d.set_item("path", py.None())?;
            d.set_item("alt_text", py.None())?;

            result.append(d)?;
        }

        Ok(result.into())
    }

    pub fn add_image(
        &mut self,
        sheet: &str,
        image_dict: &Bound<'_, PyAny>,
    ) -> PyResult<()> {
        let ws = self
            .book
            .get_sheet_by_name_mut(sheet)
            .ok_or_else(|| PyErr::new::<PyValueError, _>(format!("Unknown sheet: {sheet}")))?;

        let dict = image_dict
            .downcast::<PyDict>()
            .map_err(|_| PyErr::new::<PyValueError, _>("image must be a dict"))?;

        // Support optional wrapper key "image"
        let inner: Option<Bound<'_, PyAny>> = dict.get_item("image")?;
        let cfg: &Bound<'_, PyDict> = match &inner {
            Some(v) => v.downcast::<PyDict>().unwrap_or(dict),
            None => dict,
        };

        let path: String = cfg
            .get_item("path")?
            .ok_or_else(|| PyErr::new::<PyValueError, _>("image missing 'path'"))?
            .extract()?;
        let cell: String = cfg
            .get_item("cell")?
            .ok_or_else(|| PyErr::new::<PyValueError, _>("image missing 'cell'"))?
            .extract()?;

        let mut marker = MarkerType::default();
        marker.set_coordinate(cell);

        let mut image = Image::default();
        image.new_image(&path, marker);
        ws.add_image(image);

        Ok(())
    }
}
