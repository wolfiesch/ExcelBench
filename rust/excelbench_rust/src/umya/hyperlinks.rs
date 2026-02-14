use pyo3::exceptions::PyValueError;
use pyo3::prelude::*;
use pyo3::types::{PyDict, PyList};

use umya_spreadsheet::structs::Hyperlink;

use super::UmyaBook;

#[pymethods]
impl UmyaBook {
    pub fn read_hyperlinks(&self, py: Python<'_>, sheet: &str) -> PyResult<PyObject> {
        let ws = self
            .book
            .get_sheet_by_name(sheet)
            .ok_or_else(|| PyErr::new::<PyValueError, _>(format!("Unknown sheet: {sheet}")))?;

        let result = PyList::empty(py);

        for cell in ws.get_cell_collection() {
            if let Some(h) = cell.get_hyperlink() {
                let url = h.get_url();
                if url.is_empty() {
                    continue;
                }
                let d = PyDict::new(py);
                d.set_item("cell", cell.get_coordinate().to_string())?;
                let is_location = *h.get_location();
                d.set_item("target", url)?;
                d.set_item("display", cell.get_value().into_owned())?;
                let tooltip = h.get_tooltip();
                if tooltip.is_empty() {
                    d.set_item("tooltip", py.None())?;
                } else {
                    d.set_item("tooltip", tooltip)?;
                }
                d.set_item("internal", is_location)?;
                result.append(d)?;
            }
        }

        Ok(result.into())
    }

    pub fn add_hyperlink(
        &mut self,
        sheet: &str,
        link_dict: &Bound<'_, PyAny>,
    ) -> PyResult<()> {
        let ws = self
            .book
            .get_sheet_by_name_mut(sheet)
            .ok_or_else(|| PyErr::new::<PyValueError, _>(format!("Unknown sheet: {sheet}")))?;

        let dict = link_dict
            .downcast::<PyDict>()
            .map_err(|_| PyErr::new::<PyValueError, _>("link must be a dict"))?;

        // Support optional wrapper key "hyperlink"
        let inner: Option<Bound<'_, PyAny>> = dict.get_item("hyperlink")?;
        let cfg: &Bound<'_, PyDict> = match &inner {
            Some(v) => v.downcast::<PyDict>().unwrap_or(dict),
            None => dict,
        };

        let cell: String = cfg
            .get_item("cell")?
            .ok_or_else(|| PyErr::new::<PyValueError, _>("hyperlink missing 'cell'"))?
            .extract()?;
        let target: String = cfg
            .get_item("target")?
            .ok_or_else(|| PyErr::new::<PyValueError, _>("hyperlink missing 'target'"))?
            .extract()?;
        let display: Option<String> = cfg
            .get_item("display")?
            .and_then(|v| v.extract::<String>().ok());
        let tooltip: Option<String> = cfg
            .get_item("tooltip")?
            .and_then(|v| v.extract::<String>().ok());
        let internal: bool = cfg
            .get_item("internal")?
            .and_then(|v| v.extract::<bool>().ok())
            .unwrap_or(false);

        // Set cell value to display text
        if let Some(disp) = &display {
            ws.get_cell_mut(&*cell).set_value_string(disp);
        }

        // Create and attach hyperlink
        let mut h = Hyperlink::default();
        h.set_url(&target);
        h.set_location(internal);
        if let Some(tt) = &tooltip {
            h.set_tooltip(tt);
        }
        ws.get_cell_mut(&*cell).set_hyperlink(h);

        Ok(())
    }
}
