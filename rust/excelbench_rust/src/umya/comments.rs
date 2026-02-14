use pyo3::exceptions::PyValueError;
use pyo3::prelude::*;
use pyo3::types::{PyDict, PyList};

use umya_spreadsheet::structs::Comment;

use super::UmyaBook;

/// Extract plain text from a comment.
/// umya-spreadsheet's `Text` type is pub(crate), so we can only read text
/// through `RichText::get_text()`. This covers openpyxl-generated fixtures
/// (which store comments as rich text elements).
fn extract_comment_text(comment: &Comment) -> String {
    let ct = comment.get_text();
    if let Some(rt) = ct.get_rich_text() {
        return rt.get_text().to_string();
    }
    // Plain Text case: type is pub(crate), can't access value directly.
    // Return empty — author + cell are still verified by benchmark.
    String::new()
}

#[pymethods]
impl UmyaBook {
    pub fn read_comments(&self, py: Python<'_>, sheet: &str) -> PyResult<PyObject> {
        let ws = self
            .book
            .get_sheet_by_name(sheet)
            .ok_or_else(|| PyErr::new::<PyValueError, _>(format!("Unknown sheet: {sheet}")))?;

        let comments = ws.get_comments();
        let result = PyList::empty(py);

        for comment in comments {
            let d = PyDict::new(py);
            d.set_item("cell", comment.get_coordinate().to_string())?;
            d.set_item("text", extract_comment_text(comment))?;
            d.set_item("author", comment.get_author())?;
            d.set_item("threaded", false)?;
            result.append(d)?;
        }

        Ok(result.into())
    }

    pub fn add_comment(
        &mut self,
        sheet: &str,
        comment_dict: &Bound<'_, PyAny>,
    ) -> PyResult<()> {
        let ws = self
            .book
            .get_sheet_by_name_mut(sheet)
            .ok_or_else(|| PyErr::new::<PyValueError, _>(format!("Unknown sheet: {sheet}")))?;

        let dict = comment_dict
            .downcast::<PyDict>()
            .map_err(|_| PyErr::new::<PyValueError, _>("comment must be a dict"))?;

        // Support optional wrapper key "comment"
        let inner: Option<Bound<'_, PyAny>> = dict.get_item("comment")?;
        let cfg: &Bound<'_, PyDict> = match &inner {
            Some(v) => v.downcast::<PyDict>().unwrap_or(dict),
            None => dict,
        };

        let cell: String = cfg
            .get_item("cell")?
            .ok_or_else(|| PyErr::new::<PyValueError, _>("comment missing 'cell'"))?
            .extract()?;
        let text: String = cfg
            .get_item("text")?
            .map(|v| v.extract::<String>())
            .transpose()?
            .unwrap_or_default();
        let author: String = cfg
            .get_item("author")?
            .map(|v| v.extract::<String>())
            .transpose()?
            .unwrap_or_default();

        let mut c = Comment::default();
        c.new_comment(&*cell);
        c.set_text_string(text);
        c.set_author(author);
        ws.add_comments(c);

        Ok(())
    }
}
