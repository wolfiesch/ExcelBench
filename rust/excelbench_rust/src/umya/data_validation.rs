use pyo3::exceptions::PyValueError;
use pyo3::prelude::*;
use pyo3::types::{PyDict, PyList};

use umya_spreadsheet::structs::{
    DataValidation, DataValidationOperatorValues, DataValidationValues, EnumTrait,
};

use super::UmyaBook;

fn dv_type_to_str(t: &DataValidationValues) -> &'static str {
    match t {
        DataValidationValues::Whole => "whole",
        DataValidationValues::Decimal => "decimal",
        DataValidationValues::List => "list",
        DataValidationValues::Date => "date",
        DataValidationValues::Time => "time",
        DataValidationValues::TextLength => "textLength",
        DataValidationValues::Custom => "custom",
        DataValidationValues::None => "none",
    }
}

fn str_to_dv_type(s: &str) -> DataValidationValues {
    match s {
        "whole" => DataValidationValues::Whole,
        "decimal" => DataValidationValues::Decimal,
        "list" => DataValidationValues::List,
        "date" => DataValidationValues::Date,
        "time" => DataValidationValues::Time,
        "textLength" => DataValidationValues::TextLength,
        "custom" => DataValidationValues::Custom,
        _ => DataValidationValues::None,
    }
}

fn dv_op_to_str(op: &DataValidationOperatorValues) -> &'static str {
    match op {
        DataValidationOperatorValues::Between => "between",
        DataValidationOperatorValues::NotBetween => "notBetween",
        DataValidationOperatorValues::Equal => "equal",
        DataValidationOperatorValues::NotEqual => "notEqual",
        DataValidationOperatorValues::GreaterThan => "greaterThan",
        DataValidationOperatorValues::GreaterThanOrEqual => "greaterThanOrEqual",
        DataValidationOperatorValues::LessThan => "lessThan",
        DataValidationOperatorValues::LessThanOrEqual => "lessThanOrEqual",
    }
}

fn str_to_dv_op(s: &str) -> DataValidationOperatorValues {
    match s {
        "between" => DataValidationOperatorValues::Between,
        "notBetween" => DataValidationOperatorValues::NotBetween,
        "equal" => DataValidationOperatorValues::Equal,
        "notEqual" => DataValidationOperatorValues::NotEqual,
        "greaterThan" => DataValidationOperatorValues::GreaterThan,
        "greaterThanOrEqual" => DataValidationOperatorValues::GreaterThanOrEqual,
        "lessThanOrEqual" => DataValidationOperatorValues::LessThanOrEqual,
        _ => DataValidationOperatorValues::LessThan,
    }
}

#[pymethods]
impl UmyaBook {
    pub fn read_data_validations(&self, py: Python<'_>, sheet: &str) -> PyResult<PyObject> {
        let ws = self
            .book
            .get_sheet_by_name(sheet)
            .ok_or_else(|| PyErr::new::<PyValueError, _>(format!("Unknown sheet: {sheet}")))?;

        let result = PyList::empty(py);

        let dvs = match ws.get_data_validations() {
            Some(v) => v,
            None => return Ok(result.into()),
        };

        for dv in dvs.get_data_validation_list() {
            let d = PyDict::new(py);
            d.set_item("range", dv.get_sequence_of_references().get_sqref())?;
            d.set_item("validation_type", dv_type_to_str(dv.get_type()))?;

            let op_str = dv_op_to_str(dv.get_operator());
            let op_has_value = dv.get_operator().get_value_string() != "lessThan"
                || *dv.get_type() != DataValidationValues::List;
            if op_has_value {
                d.set_item("operator", op_str)?;
            } else {
                d.set_item("operator", py.None())?;
            }

            let f1 = dv.get_formula1();
            d.set_item("formula1", if f1.is_empty() { None } else { Some(f1) })?;
            let f2 = dv.get_formula2();
            d.set_item("formula2", if f2.is_empty() { None } else { Some(f2) })?;

            d.set_item("allow_blank", *dv.get_allow_blank())?;
            d.set_item("show_input", *dv.get_show_input_message())?;
            d.set_item("show_error", *dv.get_show_error_message())?;

            let pt = dv.get_prompt_title();
            d.set_item(
                "prompt_title",
                if pt.is_empty() { None } else { Some(pt) },
            )?;
            let p = dv.get_prompt();
            d.set_item("prompt", if p.is_empty() { None } else { Some(p) })?;
            let et = dv.get_error_title();
            d.set_item(
                "error_title",
                if et.is_empty() { None } else { Some(et) },
            )?;
            let em = dv.get_error_message();
            d.set_item("error", if em.is_empty() { None } else { Some(em) })?;

            result.append(d)?;
        }

        Ok(result.into())
    }

    pub fn add_data_validation(
        &mut self,
        sheet: &str,
        validation_dict: &Bound<'_, PyAny>,
    ) -> PyResult<()> {
        let ws = self
            .book
            .get_sheet_by_name_mut(sheet)
            .ok_or_else(|| PyErr::new::<PyValueError, _>(format!("Unknown sheet: {sheet}")))?;

        let dict = validation_dict
            .downcast::<PyDict>()
            .map_err(|_| PyErr::new::<PyValueError, _>("validation must be a dict"))?;

        // Support optional wrapper key "validation"
        let inner: Option<Bound<'_, PyAny>> = dict.get_item("validation")?;
        let cfg: &Bound<'_, PyDict> = match &inner {
            Some(v) => v.downcast::<PyDict>().unwrap_or(dict),
            None => dict,
        };

        let mut dv = DataValidation::default();

        if let Some(vt) = cfg
            .get_item("validation_type")?
            .and_then(|v| v.extract::<String>().ok())
        {
            dv.set_type(str_to_dv_type(&vt));
        }
        if let Some(op) = cfg
            .get_item("operator")?
            .and_then(|v| v.extract::<String>().ok())
        {
            dv.set_operator(str_to_dv_op(&op));
        }
        if let Some(f1) = cfg
            .get_item("formula1")?
            .and_then(|v| v.extract::<String>().ok())
        {
            dv.set_formula1(f1);
        }
        if let Some(f2) = cfg
            .get_item("formula2")?
            .and_then(|v| v.extract::<String>().ok())
        {
            dv.set_formula2(f2);
        }
        if let Some(ab) = cfg
            .get_item("allow_blank")?
            .and_then(|v| v.extract::<bool>().ok())
        {
            dv.set_allow_blank(ab);
        }
        if let Some(si) = cfg
            .get_item("show_input")?
            .and_then(|v| v.extract::<bool>().ok())
        {
            dv.set_show_input_message(si);
        }
        if let Some(se) = cfg
            .get_item("show_error")?
            .and_then(|v| v.extract::<bool>().ok())
        {
            dv.set_show_error_message(se);
        }
        if let Some(pt) = cfg
            .get_item("prompt_title")?
            .and_then(|v| v.extract::<String>().ok())
        {
            dv.set_prompt_title(pt);
        }
        if let Some(p) = cfg
            .get_item("prompt")?
            .and_then(|v| v.extract::<String>().ok())
        {
            dv.set_prompt(p);
        }
        if let Some(et) = cfg
            .get_item("error_title")?
            .and_then(|v| v.extract::<String>().ok())
        {
            dv.set_error_title(et);
        }
        if let Some(e) = cfg
            .get_item("error")?
            .and_then(|v| v.extract::<String>().ok())
        {
            dv.set_error_message(e);
        }

        // Set range
        if let Some(range) = cfg
            .get_item("range")?
            .and_then(|v| v.extract::<String>().ok())
        {
            dv.get_sequence_of_references_mut().set_sqref(range);
        }

        // Add to worksheet
        let mut dvs = ws.get_data_validations().cloned().unwrap_or_default();
        dvs.add_data_validation_list(dv);
        ws.set_data_validations(dvs);

        Ok(())
    }
}
