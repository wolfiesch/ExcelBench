use pyo3::exceptions::PyValueError;
use pyo3::prelude::*;
use pyo3::types::{PyDict, PyList};

use umya_spreadsheet::structs::{
    Color, ConditionalFormatValues, ConditionalFormatting, ConditionalFormattingOperatorValues,
    ConditionalFormattingRule, EnumTrait, Formula, Style,
};

use super::UmyaBook;

fn cf_type_to_str(t: &ConditionalFormatValues) -> &str {
    t.get_value_string()
}

fn str_to_cf_type(s: &str) -> ConditionalFormatValues {
    match s {
        "cellIs" => ConditionalFormatValues::CellIs,
        "expression" => ConditionalFormatValues::Expression,
        "colorScale" => ConditionalFormatValues::ColorScale,
        "dataBar" => ConditionalFormatValues::DataBar,
        "iconSet" => ConditionalFormatValues::IconSet,
        "top10" => ConditionalFormatValues::Top10,
        "aboveAverage" => ConditionalFormatValues::AboveAverage,
        "beginsWith" => ConditionalFormatValues::BeginsWith,
        "endsWith" => ConditionalFormatValues::EndsWith,
        "containsText" => ConditionalFormatValues::ContainsText,
        "notContainsText" => ConditionalFormatValues::NotContainsText,
        "containsBlanks" => ConditionalFormatValues::ContainsBlanks,
        "notContainsBlanks" => ConditionalFormatValues::NotContainsBlanks,
        "containsErrors" => ConditionalFormatValues::ContainsErrors,
        "notContainsErrors" => ConditionalFormatValues::NotContainsErrors,
        "duplicateValues" => ConditionalFormatValues::DuplicateValues,
        "uniqueValues" => ConditionalFormatValues::UniqueValues,
        "timePeriod" => ConditionalFormatValues::TimePeriod,
        _ => ConditionalFormatValues::Expression,
    }
}

fn cf_op_to_str(op: &ConditionalFormattingOperatorValues) -> &str {
    op.get_value_string()
}

fn str_to_cf_op(s: &str) -> ConditionalFormattingOperatorValues {
    match s {
        "between" => ConditionalFormattingOperatorValues::Between,
        "notBetween" => ConditionalFormattingOperatorValues::NotBetween,
        "equal" => ConditionalFormattingOperatorValues::Equal,
        "notEqual" => ConditionalFormattingOperatorValues::NotEqual,
        "greaterThan" => ConditionalFormattingOperatorValues::GreaterThan,
        "greaterThanOrEqual" => ConditionalFormattingOperatorValues::GreaterThanOrEqual,
        "lessThan" => ConditionalFormattingOperatorValues::LessThan,
        "lessThanOrEqual" => ConditionalFormattingOperatorValues::LessThanOrEqual,
        "beginsWith" => ConditionalFormattingOperatorValues::BeginsWith,
        "endsWith" => ConditionalFormattingOperatorValues::EndsWith,
        "containsText" => ConditionalFormattingOperatorValues::ContainsText,
        "notContains" => ConditionalFormattingOperatorValues::NotContains,
        _ => ConditionalFormattingOperatorValues::LessThan,
    }
}

fn argb_to_hex(color: &Color) -> Option<String> {
    let argb = color.get_argb();
    if argb.is_empty() || argb == "00000000" {
        return None;
    }
    // Convert ARGB "AARRGGBB" → "#RRGGBB" (strip alpha, add #)
    let rgb = if argb.len() == 8 { &argb[2..] } else { argb };
    Some(format!("#{rgb}"))
}

#[pymethods]
impl UmyaBook {
    pub fn read_conditional_formats(&self, py: Python<'_>, sheet: &str) -> PyResult<PyObject> {
        let ws = self
            .book
            .get_sheet_by_name(sheet)
            .ok_or_else(|| PyErr::new::<PyValueError, _>(format!("Unknown sheet: {sheet}")))?;

        let result = PyList::empty(py);

        for cf in ws.get_conditional_formatting_collection() {
            let range = cf.get_sequence_of_references().get_sqref();

            for rule in cf.get_conditional_collection() {
                let d = PyDict::new(py);
                d.set_item("range", &range)?;
                d.set_item("rule_type", cf_type_to_str(rule.get_type()))?;

                let op_str = cf_op_to_str(rule.get_operator());
                let has_op = op_str != "lessThan"
                    || *rule.get_type() == ConditionalFormatValues::CellIs;
                if has_op {
                    d.set_item("operator", op_str)?;
                } else {
                    d.set_item("operator", py.None())?;
                }

                // Formula
                if let Some(formula) = rule.get_formula() {
                    let text = formula.get_address_str();
                    if !text.is_empty() {
                        d.set_item("formula", &text)?;
                    } else {
                        d.set_item("formula", py.None())?;
                    }
                } else {
                    d.set_item("formula", py.None())?;
                }

                let priority = *rule.get_priority();
                if priority != 0 {
                    d.set_item("priority", priority)?;
                } else {
                    d.set_item("priority", py.None())?;
                }
                let sit = *rule.get_stop_if_true();
                if sit {
                    d.set_item("stop_if_true", true)?;
                } else {
                    d.set_item("stop_if_true", py.None())?;
                }

                // Format (bg_color, font_color)
                let fmt = PyDict::new(py);
                if let Some(style) = rule.get_style() {
                    if let Some(bg) = style.get_background_color() {
                        if let Some(hex) = argb_to_hex(bg) {
                            fmt.set_item("bg_color", hex)?;
                        }
                    }
                    if let Some(font) = style.get_font() {
                        let fc = font.get_color();
                        if let Some(hex) = argb_to_hex(fc) {
                            fmt.set_item("font_color", hex)?;
                        }
                    }
                }
                d.set_item("format", fmt)?;

                result.append(d)?;
            }
        }

        Ok(result.into())
    }

    pub fn add_conditional_format(
        &mut self,
        sheet: &str,
        rule_dict: &Bound<'_, PyAny>,
    ) -> PyResult<()> {
        let ws = self
            .book
            .get_sheet_by_name_mut(sheet)
            .ok_or_else(|| PyErr::new::<PyValueError, _>(format!("Unknown sheet: {sheet}")))?;

        let dict = rule_dict
            .downcast::<PyDict>()
            .map_err(|_| PyErr::new::<PyValueError, _>("rule must be a dict"))?;

        // Support optional wrapper key "conditional_format"
        let inner: Option<Bound<'_, PyAny>> = dict.get_item("conditional_format")?;
        let cfg: &Bound<'_, PyDict> = match &inner {
            Some(v) => v.downcast::<PyDict>().unwrap_or(dict),
            None => dict,
        };

        let mut rule = ConditionalFormattingRule::default();

        if let Some(rt) = cfg
            .get_item("rule_type")?
            .and_then(|v| v.extract::<String>().ok())
        {
            rule.set_type(str_to_cf_type(&rt));
        }
        if let Some(op) = cfg
            .get_item("operator")?
            .and_then(|v| v.extract::<String>().ok())
        {
            rule.set_operator(str_to_cf_op(&op));
        }
        if let Some(f) = cfg
            .get_item("formula")?
            .and_then(|v| v.extract::<String>().ok())
        {
            let mut formula = Formula::default();
            formula.set_string_value(f);
            rule.set_formula(formula);
        }
        if let Some(p) = cfg
            .get_item("priority")?
            .and_then(|v| v.extract::<i32>().ok())
        {
            rule.set_priority(p);
        }
        if let Some(sit) = cfg
            .get_item("stop_if_true")?
            .and_then(|v| v.extract::<bool>().ok())
        {
            rule.set_stop_if_true(sit);
        }

        // Format: bg_color, font_color
        if let Some(fmt_val) = cfg.get_item("format")? {
            if let Ok(fmt_dict) = fmt_val.downcast::<PyDict>() {
                let mut style = Style::default();
                if let Some(bg) = fmt_dict
                    .get_item("bg_color")?
                    .and_then(|v| v.extract::<String>().ok())
                {
                    let bg = bg.strip_prefix('#').unwrap_or(&bg);
                    style.set_background_color(bg);
                }
                if let Some(fc) = fmt_dict
                    .get_item("font_color")?
                    .and_then(|v| v.extract::<String>().ok())
                {
                    let fc = fc.strip_prefix('#').unwrap_or(&fc);
                    let font = style.get_font_mut();
                    font.get_color_mut().set_argb(fc);
                }
                rule.set_style(style);
            }
        }

        // Build ConditionalFormatting container
        let mut cf = ConditionalFormatting::default();
        if let Some(range) = cfg
            .get_item("range")?
            .and_then(|v| v.extract::<String>().ok())
        {
            cf.get_sequence_of_references_mut().set_sqref(range);
        }
        cf.add_conditional_collection(rule);
        ws.add_conditional_formatting_collection(cf);

        Ok(())
    }
}
