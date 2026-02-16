use pyo3::prelude::*;
use pyo3::types::{PyDict, PyList};

#[cfg(any(feature = "calamine", feature = "rust_xlsxwriter", feature = "umya", feature = "wolfxl"))]
mod util;

#[cfg(any(feature = "calamine", feature = "rust_xlsxwriter", feature = "wolfxl"))]
mod ooxml_util;

#[cfg(feature = "calamine")]
mod calamine_backend;

#[cfg(feature = "calamine")]
mod calamine_styled_backend;

#[cfg(feature = "rust_xlsxwriter")]
mod rust_xlsxwriter_backend;

#[cfg(feature = "umya")]
mod umya;

#[cfg(feature = "wolfxl")]
mod wolfxl;

fn enabled_backends() -> Vec<&'static str> {
    let mut out: Vec<&'static str> = Vec::new();
    if cfg!(feature = "calamine") {
        out.push("calamine");
    }
    if cfg!(feature = "rust_xlsxwriter") {
        out.push("rust_xlsxwriter");
    }
    if cfg!(feature = "umya") {
        out.push("umya-spreadsheet");
    }
    if cfg!(feature = "wolfxl") {
        out.push("wolfxl");
    }
    out
}

#[pyfunction]
fn build_info(py: Python<'_>) -> PyResult<PyObject> {
    // Stable keys so Python adapters can depend on this shape.
    let info = PyDict::new(py);
    info.set_item("package", "excelbench_rust")?;
    info.set_item("package_version", env!("CARGO_PKG_VERSION"))?;

    let enabled = enabled_backends();
    info.set_item("enabled_backends", PyList::new(py, enabled)?)?;

    // Backend version reporting can be filled in later.
    // Keep a dict in place now so consumers can read it unconditionally.
    let backends = PyDict::new(py);
    backends.set_item(
        "calamine",
        if cfg!(feature = "calamine") {
            "enabled"
        } else {
            "disabled"
        },
    )?;
    backends.set_item(
        "rust_xlsxwriter",
        if cfg!(feature = "rust_xlsxwriter") {
            "enabled"
        } else {
            "disabled"
        },
    )?;
    backends.set_item(
        "umya-spreadsheet",
        if cfg!(feature = "umya") {
            "enabled"
        } else {
            "disabled"
        },
    )?;
    info.set_item("backends", backends)?;

    let versions = PyDict::new(py);
    versions.set_item("calamine", option_env!("EXCELBENCH_DEP_CALAMINE_VERSION"))?;
    versions.set_item(
        "rust_xlsxwriter",
        option_env!("EXCELBENCH_DEP_RUST_XLSXWRITER_VERSION"),
    )?;
    versions.set_item(
        "umya-spreadsheet",
        option_env!("EXCELBENCH_DEP_UMYA_VERSION"),
    )?;
    info.set_item("backend_versions", versions)?;

    Ok(info.into())
}

#[pymodule]
fn excelbench_rust(m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add("__version__", env!("CARGO_PKG_VERSION"))?;
    m.add_function(wrap_pyfunction!(build_info, m)?)?;

    #[cfg(feature = "calamine")]
    {
        m.add_class::<calamine_backend::CalamineBook>()?;
        m.add_class::<calamine_styled_backend::CalamineStyledBook>()?;
    }

    #[cfg(feature = "rust_xlsxwriter")]
    {
        m.add_class::<rust_xlsxwriter_backend::RustXlsxWriterBook>()?;
    }

    #[cfg(feature = "umya")]
    {
        m.add_class::<umya::UmyaBook>()?;
    }

    #[cfg(feature = "wolfxl")]
    {
        m.add_class::<wolfxl::XlsxPatcher>()?;
    }

    Ok(())
}
