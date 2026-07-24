//! xlsb_reader_rs — Rust backend for the `xlsb_reader` Python package.
//!
//! This crate is compiled as a Python extension module (via `pyo3`/
//! `maturin`) that provides fast `.xlsb`/`.xlsx`/`.xlsm` parsing as a
//! drop-in replacement for `xlsb_reader`'s pure-Python implementation.
//! `xlsb_reader/__init__.py` imports this module (if installed) and
//! re-exports its `XlsbWorkbook`/`XlsxWorkbook`/`col_to_letter`, falling
//! back to the pure-Python versions otherwise — see that file's backend
//! dispatch logic for the exact contract this module must satisfy.
//!
//! # Module layout
//!
//!   - [`common`] — shared data model, byte-cursor `Reader`, error type,
//!     and PyO3 conversion helpers. **Read this module first** — it is the
//!     contract every other module builds on.
//!   - [`xlsb`] — `.xlsb` parsing + the `XlsbWorkbook` pyclass (Wave 1A).
//!   - [`xlsx`] — `.xlsx`/`.xlsm` parsing + the `XlsxWorkbook` pyclass
//!     (Wave 1B).
//!   - [`vba`] — VBA module extraction from embedded OLE containers
//!     (Wave 1C).
//!   - [`render`] — `to_dict`/`to_json`/`to_markdown` assembly (Wave 1D).
//!   - [`spec`] — the `extract_spec` structural-spec pipeline (Wave 1D).
//!
//! Everything below Wave 0 (all five modules above) currently contains
//! working stubs: parsing always returns an empty [`common::WorkbookData`],
//! and the one-shot render/spec functions return empty results. This is
//! intentional — Wave 0's job is to get the *shape* of the crate, its
//! public API, and the Python-side backend dispatch correct and building
//! end-to-end; later waves fill in the actual parsing/rendering logic
//! behind these exact signatures.

mod common;
mod render;
mod spec;
mod vba;
mod xlsb;
mod xlsx;

use pyo3::prelude::*;

use common::WorkbookData;

fn default_include() -> Vec<String> {
    render::DEFAULT_INCLUDE
        .iter()
        .map(|s| s.to_string())
        .collect()
}

fn is_xlsx_like(path: &std::path::Path) -> bool {
    match path.extension().and_then(|e| e.to_str()) {
        Some(ext) => {
            let ext = ext.to_lowercase();
            ext == "xlsx" || ext == "xlsm"
        }
        None => false,
    }
}

/// Parse `path` with whichever of [`xlsb::parse_xlsb`]/[`xlsx::parse_xlsx`]
/// matches its extension — mirrors
/// `xlsb_reader._spec_extractor.open_workbook`'s dispatch rule (`.xlsx`/
/// `.xlsm` -> xlsx reader, everything else -> xlsb reader).
fn parse_any(path: &str) -> common::Result<WorkbookData> {
    let path = std::path::Path::new(path);
    if is_xlsx_like(path) {
        xlsx::parse_xlsx(path)
    } else {
        xlsb::parse_xlsb(path)
    }
}

/// Convert a 0-based column index to Excel column letter(s), e.g.
/// `0 -> "A"`, `26 -> "AA"`. Matches `xlsb_reader.col_to_letter` (and
/// `xlsb_reader._reader.col_to_letter`/`xlsb_reader._xlsx_reader.col_to_letter`)
/// exactly.
#[pyfunction]
fn col_to_letter(col: u32) -> String {
    common::col_to_letter(col)
}

/// One-shot: open the workbook at `path` and return a `dict` of the
/// requested sections. Fast-path implementation of
/// `xlsb_reader._render.to_dict`, used by `xlsb_reader.to_dict()` in
/// `__init__.py` when this extension is installed.
#[pyfunction]
#[pyo3(signature = (path, include=None, sheet=None))]
fn to_dict(
    py: Python<'_>,
    path: String,
    include: Option<Vec<String>>,
    sheet: Option<String>,
) -> PyResult<Py<PyAny>> {
    let data = parse_any(&path)?;
    let include = include.unwrap_or_else(default_include);
    let supports_vba = is_xlsx_like(std::path::Path::new(&path));
    render::to_dict_py(py, &data, &include, sheet.as_deref(), supports_vba)
}

/// One-shot: same sections as [`to_dict`], returned as a JSON string.
/// Fast-path implementation of `xlsb_reader._render.to_json`.
#[pyfunction]
#[pyo3(signature = (path, include=None, sheet=None))]
fn to_json(path: String, include: Option<Vec<String>>, sheet: Option<String>) -> PyResult<String> {
    let data = parse_any(&path)?;
    let include = include.unwrap_or_else(default_include);
    let supports_vba = is_xlsx_like(std::path::Path::new(&path));
    let value = render::to_dict_value(&data, &include, sheet.as_deref(), supports_vba);
    // Match Python's `json.dumps(data, ensure_ascii=False, indent=2,
    // sort_keys=True)`: 2-space pretty-printing (serde_json's default for
    // `to_string_pretty`) with raw (non-escaped) UTF-8 (serde_json's
    // default) and keys already sorted since `Map`/`Value::Object` is a
    // `BTreeMap` under the hood (the `preserve_order` feature is not
    // enabled — see Cargo.toml).
    serde_json::to_string_pretty(&value)
        .map_err(|e| common::XlsbError::Parse(e.to_string()).into())
}

/// One-shot: same sections as [`to_dict`], rendered as a Markdown report.
/// Fast-path implementation of `xlsb_reader._render.to_markdown`.
#[pyfunction]
#[pyo3(signature = (path, include=None, sheet=None))]
fn to_markdown(
    path: String,
    include: Option<Vec<String>>,
    sheet: Option<String>,
) -> PyResult<String> {
    let data = parse_any(&path)?;
    let include = include.unwrap_or_else(default_include);
    let supports_vba = is_xlsx_like(std::path::Path::new(&path));
    Ok(render::to_markdown_string(&data, &include, sheet.as_deref(), supports_vba))
}

/// One-shot: extract a token-efficient structural spec from the workbook
/// at `path`. Fast-path implementation of
/// `xlsb_reader._spec_extractor.build_spec`, used by
/// `xlsb_reader.extract_spec()` in `__init__.py`.
#[pyfunction]
#[pyo3(signature = (path, sample_rows=50, sheets=None))]
fn extract_spec(path: String, sample_rows: usize, sheets: Option<String>) -> PyResult<String> {
    let data = parse_any(&path)?;
    let sheets = sheets.unwrap_or_else(|| "all".to_string());
    let file_path = std::path::Path::new(&path);
    let file_name = file_path
        .file_name()
        .and_then(|s| s.to_str())
        .unwrap_or(&path)
        .to_string();
    let format_ext = file_path
        .extension()
        .and_then(|s| s.to_str())
        .unwrap_or("")
        .to_lowercase();
    let spec = spec::extract_spec(&data, &file_name, &format_ext, sample_rows, &sheets)?;
    Ok(spec)
}

/// Python module entry point. Registers the public API surface Wave 0
/// commits `xlsb_reader/__init__.py` to depending on:
/// `XlsbWorkbook`, `XlsxWorkbook`, `col_to_letter`, and the one-shot
/// `to_dict`/`to_json`/`to_markdown`/`extract_spec` functions.
#[pymodule]
fn xlsb_reader_rs(m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add_class::<xlsb::XlsbWorkbook>()?;
    m.add_class::<xlsx::XlsxWorkbook>()?;
    m.add_function(wrap_pyfunction!(col_to_letter, m)?)?;
    m.add_function(wrap_pyfunction!(to_dict, m)?)?;
    m.add_function(wrap_pyfunction!(to_json, m)?)?;
    m.add_function(wrap_pyfunction!(to_markdown, m)?)?;
    m.add_function(wrap_pyfunction!(extract_spec, m)?)?;
    Ok(())
}
