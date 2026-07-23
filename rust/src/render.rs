//! render.rs — assembling parsed workbook data into `dict`/JSON/Markdown,
//! and the CLI's rendering logic.
//!
//! TODO(Wave 1D): Port `xlsb_reader/_render.py` (and the argparse-free parts
//! of `xlsb_reader/_cli.py`) in full:
//!   - `_cellmap`/`_cellmap_any`: `{(row,col): v}` -> `{"A1": v, ...}` via
//!     `common::col_to_letter`, sorted by `(row, col)`.
//!   - `_collect_formulas`/`_collect_values`/`_collect_pivots`/
//!     `_collect_filters`: per-section assembly, honoring an optional
//!     `sheet` filter.
//!   - `to_dict`/`to_json`/`to_markdown`'s section-inclusion logic (the
//!     `include` iterable of `formulas`/`values`/`pivots`/`filters`/`vba`).
//!   - `_as_markdown` byte-for-byte, including the "(no formulas found)"
//!     fallback and the VBA-modules appendix behaviour.
//!
//! These one-shot functions are what back the crate's `to_dict`/`to_json`/
//! `to_markdown` `#[pyfunction]`s in `lib.rs` (the "fast path" used by
//! `xlsb_reader.to_dict`/`to_json`/`to_markdown` in `__init__.py` when this
//! extension is installed) — output MUST match the pure-Python
//! `xlsb_reader._render` module byte-for-byte for the same `include`/
//! `sheet` arguments, since either backend can be selected transparently.

use std::collections::BTreeMap;

use pyo3::prelude::*;
use serde_json::{Map, Value};

use crate::common::{self, WorkbookData};

/// Section names accepted by `include=`, mirroring `_render.py`'s
/// `_ALL_SECTIONS`.
pub const ALL_SECTIONS: [&str; 5] = ["formulas", "values", "pivots", "filters", "vba"];

/// Default `include=` value, mirroring `_render.py`'s `_DEFAULT_INCLUDE`
/// and the CLI's `--include formulas,values,pivots` default.
pub const DEFAULT_INCLUDE: [&str; 3] = ["formulas", "values", "pivots"];

fn normalize_include(include: &[String]) -> std::collections::BTreeSet<String> {
    include
        .iter()
        .map(|s| s.trim().to_lowercase())
        .filter(|s| !s.is_empty())
        .collect()
}

/// Assemble the same `dict` shape `xlsb_reader._render.to_dict` produces,
/// as a `serde_json::Value` object — used both to build the Python `dict`
/// ([`to_dict_py`]) and to serialize directly to a JSON string (the crate's
/// `to_json` `#[pyfunction]`, via `serde_json::to_string`).
///
/// TODO(Wave 1D): implement for real. Current stub always returns an empty
/// JSON object regardless of `include`/`sheet`.
pub fn to_dict_value(_data: &WorkbookData, _include: &[String], _sheet: Option<&str>) -> Value {
    // STUB: Wave 1D replaces this with the real renderer. Once implemented,
    // this should build a `Map<String, Value>` with keys among
    // "formulas" / "values" / "pivot_tables" / "filters" / "vba_modules"
    // exactly as `_render._gather` does, respecting `normalize_include`.
    let _ = normalize_include(_include);
    Value::Object(Map::new())
}

/// Convert a [`to_dict_value`] result into a Python `dict`, used by the
/// crate's `to_dict` `#[pyfunction]` in `lib.rs`.
pub fn to_dict_py(
    py: Python<'_>,
    data: &WorkbookData,
    include: &[String],
    sheet: Option<&str>,
) -> PyResult<Py<PyAny>> {
    let value = to_dict_value(data, include, sheet);
    common::json_value_to_py(py, &value)
}

/// Render the same Markdown report `xlsb_reader._render.to_markdown`
/// produces.
///
/// TODO(Wave 1D): implement for real, matching `_render._as_markdown` (plus
/// the VBA-appendix behaviour in `to_markdown`) byte-for-byte. Current stub
/// always returns an empty string.
pub fn to_markdown_string(
    _data: &WorkbookData,
    _include: &[String],
    _sheet: Option<&str>,
) -> String {
    String::new()
}

/// Helper for Wave 1D: build the `{"A1": value, ...}` cell map a section
/// like `formulas`/`values` embeds per sheet, sorted by `(row, col)` like
/// `_render._cellmap`/`_cellmap_any`.
///
/// Provided now (rather than left for Wave 1D to write from scratch) since
/// it only depends on [`common::col_to_letter`], which is already final.
pub fn cell_map_to_json<T: Clone + Into<Value>>(map: &BTreeMap<(u32, u32), T>) -> Value {
    let mut obj = Map::new();
    for (&(row, col), value) in map.iter() {
        let key = format!("{}{}", common::col_to_letter(col), row + 1);
        obj.insert(key, value.clone().into());
    }
    Value::Object(obj)
}
