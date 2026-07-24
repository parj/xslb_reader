//! render.rs — assembling parsed workbook data into `dict`/JSON/Markdown,
//! and the CLI's rendering logic.
//!
//! Port of `xlsb_reader/_render.py`. These one-shot functions back the
//! crate's `to_dict`/`to_json`/`to_markdown` `#[pyfunction]`s in `lib.rs`
//! (the "fast path" used by `xlsb_reader.to_dict`/`to_json`/`to_markdown`
//! in `__init__.py` when this extension is installed) — output MUST match
//! the pure-Python `xlsb_reader._render` module byte-for-byte for the same
//! `include`/`sheet` arguments, since either backend can be selected
//! transparently.

use std::collections::BTreeSet;

use pyo3::prelude::*;
use serde_json::{Map, Value};

use crate::common::{self, WorkbookData};

/// Section names accepted by `include=`, mirroring `_render.py`'s
/// `_ALL_SECTIONS`.
#[allow(dead_code)]
pub const ALL_SECTIONS: [&str; 5] = ["formulas", "values", "pivots", "filters", "vba"];

/// Default `include=` value, mirroring `_render.py`'s `_DEFAULT_INCLUDE`
/// and the CLI's `--include formulas,values,pivots` default.
pub const DEFAULT_INCLUDE: [&str; 3] = ["formulas", "values", "pivots"];

fn normalize_include(include: &[String]) -> BTreeSet<String> {
    include
        .iter()
        .map(|s| s.trim().to_lowercase())
        .filter(|s| !s.is_empty())
        .collect()
}

/// Port of `_render._cellmap`/`_cellmap_any`: `{(row,col): v}` ->
/// `{"A1": v, ...}`, sorted by `(row, col)` (guaranteed here since the
/// source is already a `BTreeMap`).
fn cellmap_to_json<T: Clone + Into<Value>>(map: &common::CellMap<T>) -> Value {
    let mut obj = Map::new();
    for (&(row, col), value) in map.iter() {
        let key = format!("{}{}", common::col_to_letter(col), row + 1);
        obj.insert(key, value.clone().into());
    }
    Value::Object(obj)
}

impl From<common::CellValue> for Value {
    fn from(v: common::CellValue) -> Value {
        match v {
            common::CellValue::Int(i) => Value::from(i),
            common::CellValue::Float(f) => {
                serde_json::Number::from_f64(f).map(Value::Number).unwrap_or(Value::Null)
            }
            common::CellValue::Str(s) => Value::String(s),
            common::CellValue::Bool(b) => Value::Bool(b),
        }
    }
}

/// Port of `_render._collect_formulas`: `{sheet: {"A1": formula, ...}, ...}`,
/// only for sheets with at least one formula, optionally filtered to one
/// sheet.
fn collect_formulas(data: &WorkbookData, filter_sheet: Option<&str>) -> Map<String, Value> {
    let mut out = Map::new();
    for (sheet_name, formulas) in &data.formulas {
        if let Some(f) = filter_sheet {
            if sheet_name != f {
                continue;
            }
        }
        if !formulas.is_empty() {
            out.insert(sheet_name.clone(), cellmap_to_json(formulas));
        }
    }
    out
}

/// Port of `_render._collect_values`.
fn collect_values(data: &WorkbookData, filter_sheet: Option<&str>) -> Map<String, Value> {
    let mut out = Map::new();
    for (sheet_name, values) in &data.values {
        if let Some(f) = filter_sheet {
            if sheet_name != f {
                continue;
            }
        }
        if !values.is_empty() {
            out.insert(sheet_name.clone(), cellmap_to_json(values));
        }
    }
    out
}

/// Port of `_render._collect_pivots`: filters the flat pivot list by the
/// pivot's own `"sheet"` key (not by any sheet association on our side).
fn collect_pivots(data: &WorkbookData, filter_sheet: Option<&str>) -> Vec<Value> {
    data.pivots
        .iter()
        .filter(|pt| match filter_sheet {
            Some(f) => pt.get("sheet").and_then(Value::as_str) == Some(f),
            None => true,
        })
        .cloned()
        .collect()
}

/// Port of `_render._collect_filters`. `data.filters` always carries
/// `(sheet_name, info_or_None)` pairs regardless of backend (see
/// `WorkbookData::filters`'s doc comment) — unlike the pure-Python
/// `XlsbWorkbook.iter_filters()`, whose tuple-shaped items lack an
/// embedded `"sheet"` key on the info dict itself (only `XlsxWorkbook`'s do).
/// Always insert/overwrite `"sheet"` from our own tuple so the output is
/// uniformly shaped like XLSX's, matching `_render._collect_filters`'s
/// normalized Python-side behavior.
fn collect_filters(data: &WorkbookData, filter_sheet: Option<&str>) -> Vec<Value> {
    data.filters
        .iter()
        .filter_map(|(sheet_name, info)| info.as_ref().map(|v| (sheet_name, v)))
        .filter(|(sheet_name, _)| match filter_sheet {
            Some(f) => sheet_name.as_str() == f,
            None => true,
        })
        .map(|(sheet_name, v)| {
            let mut obj = match v {
                Value::Object(m) => m.clone(),
                _ => Map::new(),
            };
            obj.insert("sheet".to_string(), Value::String(sheet_name.clone()));
            Value::Object(obj)
        })
        .collect()
}

/// Order-preserving counterpart to [`collect_formulas`]/[`collect_values`]
/// for Markdown rendering — see [`as_markdown`]'s doc comment for why this
/// can't reuse the JSON-oriented `Map` those build.
fn collect_formulas_ordered<'a>(
    data: &'a WorkbookData,
    filter_sheet: Option<&str>,
) -> Vec<(&'a str, &'a common::CellMap<String>)> {
    data.formulas
        .iter()
        .filter(|(sheet_name, _)| match filter_sheet {
            Some(f) => sheet_name == f,
            None => true,
        })
        .filter(|(_, formulas)| !formulas.is_empty())
        .map(|(sheet_name, formulas)| (sheet_name.as_str(), formulas))
        .collect()
}

/// Order-preserving counterpart to [`collect_values`] — see
/// [`collect_formulas_ordered`].
fn collect_values_ordered<'a>(
    data: &'a WorkbookData,
    filter_sheet: Option<&str>,
) -> Vec<(&'a str, &'a common::CellMap<common::CellValue>)> {
    data.values
        .iter()
        .filter(|(sheet_name, _)| match filter_sheet {
            Some(f) => sheet_name == f,
            None => true,
        })
        .filter(|(_, values)| !values.is_empty())
        .map(|(sheet_name, values)| (sheet_name.as_str(), values))
        .collect()
}

struct Gathered {
    data: Map<String, Value>,
    formulas: Map<String, Value>,
    values: Map<String, Value>,
    pivots: Vec<Value>,
    filters: Vec<Value>,
    vba: Map<String, Value>,
}

/// Port of `_render._gather`.
///
/// `supports_vba` mirrors Python's `hasattr(wb, "iter_vba_modules")` guard:
/// `XlsbWorkbook` has no such method (so `"vba_modules"` is never added to
/// `to_dict()`'s output, even with `include=["vba"]`), while `XlsxWorkbook`
/// does. The one-shot functions in `lib.rs` determine this from the file
/// extension the same way `_spec_extractor.open_workbook` picks a reader.
fn gather(data: &WorkbookData, include: &[String], sheet: Option<&str>, supports_vba: bool) -> Gathered {
    let includes = normalize_include(include);

    let mut out = Gathered {
        data: Map::new(),
        formulas: Map::new(),
        values: Map::new(),
        pivots: Vec::new(),
        filters: Vec::new(),
        vba: Map::new(),
    };

    if includes.contains("formulas") {
        out.formulas = collect_formulas(data, sheet);
        out.data
            .insert("formulas".to_string(), Value::Object(out.formulas.clone()));
    }
    if includes.contains("values") {
        out.values = collect_values(data, sheet);
        out.data
            .insert("values".to_string(), Value::Object(out.values.clone()));
    }
    if includes.contains("pivots") {
        out.pivots = collect_pivots(data, sheet);
        out.data
            .insert("pivot_tables".to_string(), Value::Array(out.pivots.clone()));
    }
    if includes.contains("filters") {
        out.filters = collect_filters(data, sheet);
        out.data
            .insert("filters".to_string(), Value::Array(out.filters.clone()));
    }
    if includes.contains("vba") && supports_vba {
        out.vba = data
            .vba
            .iter()
            .map(|(k, v)| (k.clone(), Value::String(v.clone())))
            .collect();
        out.data
            .insert("vba_modules".to_string(), Value::Object(out.vba.clone()));
    }

    out
}

/// Assemble the same `dict` shape `xlsb_reader._render.to_dict` produces,
/// as a `serde_json::Value` object — used both to build the Python `dict`
/// ([`to_dict_py`]) and to serialize directly to a JSON string (the crate's
/// `to_json` `#[pyfunction]`, via `serde_json::to_string`).
pub fn to_dict_value(
    data: &WorkbookData,
    include: &[String],
    sheet: Option<&str>,
    supports_vba: bool,
) -> Value {
    Value::Object(gather(data, include, sheet, supports_vba).data)
}

/// Convert a [`to_dict_value`] result into a Python `dict`, used by the
/// crate's `to_dict` `#[pyfunction]` in `lib.rs`.
pub fn to_dict_py(
    py: Python<'_>,
    data: &WorkbookData,
    include: &[String],
    sheet: Option<&str>,
    supports_vba: bool,
) -> PyResult<Py<PyAny>> {
    let value = to_dict_value(data, include, sheet, supports_vba);
    common::json_value_to_py(py, &value)
}

fn json_display(v: &Value) -> String {
    match v {
        Value::Null => "None".to_string(),
        Value::Bool(b) => {
            if *b {
                "True".to_string()
            } else {
                "False".to_string()
            }
        }
        Value::String(s) => s.clone(),
        Value::Number(n) => n.to_string(),
        // Markdown rendering never formats a full array/object inline in
        // this module (every call site below extracts scalar fields first);
        // this arm only exists for exhaustiveness.
        other => other.to_string(),
    }
}

/// `str(value)` for a raw cell value, matching Python's f-string formatting
/// of formula/value cells in `_as_markdown` exactly — in particular using
/// `xlsb::python_float_repr` for floats (CPython's `repr(float)` algorithm)
/// rather than any JSON-oriented float formatting, since `str()`/f-strings
/// and `json.dumps` do not always agree on float text (see
/// [`to_dict_value`]'s doc comment for the JSON-side caveat).
fn cellvalue_str(v: &common::CellValue) -> String {
    match v {
        common::CellValue::Int(i) => i.to_string(),
        common::CellValue::Float(f) => crate::xlsb::python_float_repr(*f),
        common::CellValue::Str(s) => s.clone(),
        common::CellValue::Bool(b) => {
            if *b {
                "True".to_string()
            } else {
                "False".to_string()
            }
        }
    }
}

/// Port of `_render._as_markdown`.
///
/// `formulas`/`values` are `(sheet_name, cell_map)` pairs in **workbook tab
/// order**, each `cell_map` in **(row, col) order** — i.e. the raw
/// `WorkbookData` ordering, not the alphabetically-sorted-by-label ordering
/// `to_dict_value`'s JSON `Map` imposes. Python's `_collect_formulas`/
/// `_collect_values` build plain `dict`s that preserve exactly this
/// insertion order, and `_as_markdown` renders them by iterating that dict
/// directly — reusing the JSON `Map` here (sorted by string key, so e.g.
/// `"A10"` would sort before `"A2"`) would silently reorder the output.
fn as_markdown(
    sheets: &[String],
    formulas: Option<&[(&str, &common::CellMap<String>)]>,
    values: Option<&[(&str, &common::CellMap<common::CellValue>)]>,
    pivots: Option<&[Value]>,
    filters: Option<&[Value]>,
) -> String {
    let mut lines: Vec<String> = vec![format!("Sheets: {}", sheets.join(", ")), String::new()];
    let mut emitted = false;

    if let Some(formulas) = formulas {
        if !formulas.is_empty() {
            emitted = true;
            lines.push("## Formulas".to_string());
            lines.push(String::new());
            for (sheet, cell_formulas) in formulas {
                lines.push(format!("### {sheet}"));
                for (&(row, col), formula) in cell_formulas.iter() {
                    let cell = format!("{}{}", common::col_to_letter(col), row + 1);
                    lines.push(format!("- `{cell}`: `{formula}`"));
                }
                lines.push(String::new());
            }
        }
    }
    if let Some(values) = values {
        if !values.is_empty() {
            emitted = true;
            lines.push("## Values".to_string());
            lines.push(String::new());
            for (sheet, cell_values) in values {
                lines.push(format!("### {sheet}"));
                for (&(row, col), value) in cell_values.iter() {
                    let cell = format!("{}{}", common::col_to_letter(col), row + 1);
                    lines.push(format!("- `{cell}`: `{}`", cellvalue_str(value)));
                }
                lines.push(String::new());
            }
        }
    }
    if let Some(pivots) = pivots {
        if !pivots.is_empty() {
            emitted = true;
            lines.push("## Pivot Tables".to_string());
            lines.push(String::new());
            for pt in pivots {
                let name = pt
                    .get("name")
                    .and_then(Value::as_str)
                    .filter(|s| !s.is_empty())
                    .unwrap_or("<unnamed>");
                let sheet = pt.get("sheet").map(json_display).unwrap_or_default();
                let cache_id = pt.get("cache_id").map(json_display).unwrap_or_default();
                lines.push(format!(
                    "- `{name}` (sheet: `{sheet}`, cache_id: `{cache_id}`)"
                ));
                if let Some(location) = pt.get("location").filter(|l| l.is_object()) {
                    if let Some(rfx) = location.get("rfx_geom").filter(|r| r.is_object()) {
                        let top_left = rfx.get("top_left").and_then(Value::as_str);
                        let bottom_right = rfx.get("bottom_right").and_then(Value::as_str);
                        if let (Some(tl), Some(br)) = (top_left, bottom_right) {
                            if !tl.is_empty() && !br.is_empty() {
                                let pivot_fields =
                                    pt.get("pivot_fields").map(json_display).unwrap_or_default();
                                let pivot_items =
                                    pt.get("pivot_items").map(json_display).unwrap_or_default();
                                lines.push(format!(
                                    "  body: `{tl}:{br}`; fields: `{pivot_fields}`; items: `{pivot_items}`"
                                ));
                            }
                        }
                    }
                }
            }
        }
    }
    if let Some(filters) = filters {
        if !filters.is_empty() {
            emitted = true;
            lines.push("## Filters".to_string());
            lines.push(String::new());
            let mut by_sheet: Vec<(String, Vec<&Value>)> = Vec::new();
            for f in filters {
                let sheet = f
                    .get("sheet")
                    .and_then(Value::as_str)
                    .filter(|s| !s.is_empty())
                    .unwrap_or("<unknown>")
                    .to_string();
                match by_sheet.iter_mut().find(|(s, _)| *s == sheet) {
                    Some((_, v)) => v.push(f),
                    None => by_sheet.push((sheet, vec![f])),
                }
            }
            for (sheet, sheet_filters) in &by_sheet {
                lines.push(format!("### {sheet}"));
                for f in sheet_filters {
                    let reference = f.get("ref").and_then(Value::as_str).unwrap_or("");
                    let columns = f.get("columns").and_then(Value::as_array);
                    match columns {
                        Some(cols) if !cols.is_empty() => {
                            for col in cols {
                                let col_id = col
                                    .get("col_id")
                                    .map(json_display)
                                    .unwrap_or_else(|| "?".to_string());
                                let col_type =
                                    col.get("type").and_then(Value::as_str).unwrap_or("");
                                let conditions = col.get("conditions").and_then(Value::as_array);
                                let cond_str = match conditions {
                                    Some(conds) if !conds.is_empty() => conds
                                        .iter()
                                        .map(|c| {
                                            let op = c
                                                .get("operator")
                                                .and_then(Value::as_str)
                                                .unwrap_or("");
                                            let val = c
                                                .get("val")
                                                .map(json_display)
                                                .unwrap_or_default();
                                            format!("{op} {val}")
                                        })
                                        .collect::<Vec<_>>()
                                        .join("; "),
                                    // Python: `str(col.get("attrs", ""))` —
                                    // the default is an empty string, not
                                    // `None`.
                                    _ => col
                                        .get("attrs")
                                        .map(python_dict_repr)
                                        .unwrap_or_default(),
                                };
                                lines.push(format!(
                                    "- `{reference}` — column {col_id}: {col_type} {cond_str}"
                                ));
                            }
                        }
                        _ => lines.push(format!("- `{reference}`")),
                    }
                }
                lines.push(String::new());
            }
        }
    }

    if !emitted {
        lines.push("(no formulas found)".to_string());
        return lines.join("\n");
    }
    format!("{}\n", lines.join("\n").trim_end())
}

/// Render a JSON object/value the way Python's `str(dict)` would for the
/// `col.get("attrs")` fallback in [`as_markdown`] (a plain attrs dict with
/// no custom filter/discrete-values shape) — Python prints dict literals
/// with single-quoted string keys/values, e.g. `{'val': '5', 'type': '0'}`.
fn python_dict_repr(v: &Value) -> String {
    match v {
        Value::Object(map) => {
            let entries: Vec<String> = map
                .iter()
                .map(|(k, v)| format!("'{k}': {}", python_repr(v)))
                .collect();
            format!("{{{}}}", entries.join(", "))
        }
        other => python_repr(other),
    }
}

fn python_repr(v: &Value) -> String {
    match v {
        Value::String(s) => format!("'{s}'"),
        Value::Null => "None".to_string(),
        Value::Bool(b) => {
            if *b {
                "True".to_string()
            } else {
                "False".to_string()
            }
        }
        Value::Number(n) => n.to_string(),
        Value::Array(items) => {
            format!(
                "[{}]",
                items.iter().map(python_repr).collect::<Vec<_>>().join(", ")
            )
        }
        Value::Object(_) => python_dict_repr(v),
    }
}

/// Render the same Markdown report `xlsb_reader._render.to_markdown`
/// produces. `supports_vba` — see [`gather`]'s doc comment.
pub fn to_markdown_string(
    data: &WorkbookData,
    include: &[String],
    sheet: Option<&str>,
    supports_vba: bool,
) -> String {
    let includes = normalize_include(include);
    let gathered = gather(data, include, sheet, supports_vba);

    let formulas = includes
        .contains("formulas")
        .then(|| collect_formulas_ordered(data, sheet));
    let values = includes
        .contains("values")
        .then(|| collect_values_ordered(data, sheet));

    let mut md = as_markdown(
        &data.sheet_names,
        formulas.as_deref(),
        values.as_deref(),
        includes.contains("pivots").then_some(gathered.pivots.as_slice()),
        includes.contains("filters").then_some(gathered.filters.as_slice()),
    );

    if includes.contains("vba") && supports_vba && !data.vba.is_empty() {
        let mut vba_lines = vec!["## VBA Modules".to_string(), String::new()];
        for (mod_name, src) in &data.vba {
            vba_lines.push(format!("### {mod_name}"));
            vba_lines.push("```vba".to_string());
            vba_lines.push(src.trim_end().to_string());
            vba_lines.push("```".to_string());
            vba_lines.push(String::new());
        }
        md = format!("{}\n\n{}", md.trim_end(), vba_lines.join("\n"));
    }
    md
}
