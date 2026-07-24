//! xlsx.rs — `.xlsx`/`.xlsm` (Excel Open XML) support.
//!
//! Port of `xlsb_reader/_xlsx_reader.py`. Parses the OOXML container (a
//! plain ZIP of XML parts, same part paths as Python) with the `zip` crate
//! and `quick-xml`.
//!
//! `quick-xml` is a streaming/event parser, whereas the Python source uses
//! `xml.etree.ElementTree`'s DOM-style `.find()`/iteration API. To keep this
//! port a faithful, line-for-line mirror of the Python logic (and avoid
//! hand-rolling a bespoke event state machine per XML shape), [`parse_xml_tree`]
//! builds a small generic DOM (namespace-prefix-stripped, like Python's
//! `_ns()` helper) once per part, then every parser below walks it with
//! `.find()`/`.find_all()` helpers analogous to `ElementTree.find()`.

use std::collections::BTreeMap;
use std::io::Read;
use std::path::{Path, PathBuf};
use std::sync::{Mutex, OnceLock};

use pyo3::prelude::*;
use pyo3::types::{PyDict, PyList, PyTuple};
use quick_xml::events::Event;
use quick_xml::reader::Reader as XmlReader;
use serde_json::{Map, Value};

use crate::common::{self, CellValue, Result, WorkbookData, XlsbError};
use crate::vba;

// ---------------------------------------------------------------------------
// Cell reference helpers — port of `_col_from_str` / `_cell_ref_to_row_col`.
// ---------------------------------------------------------------------------

/// Convert Excel column letters to a 0-based index, e.g. `"A" -> 0`, `"AA" -> 26`.
fn col_from_str(s: &str) -> i64 {
    let mut result: i64 = 0;
    for ch in s.chars() {
        result = result * 26 + (ch as i64 - 'A' as i64 + 1);
    }
    result - 1
}

/// Parse a cell reference like `"K3"` into 0-based `(row, col)`.
fn cell_ref_to_row_col(reference: &str) -> (i64, i64) {
    let split = reference
        .find(|c: char| !c.is_ascii_alphabetic())
        .unwrap_or(reference.len());
    let (col_str, row_str) = reference.split_at(split);
    let col = col_from_str(col_str);
    let row = row_str.parse::<i64>().unwrap_or(1) - 1;
    (row, col)
}

// ---------------------------------------------------------------------------
// Minimal generic XML tree — see module doc comment.
// ---------------------------------------------------------------------------

struct XmlNode {
    tag: String,
    attrs: Vec<(String, String)>,
    text: String,
    children: Vec<XmlNode>,
}

impl XmlNode {
    fn attr(&self, name: &str) -> Option<&str> {
        self.attrs
            .iter()
            .find(|(k, _)| k == name)
            .map(|(_, v)| v.as_str())
    }

    fn attr_or<'a>(&'a self, name: &str, default: &'a str) -> &'a str {
        self.attr(name).unwrap_or(default)
    }

    fn find(&self, tag: &str) -> Option<&XmlNode> {
        self.children.iter().find(|c| c.tag == tag)
    }

    fn find_all<'a>(&'a self, tag: &'a str) -> impl Iterator<Item = &'a XmlNode> {
        self.children.iter().filter(move |c| c.tag == tag)
    }

    /// `el.text or ""` — Python's fallback for a `None` text node.
    fn text_or_empty(&self) -> &str {
        &self.text
    }
}

/// Strip an XML namespace prefix (`"r:id"` -> `"id"`), matching Python's
/// `_ns()` local-name stripping. Namespace *disambiguation* is not needed
/// here since each of these part types only ever uses one prefix per
/// meaning in practice (see module doc comment).
fn local_name(qname: &[u8]) -> String {
    let s = String::from_utf8_lossy(qname);
    match s.rfind(':') {
        Some(i) => s[i + 1..].to_string(),
        None => s.to_string(),
    }
}

/// Parse XML bytes into a small generic tree, or `None` on malformed XML —
/// mirrors `except ET.ParseError: return {}/[]` at every Python call site.
fn parse_xml_tree(data: &[u8]) -> Option<XmlNode> {
    let mut reader = XmlReader::from_reader(data);
    let mut stack: Vec<XmlNode> = Vec::new();
    let mut root: Option<XmlNode> = None;
    let mut buf = Vec::new();

    fn read_attrs(e: &quick_xml::events::BytesStart<'_>) -> Vec<(String, String)> {
        e.attributes()
            .flatten()
            .map(|a| {
                let key = local_name(a.key.as_ref());
                let val = String::from_utf8_lossy(&a.value).into_owned();
                let val = quick_xml::escape::unescape(&val)
                    .map(|c| c.into_owned())
                    .unwrap_or(val);
                (key, val)
            })
            .collect()
    }

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Eof) => break,
            Ok(Event::Start(e)) => {
                let tag = local_name(e.name().as_ref());
                let attrs = read_attrs(&e);
                stack.push(XmlNode {
                    tag,
                    attrs,
                    text: String::new(),
                    children: Vec::new(),
                });
            }
            Ok(Event::Empty(e)) => {
                let tag = local_name(e.name().as_ref());
                let attrs = read_attrs(&e);
                let node = XmlNode {
                    tag,
                    attrs,
                    text: String::new(),
                    children: Vec::new(),
                };
                match stack.last_mut() {
                    Some(parent) => parent.children.push(node),
                    None => root = Some(node),
                }
            }
            Ok(Event::Text(t)) => {
                if let Some(top) = stack.last_mut() {
                    if let Ok(decoded) = t.decode() {
                        match quick_xml::escape::unescape(&decoded) {
                            Ok(txt) => top.text.push_str(&txt),
                            Err(_) => top.text.push_str(&decoded),
                        }
                    }
                }
            }
            Ok(Event::CData(t)) => {
                if let Some(top) = stack.last_mut() {
                    top.text.push_str(&String::from_utf8_lossy(&t.into_inner()));
                }
            }
            // quick-xml emits `&entity;`/`&#N;` references embedded in text
            // as a separate `GeneralRef` event rather than folding them into
            // the surrounding `Text` event — without this arm they would
            // silently vanish (e.g. `A3&gt;0` -> `A30` instead of `A3>0`).
            Ok(Event::GeneralRef(r)) => {
                if let Some(top) = stack.last_mut() {
                    match r.resolve_char_ref() {
                        Ok(Some(ch)) => top.text.push(ch),
                        Ok(None) => {
                            if let Ok(name) = r.decode() {
                                if let Some(resolved) =
                                    quick_xml::escape::resolve_predefined_entity(&name)
                                {
                                    top.text.push_str(resolved);
                                }
                            }
                        }
                        Err(_) => {}
                    }
                }
            }
            Ok(Event::End(_)) => match stack.pop() {
                Some(node) => match stack.last_mut() {
                    Some(parent) => parent.children.push(node),
                    None => root = Some(node),
                },
                None => return None,
            },
            Ok(_) => {}
            Err(_) => return None,
        }
        buf.clear();
    }
    root
}

// ---------------------------------------------------------------------------
// Relationships — port of `_parse_rels` / `_resolve_rel_target`.
// ---------------------------------------------------------------------------

/// Return `[(Id, Target), ...]` from a `.rels` XML file, document order.
fn parse_rels(data: &[u8]) -> Vec<(String, String)> {
    let root = match parse_xml_tree(data) {
        Some(r) => r,
        None => return Vec::new(),
    };
    root.children
        .iter()
        .filter(|c| c.tag == "Relationship")
        .filter_map(|c| {
            let id = c.attr("Id").unwrap_or("");
            let target = c.attr("Target").unwrap_or("");
            if id.is_empty() || target.is_empty() {
                None
            } else {
                Some((id.to_string(), target.to_string()))
            }
        })
        .collect()
}

fn rels_get<'a>(rels: &'a [(String, String)], id: &str) -> Option<&'a str> {
    rels.iter().find(|(k, _)| k == id).map(|(_, v)| v.as_str())
}

/// Resolve a relationship target relative to `part_path`'s directory.
fn resolve_rel_target(part_path: &str, target: &str) -> String {
    let base = match part_path.rfind('/') {
        Some(i) => &part_path[..i],
        None => "",
    };
    if let Some(stripped) = target.strip_prefix('/') {
        return stripped.trim_start_matches('/').to_string();
    }
    let joined = format!("{base}/{target}");
    let mut resolved: Vec<&str> = Vec::new();
    for part in joined.split('/') {
        if part == ".." {
            resolved.pop();
        } else if !part.is_empty() && part != "." {
            resolved.push(part);
        }
    }
    resolved.join("/")
}

fn dirname(p: &str) -> String {
    match p.rfind('/') {
        Some(i) => p[..i].to_string(),
        None => String::new(),
    }
}

fn basename(p: &str) -> String {
    match p.rfind('/') {
        Some(i) => p[i + 1..].to_string(),
        None => p.to_string(),
    }
}

// ---------------------------------------------------------------------------
// Shared string table — port of `_parse_shared_strings`.
// ---------------------------------------------------------------------------

fn parse_shared_strings(data: &[u8]) -> Vec<String> {
    let root = match parse_xml_tree(data) {
        Some(r) => r,
        None => return Vec::new(),
    };
    root.find_all("si")
        .map(|si| {
            let runs: Vec<&XmlNode> = si.find_all("r").collect();
            if !runs.is_empty() {
                runs.iter()
                    .flat_map(|r| r.find_all("t"))
                    .map(|t| t.text_or_empty())
                    .collect::<String>()
            } else {
                si.find("t").map(|t| t.text.clone()).unwrap_or_default()
            }
        })
        .collect()
}

// ---------------------------------------------------------------------------
// Workbook XML — port of `_parse_workbook`.
// ---------------------------------------------------------------------------

fn parse_workbook_xml(data: &[u8]) -> (Vec<(String, String)>, BTreeMap<String, String>) {
    let root = match parse_xml_tree(data) {
        Some(r) => r,
        None => return (Vec::new(), BTreeMap::new()),
    };
    let mut sheets = Vec::new();
    let mut defined_names = BTreeMap::new();

    if let Some(sheets_el) = root.find("sheets") {
        for sheet_el in sheets_el.find_all("sheet") {
            let name = sheet_el.attr_or("name", "");
            let rid = sheet_el.attr_or("id", "");
            if !name.is_empty() && !rid.is_empty() {
                sheets.push((name.to_string(), rid.to_string()));
            }
        }
    }
    if let Some(dn_el) = root.find("definedNames") {
        for dn in dn_el.find_all("definedName") {
            let name = dn.attr_or("name", "");
            if !name.is_empty() {
                defined_names.insert(name.to_string(), dn.text.clone());
            }
        }
    }
    (sheets, defined_names)
}

// ---------------------------------------------------------------------------
// Shared-formula cell-reference shifting — port of `_shift_ref` /
// `_expand_shared_formula`.
// ---------------------------------------------------------------------------

fn cell_ref_regex() -> &'static regex::Regex {
    static RE: OnceLock<regex::Regex> = OnceLock::new();
    RE.get_or_init(|| regex::Regex::new(r"(\$?)([A-Z]+)(\$?)(\d+)").unwrap())
}

fn expand_shared_formula(formula: &str, dr: i64, dc: i64) -> String {
    if dr == 0 && dc == 0 {
        return formula.to_string();
    }
    cell_ref_regex()
        .replace_all(formula, |caps: &regex::Captures| {
            let col_abs = &caps[1];
            let col_str = &caps[2];
            let row_abs = &caps[3];
            let row_str = &caps[4];
            let mut col = col_from_str(col_str);
            let mut row = row_str.parse::<i64>().unwrap_or(1) - 1;
            if col_abs.is_empty() {
                col += dc;
            }
            if row_abs.is_empty() {
                row += dr;
            }
            format!(
                "{col_abs}{}{row_abs}{}",
                common::col_to_letter(col.max(0) as u32),
                row + 1
            )
        })
        .into_owned()
}

// ---------------------------------------------------------------------------
// Worksheet XML — formulas. Port of `_parse_worksheet_formulas`.
// ---------------------------------------------------------------------------

fn parse_worksheet_formulas(data: &[u8]) -> BTreeMap<(u32, u32), String> {
    let root = match parse_xml_tree(data) {
        Some(r) => r,
        None => return BTreeMap::new(),
    };
    let mut formulas: BTreeMap<(u32, u32), String> = BTreeMap::new();
    let sheet_data = match root.find("sheetData") {
        Some(sd) => sd,
        None => return formulas,
    };

    // si -> (formula_text, anchor_row, anchor_col)
    let mut shared_formulas: BTreeMap<String, (String, i64, i64)> = BTreeMap::new();

    for row_el in sheet_data.find_all("row") {
        for c_el in row_el.find_all("c") {
            let reference = c_el.attr_or("r", "");
            if reference.is_empty() {
                continue;
            }
            let (row, col) = cell_ref_to_row_col(reference);

            let f_el = match c_el.find("f") {
                Some(f) => f,
                None => continue,
            };
            let f_text = f_el.text_or_empty();
            let f_type = f_el.attr_or("t", "normal");
            let f_si = f_el.attr_or("si", "");
            let f_ref = f_el.attr_or("ref", "");
            let f_ca = f_el.attr_or("ca", "0");

            let strip_udf = |s: &str| -> String {
                if f_ca == "1" {
                    s.replace("_xludf.", "")
                } else {
                    s.to_string()
                }
            };

            match f_type {
                "array" => {
                    let formula_text = strip_udf(f_text);
                    formulas.insert(
                        (row as u32, col as u32),
                        format!("{{={formula_text}}}"),
                    );
                }
                "shared" => {
                    if !f_ref.is_empty() {
                        let formula_text = strip_udf(f_text);
                        if !f_si.is_empty() {
                            shared_formulas
                                .insert(f_si.to_string(), (formula_text.clone(), row, col));
                        }
                        formulas.insert((row as u32, col as u32), format!("={formula_text}"));
                    } else if !f_si.is_empty() {
                        if let Some((anchor_formula, anchor_row, anchor_col)) =
                            shared_formulas.get(f_si)
                        {
                            let dr = row - anchor_row;
                            let dc = col - anchor_col;
                            let expanded = expand_shared_formula(anchor_formula, dr, dc);
                            formulas.insert((row as u32, col as u32), format!("={expanded}"));
                        }
                        // else: shared formula not seen yet — matches Python (missing entry).
                    }
                }
                _ => {
                    let formula_text = strip_udf(f_text);
                    formulas.insert((row as u32, col as u32), format!("={formula_text}"));
                }
            }
        }
    }
    formulas
}

// ---------------------------------------------------------------------------
// Worksheet XML — values. Port of `_parse_worksheet_values`.
// ---------------------------------------------------------------------------

fn numeric_cell_value(v_text: &str) -> CellValue {
    match v_text.parse::<f64>() {
        Ok(f) if f.fract() == 0.0 && f.abs() < 1e15 => CellValue::Int(f as i64),
        Ok(f) => CellValue::Float(f),
        Err(_) => CellValue::Str(v_text.to_string()),
    }
}

fn parse_worksheet_values(data: &[u8], sst: &[String]) -> BTreeMap<(u32, u32), CellValue> {
    let root = match parse_xml_tree(data) {
        Some(r) => r,
        None => return BTreeMap::new(),
    };
    let mut values: BTreeMap<(u32, u32), CellValue> = BTreeMap::new();
    let sheet_data = match root.find("sheetData") {
        Some(sd) => sd,
        None => return values,
    };

    for row_el in sheet_data.find_all("row") {
        for c_el in row_el.find_all("c") {
            let reference = c_el.attr_or("r", "");
            if reference.is_empty() {
                continue;
            }
            let (row, col) = cell_ref_to_row_col(reference);
            let key = (row as u32, col as u32);
            let cell_type = c_el.attr_or("t", "");

            if cell_type == "inlineStr" {
                let text = c_el
                    .find("is")
                    .and_then(|is_el| is_el.find("t"))
                    .map(|t| t.text.clone())
                    .unwrap_or_default();
                values.insert(key, CellValue::Str(text));
                continue;
            }

            let v_text = match c_el.find("v") {
                Some(v_el) => v_el.text.clone(),
                None => continue, // no <v> — blank cell, matches Python's `continue`.
            };

            let value = match cell_type {
                "s" => {
                    let idx: Option<usize> = v_text.parse().ok();
                    match idx.and_then(|i| sst.get(i)) {
                        Some(s) => CellValue::Str(s.clone()),
                        None => CellValue::Str(String::new()),
                    }
                }
                // Python does `bool(int(v_text))` unguarded here; real OOXML
                // files always encode boolean cells as "0"/"1", so a parse
                // failure (malformed input) falls back to `false` rather
                // than propagating a panic/exception.
                "b" => CellValue::Bool(v_text.trim().parse::<i64>().unwrap_or(0) != 0),
                "e" => CellValue::Str(v_text),
                "str" => CellValue::Str(v_text),
                _ => numeric_cell_value(&v_text),
            };
            values.insert(key, value);
        }
    }
    values
}

// ---------------------------------------------------------------------------
// Pivot table XML — port of `_parse_pivot_table_xml`.
// ---------------------------------------------------------------------------

fn parse_pivot_table_xml(data: &[u8]) -> Map<String, Value> {
    let mut meta = Map::new();
    let root = match parse_xml_tree(data) {
        Some(r) => r,
        None => return meta,
    };

    meta.insert("name".to_string(), Value::String(root.attr_or("name", "").to_string()));
    let cache_id: i64 = root.attr_or("cacheId", "0").parse().unwrap_or(0);
    meta.insert("cache_id".to_string(), Value::from(cache_id));
    meta.insert(
        "data_caption".to_string(),
        Value::String(root.attr_or("dataCaption", "").to_string()),
    );

    if let Some(loc_el) = root.find("location") {
        let reference = loc_el.attr_or("ref", "");
        let (top_left, bottom_right) = match reference.split_once(':') {
            Some((a, b)) => (a, b),
            None => (reference, reference),
        };
        let rw_first_head: i64 = loc_el.attr_or("firstHeaderRow", "0").parse().unwrap_or(0);
        let rw_first_data: i64 = loc_el.attr_or("firstDataRow", "0").parse().unwrap_or(0);
        let col_first_data_int: i64 = loc_el.attr_or("firstDataCol", "0").parse().unwrap_or(0);
        let col_first_data_letter = if !top_left.is_empty() {
            let (_tl_row, tl_col) = cell_ref_to_row_col(top_left);
            common::col_to_letter((tl_col + col_first_data_int).max(0) as u32)
        } else {
            common::col_to_letter(col_first_data_int.max(0) as u32)
        };

        let mut rfx_geom = Map::new();
        rfx_geom.insert("top_left".to_string(), Value::String(top_left.to_string()));
        rfx_geom.insert(
            "bottom_right".to_string(),
            Value::String(bottom_right.to_string()),
        );
        let mut location = Map::new();
        location.insert("rfx_geom".to_string(), Value::Object(rfx_geom));
        location.insert("rw_first_head".to_string(), Value::from(rw_first_head));
        location.insert("rw_first_data".to_string(), Value::from(rw_first_data));
        location.insert("col_first_data".to_string(), Value::String(col_first_data_letter));
        location.insert("page_rows".to_string(), Value::from(0));
        location.insert("page_cols".to_string(), Value::from(0));
        meta.insert("location".to_string(), Value::Object(location));
    } else {
        meta.insert("location".to_string(), Value::Null);
    }

    let mut pivot_fields_count: i64 = 0;
    let mut pivot_items_count: i64 = 0;
    let mut field_filters: Vec<Value> = Vec::new();
    if let Some(pf_el) = root.find("pivotFields") {
        for pf in pf_el.find_all("pivotField") {
            let field_index = pivot_fields_count;
            pivot_fields_count += 1;
            let mut hidden_indices: Vec<i64> = Vec::new();
            if let Some(items_el) = pf.find("items") {
                for item in items_el.find_all("item") {
                    pivot_items_count += 1;
                    if item.attr("h") == Some("1") {
                        if let Some(x) = item.attr("x").and_then(|v| v.parse::<i64>().ok()) {
                            hidden_indices.push(x);
                        }
                    }
                }
            }
            if !hidden_indices.is_empty() {
                let mut ff = Map::new();
                ff.insert("field_index".to_string(), Value::from(field_index));
                ff.insert(
                    "hidden_indices".to_string(),
                    Value::Array(hidden_indices.into_iter().map(Value::from).collect()),
                );
                field_filters.push(Value::Object(ff));
            }
        }
    }
    meta.insert("pivot_fields".to_string(), Value::from(pivot_fields_count));
    meta.insert("pivot_items".to_string(), Value::from(pivot_items_count));
    meta.insert("field_filters".to_string(), Value::Array(field_filters));

    let mut row_fields: Vec<Value> = Vec::new();
    if let Some(rf_el) = root.find("rowFields") {
        for field_el in rf_el.find_all("field") {
            if let Ok(x) = field_el.attr_or("x", "0").parse::<i64>() {
                row_fields.push(Value::from(x));
            }
        }
    }
    meta.insert("row_fields".to_string(), Value::Array(row_fields));

    let mut col_fields: Vec<Value> = Vec::new();
    if let Some(cf_el) = root.find("colFields") {
        for field_el in cf_el.find_all("field") {
            if let Ok(x) = field_el.attr_or("x", "0").parse::<i64>() {
                if x != -2 {
                    col_fields.push(Value::from(x));
                }
            }
        }
    }
    meta.insert("col_fields".to_string(), Value::Array(col_fields));

    let mut data_fields: Vec<Value> = Vec::new();
    if let Some(df_el) = root.find("dataFields") {
        for df in df_el.find_all("dataField") {
            let mut entry = Map::new();
            entry.insert(
                "name".to_string(),
                Value::String(df.attr_or("name", "").to_string()),
            );
            let fld: i64 = df.attr_or("fld", "0").parse().unwrap_or(0);
            entry.insert("fld".to_string(), Value::from(fld));
            entry.insert(
                "subtotal".to_string(),
                Value::String(df.attr_or("subtotal", "sum").to_string()),
            );
            data_fields.push(Value::Object(entry));
        }
    }
    meta.insert("data_fields".to_string(), Value::Array(data_fields));

    let mut report_filters: Vec<Value> = Vec::new();
    if let Some(filters_el) = root.find("filters") {
        for filt_el in filters_el.find_all("filter") {
            let fld_idx: i64 = filt_el.attr_or("fld", "0").parse().unwrap_or(0);
            let mut entry = Map::new();
            entry.insert("fld".to_string(), Value::from(fld_idx));
            entry.insert(
                "type".to_string(),
                Value::String(filt_el.attr_or("type", "").to_string()),
            );
            let custom_filter = filt_el
                .find("autoFilter")
                .and_then(|a| a.find("filterColumn"))
                .and_then(|a| a.find("customFilters"))
                .and_then(|a| a.find("customFilter"));
            if let Some(cf) = custom_filter {
                entry.insert(
                    "operator".to_string(),
                    match cf.attr("operator") {
                        Some(v) => Value::String(v.to_string()),
                        None => Value::Null,
                    },
                );
                entry.insert(
                    "value".to_string(),
                    match cf.attr("val") {
                        Some(v) => Value::String(v.to_string()),
                        None => Value::Null,
                    },
                );
            }
            report_filters.push(Value::Object(entry));
        }
    }
    meta.insert("report_filters".to_string(), Value::Array(report_filters));

    meta
}

// ---------------------------------------------------------------------------
// Pivot cache definition XML — port of `_parse_pivot_cache_fields_xml`.
// ---------------------------------------------------------------------------

fn parse_pivot_cache_fields_xml(data: &[u8]) -> Vec<Value> {
    let root = match parse_xml_tree(data) {
        Some(r) => r,
        None => return Vec::new(),
    };
    let cache_fields_el = match root.find("cacheFields") {
        Some(el) => el,
        None => return Vec::new(),
    };
    cache_fields_el
        .find_all("cacheField")
        .map(|cf| {
            let mut shared_items: Vec<Value> = Vec::new();
            if let Some(si_el) = cf.find("sharedItems") {
                for item in &si_el.children {
                    match item.tag.as_str() {
                        "s" => shared_items.push(Value::String(item.attr_or("v", "").to_string())),
                        "n" => {
                            let v: f64 = item.attr_or("v", "0").parse().unwrap_or(0.0);
                            shared_items.push(
                                serde_json::Number::from_f64(v)
                                    .map(Value::Number)
                                    .unwrap_or(Value::String(item.attr_or("v", "").to_string())),
                            );
                        }
                        "b" => shared_items.push(Value::Bool(item.attr("v") == Some("1"))),
                        _ => {}
                    }
                }
            }
            let mut entry = Map::new();
            entry.insert(
                "name".to_string(),
                Value::String(cf.attr_or("name", "").to_string()),
            );
            entry.insert("shared_items".to_string(), Value::Array(shared_items));
            Value::Object(entry)
        })
        .collect()
}

// ---------------------------------------------------------------------------
// Auto-filter XML — port of `_parse_auto_filter`.
// ---------------------------------------------------------------------------

fn parse_auto_filter(af_el: &XmlNode, sheet_name: &str) -> Value {
    let reference = af_el.attr_or("ref", "");
    let mut columns: Vec<Value> = Vec::new();

    for fc_el in af_el.find_all("filterColumn") {
        let col_id: i64 = fc_el.attr_or("colId", "0").parse().unwrap_or(0);
        let mut col_info = Map::new();
        col_info.insert("col_id".to_string(), Value::from(col_id));

        // Only the first filter-type child is processed, matching Python's
        // `break` after the first iteration of the inner loop.
        if let Some(filter_child) = fc_el.children.first() {
            match filter_child.tag.as_str() {
                "customFilters" => {
                    col_info.insert("type".to_string(), Value::String("custom".to_string()));
                    let conditions: Vec<Value> = filter_child
                        .find_all("customFilter")
                        .map(|cf| {
                            let mut c = Map::new();
                            c.insert(
                                "operator".to_string(),
                                Value::String(cf.attr_or("operator", "equal").to_string()),
                            );
                            c.insert(
                                "val".to_string(),
                                Value::String(cf.attr_or("val", "").to_string()),
                            );
                            Value::Object(c)
                        })
                        .collect();
                    col_info.insert("conditions".to_string(), Value::Array(conditions));
                }
                "filters" => {
                    col_info.insert("type".to_string(), Value::String("discrete".to_string()));
                    let vals: Vec<Value> = filter_child
                        .find_all("filter")
                        .map(|f| Value::String(f.attr_or("val", "").to_string()))
                        .collect();
                    col_info.insert("values".to_string(), Value::Array(vals));
                }
                "top10" => {
                    col_info.insert("type".to_string(), Value::String("top10".to_string()));
                    col_info.insert("attrs".to_string(), attrs_to_json(filter_child));
                }
                "dynamicFilter" => {
                    col_info.insert("type".to_string(), Value::String("dynamic".to_string()));
                    col_info.insert("attrs".to_string(), attrs_to_json(filter_child));
                }
                _ => {}
            }
        }
        columns.push(Value::Object(col_info));
    }

    let mut result = Map::new();
    result.insert("sheet".to_string(), Value::String(sheet_name.to_string()));
    result.insert("ref".to_string(), Value::String(reference.to_string()));
    result.insert("columns".to_string(), Value::Array(columns));
    Value::Object(result)
}

fn attrs_to_json(node: &XmlNode) -> Value {
    let mut m = Map::new();
    for (k, v) in &node.attrs {
        m.insert(k.clone(), Value::String(v.clone()));
    }
    Value::Object(m)
}

// ---------------------------------------------------------------------------
// Top-level parse: open the ZIP, orchestrate the reads above.
// ---------------------------------------------------------------------------

type ZipArchive = zip::ZipArchive<std::io::BufReader<std::fs::File>>;

/// Read a part by name, with the same case-insensitive fallback as
/// `XlsxWorkbook._read_part`.
fn read_part(zf: &mut ZipArchive, path: &str) -> Result<Vec<u8>> {
    let norm = path.replace('\\', "/");
    let norm = norm.trim_start_matches('/');
    if let Ok(mut f) = zf.by_name(norm) {
        let mut buf = Vec::with_capacity(f.size() as usize);
        f.read_to_end(&mut buf)?;
        return Ok(buf);
    }
    let lo = norm.to_lowercase();
    let found = zf
        .file_names()
        .find(|n| n.to_lowercase() == lo)
        .map(|s| s.to_string());
    match found {
        Some(n) => {
            let mut f = zf.by_name(&n)?;
            let mut buf = Vec::with_capacity(f.size() as usize);
            f.read_to_end(&mut buf)?;
            Ok(buf)
        }
        None => Err(XlsbError::Parse(format!("Not found in XLSX zip: {path}"))),
    }
}

/// Parse a `.xlsx`/`.xlsm` file at `path` into a [`WorkbookData`]. Port of
/// `XlsxWorkbook.__init__`/`_init_workbook`/`_init_sst` plus eager versions
/// of all four `iter_*` methods (the pyclass caches the whole
/// [`WorkbookData`] up front rather than re-parsing per call).
pub fn parse_xlsx(path: &Path) -> Result<WorkbookData> {
    let file = std::fs::File::open(path)?;
    let mut zf = zip::ZipArchive::new(std::io::BufReader::new(file))?;

    let wb_data = read_part(&mut zf, "xl/workbook.xml")?;
    let (sheets, _defined_names) = parse_workbook_xml(&wb_data);

    let rels = read_part(&mut zf, "xl/_rels/workbook.xml.rels")
        .map(|b| parse_rels(&b))
        .unwrap_or_default();

    let mut paths: BTreeMap<String, String> = BTreeMap::new();
    for (_name, rid) in &sheets {
        if let Some(target) = rels_get(&rels, rid) {
            if !target.is_empty() {
                let full = if let Some(stripped) = target.strip_prefix('/') {
                    stripped.to_string()
                } else if target.starts_with("xl/") {
                    target.to_string()
                } else {
                    format!("xl/{target}")
                };
                paths.insert(rid.clone(), full.replace("//", "/"));
            }
        }
    }

    let sst = read_part(&mut zf, "xl/sharedStrings.xml")
        .map(|b| parse_shared_strings(&b))
        .unwrap_or_default();

    let sheet_names: Vec<String> = sheets.iter().map(|(n, _)| n.clone()).collect();
    let mut formulas = Vec::new();
    let mut values = Vec::new();
    let mut filters = Vec::new();

    for (sheet_name, rid) in &sheets {
        let ws_data = match paths.get(rid).and_then(|p| read_part(&mut zf, p).ok()) {
            Some(d) => d,
            None => {
                formulas.push((sheet_name.clone(), BTreeMap::new()));
                values.push((sheet_name.clone(), BTreeMap::new()));
                filters.push((sheet_name.clone(), None));
                continue;
            }
        };
        formulas.push((sheet_name.clone(), parse_worksheet_formulas(&ws_data)));
        values.push((sheet_name.clone(), parse_worksheet_values(&ws_data, &sst)));

        let filter_info = parse_xml_tree(&ws_data)
            .and_then(|root| root.find("autoFilter").map(|af| parse_auto_filter(af, sheet_name)));
        filters.push((sheet_name.clone(), filter_info));
    }

    let pivots = parse_pivot_tables(&mut zf, &sheets, &paths);

    let vba = read_part(&mut zf, "xl/vbaProject.bin")
        .ok()
        .and_then(|bin| vba::read_vba_modules(&bin).ok())
        .unwrap_or_default();

    Ok(WorkbookData {
        sheet_names,
        formulas,
        values,
        filters,
        pivots,
        vba,
    })
}

/// Discover PivotTable parts through worksheet relationships and parse
/// them, linked to sheets where possible. Port of
/// `XlsxWorkbook.iter_pivot_tables`.
fn parse_pivot_tables(
    zf: &mut ZipArchive,
    sheets: &[(String, String)],
    paths: &BTreeMap<String, String>,
) -> Vec<Value> {
    let mut order: Vec<String> = Vec::new();
    let mut sheet_by_path: BTreeMap<String, String> = BTreeMap::new();
    for (name, rid) in sheets {
        if let Some(zp) = paths.get(rid) {
            if !sheet_by_path.contains_key(zp) {
                order.push(zp.clone());
            }
            sheet_by_path.insert(zp.clone(), name.clone());
        }
    }

    let mut seen_parts: std::collections::BTreeSet<String> = std::collections::BTreeSet::new();
    let mut cache_fields_by_part: BTreeMap<String, Vec<Value>> = BTreeMap::new();
    let mut pivots: Vec<Value> = Vec::new();

    for sheet_path in &order {
        let sheet_name = match sheet_by_path.get(sheet_path) {
            Some(n) => n.clone(),
            None => continue,
        };
        let rel_path = format!("{}/_rels/{}.rels", dirname(sheet_path), basename(sheet_path));
        let rels = match read_part(zf, &rel_path) {
            Ok(b) => parse_rels(&b),
            Err(_) => continue,
        };

        for (_id, target) in &rels {
            let resolved = resolve_rel_target(sheet_path, target);
            if !resolved.contains("/pivotTables/") {
                continue;
            }
            if seen_parts.contains(&resolved) {
                continue;
            }
            seen_parts.insert(resolved.clone());
            let pt_data = match read_part(zf, &resolved) {
                Ok(b) => b,
                Err(_) => continue,
            };
            let mut meta = parse_pivot_table_xml(&pt_data);
            meta.insert("sheet".to_string(), Value::String(sheet_name.clone()));
            meta.insert("part".to_string(), Value::String(resolved.clone()));

            let pt_rel_path = format!(
                "{}/_rels/{}.rels",
                dirname(&resolved),
                basename(&resolved)
            );
            let pt_rels = read_part(zf, &pt_rel_path)
                .map(|b| parse_rels(&b))
                .unwrap_or_default();

            let mut cache_def: Option<String> = None;
            for (_id, t) in &pt_rels {
                let rr = resolve_rel_target(&resolved, t);
                if rr.contains("/pivotCache/pivotCacheDefinition") {
                    cache_def = Some(rr);
                    break;
                }
            }
            if let Some(cd) = cache_def {
                meta.insert(
                    "pivot_cache_definition".to_string(),
                    Value::String(cd.clone()),
                );
                if !cache_fields_by_part.contains_key(&cd) {
                    let fields = match read_part(zf, &cd) {
                        Ok(cdata) => parse_pivot_cache_fields_xml(&cdata),
                        Err(_) => Vec::new(),
                    };
                    cache_fields_by_part.insert(cd.clone(), fields);
                }
                meta.insert(
                    "cache_fields".to_string(),
                    Value::Array(cache_fields_by_part[&cd].clone()),
                );
            }

            pivots.push(Value::Object(meta));
        }
    }
    pivots
}

// ---------------------------------------------------------------------------
// Public API — the `XlsxWorkbook` pyclass.
// ---------------------------------------------------------------------------

/// Mirrors `xlsb_reader._xlsx_reader.XlsxWorkbook`.
///
/// Holds the source file path and lazily parses it into a [`WorkbookData`]
/// on first use (cached for the lifetime of the object).
#[pyclass(module = "xlsb_reader_rs")]
pub struct XlsxWorkbook {
    path: PathBuf,
    data: Mutex<Option<WorkbookData>>,
}

impl XlsxWorkbook {
    fn ensure_parsed(&self) -> Result<()> {
        let mut guard = self.data.lock().expect("XlsxWorkbook mutex poisoned");
        if guard.is_none() {
            *guard = Some(parse_xlsx(&self.path)?);
        }
        Ok(())
    }
}

#[pymethods]
impl XlsxWorkbook {
    /// `path` accepts anything `os.PathLike` (`str` or `pathlib.Path`),
    /// same as `XlsxWorkbook.__init__(self, path: "os.PathLike")` in
    /// `_xlsx_reader.py`.
    #[new]
    fn new(path: PathBuf) -> Self {
        XlsxWorkbook {
            path,
            data: Mutex::new(None),
        }
    }

    #[getter]
    fn sheet_names(&self) -> PyResult<Vec<String>> {
        self.ensure_parsed()?;
        let guard = self.data.lock().expect("XlsxWorkbook mutex poisoned");
        Ok(guard.as_ref().expect("just parsed").sheet_names.clone())
    }

    fn iter_formulas(&self, py: Python<'_>) -> PyResult<Py<PyList>> {
        self.ensure_parsed()?;
        let guard = self.data.lock().expect("XlsxWorkbook mutex poisoned");
        common::formulas_to_py(py, &guard.as_ref().expect("just parsed").formulas)
    }

    fn iter_values(&self, py: Python<'_>) -> PyResult<Py<PyList>> {
        self.ensure_parsed()?;
        let guard = self.data.lock().expect("XlsxWorkbook mutex poisoned");
        common::values_to_py(py, &guard.as_ref().expect("just parsed").values)
    }

    /// NOTE: unlike `XlsbWorkbook.iter_filters()`, the pure-Python
    /// `XlsxWorkbook.iter_filters()` yields plain filter dicts (each
    /// already carrying its own `"sheet"` key) and only for sheets that
    /// actually have an `<autoFilter>` — it does *not* yield a
    /// `(sheet_name, None)` pair for every other sheet. Reproduce that
    /// here by dropping the `None` entries and unwrapping the tuple.
    fn iter_filters(&self, py: Python<'_>) -> PyResult<Py<PyList>> {
        self.ensure_parsed()?;
        let guard = self.data.lock().expect("XlsxWorkbook mutex poisoned");
        let data = guard.as_ref().expect("just parsed");
        let list = PyList::empty(py);
        for (_sheet, info) in &data.filters {
            if let Some(v) = info {
                list.append(common::json_value_to_py(py, v)?)?;
            }
        }
        Ok(list.into())
    }

    fn iter_pivot_tables(&self, py: Python<'_>) -> PyResult<Py<PyList>> {
        self.ensure_parsed()?;
        let guard = self.data.lock().expect("XlsxWorkbook mutex poisoned");
        common::pivots_to_py(py, &guard.as_ref().expect("just parsed").pivots)
    }

    /// `Dict[str, str]` — see `XlsxWorkbook.iter_vba_modules` in
    /// `_xlsx_reader.py`. Returns `{}` when there is no embedded VBA
    /// project, same as Python.
    fn iter_vba_modules(&self, py: Python<'_>) -> PyResult<Py<PyDict>> {
        self.ensure_parsed()?;
        let guard = self.data.lock().expect("XlsxWorkbook mutex poisoned");
        common::vba_to_py(py, &guard.as_ref().expect("just parsed").vba)
    }

    fn close(&self) -> PyResult<()> {
        Ok(())
    }

    fn __enter__(slf: PyRef<'_, Self>) -> PyRef<'_, Self> {
        slf
    }

    #[pyo3(signature = (*_args, **_kwargs))]
    fn __exit__(
        &self,
        _args: &Bound<'_, PyTuple>,
        _kwargs: Option<&Bound<'_, PyDict>>,
    ) -> PyResult<bool> {
        self.close()?;
        Ok(false)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn xml_tree_resolves_entities_in_text() {
        // Regression test: quick-xml emits `&gt;`/`&lt;` as separate
        // `GeneralRef` events rather than folding them into the
        // surrounding `Text` event. Before handling `GeneralRef`
        // explicitly, this collapsed to "AND(A30,A40)" (both entities
        // silently dropped).
        let xml = br#"<f>AND(A3&gt;0,A4&lt;0)</f>"#;
        let tree = parse_xml_tree(xml).expect("valid xml");
        assert_eq!(tree.text, "AND(A3>0,A4<0)");
    }

    #[test]
    fn xml_tree_resolves_numeric_char_ref() {
        let xml = b"<t>a&#65;b&#x42;c</t>";
        let tree = parse_xml_tree(xml).expect("valid xml");
        assert_eq!(tree.text, "aAbBc");
    }

    #[test]
    fn shared_strings_rich_text_runs_concatenate() {
        let xml = br#"<sst><si><r><t>Hello </t></r><r><t>World</t></r></si><si><t>Plain</t></si></sst>"#;
        assert_eq!(parse_shared_strings(xml), vec!["Hello World", "Plain"]);
    }

    #[test]
    fn cell_ref_parsing() {
        assert_eq!(cell_ref_to_row_col("A1"), (0, 0));
        assert_eq!(cell_ref_to_row_col("K3"), (2, 10));
        assert_eq!(cell_ref_to_row_col("AA10"), (9, 26));
    }

    #[test]
    fn numeric_cell_value_whole_number_becomes_int() {
        assert_eq!(numeric_cell_value("42"), CellValue::Int(42));
        assert_eq!(numeric_cell_value("42.5"), CellValue::Float(42.5));
        assert_eq!(
            numeric_cell_value("not-a-number"),
            CellValue::Str("not-a-number".to_string())
        );
    }

    #[test]
    fn expand_shared_formula_shifts_relative_refs_only() {
        assert_eq!(expand_shared_formula("A1+B1", 1, 0), "A2+B2");
        assert_eq!(expand_shared_formula("$A$1+B1", 1, 1), "$A$1+C2");
        assert_eq!(expand_shared_formula("A1", 0, 0), "A1");
    }

    #[test]
    fn parse_worksheet_values_typed_cells() {
        let xml = br#"<worksheet><sheetData>
            <row r="1">
                <c r="A1" t="s"><v>0</v></c>
                <c r="B1" t="b"><v>1</v></c>
                <c r="C1"><v>3.5</v></c>
                <c r="D1" t="str"><v>result</v></c>
                <c r="E1" t="inlineStr"><is><t>inline</t></is></c>
            </row>
        </sheetData></worksheet>"#;
        let sst = vec!["shared0".to_string()];
        let values = parse_worksheet_values(xml, &sst);
        assert_eq!(values[&(0, 0)], CellValue::Str("shared0".to_string()));
        assert_eq!(values[&(0, 1)], CellValue::Bool(true));
        assert_eq!(values[&(0, 2)], CellValue::Float(3.5));
        assert_eq!(values[&(0, 3)], CellValue::Str("result".to_string()));
        assert_eq!(values[&(0, 4)], CellValue::Str("inline".to_string()));
    }

    #[test]
    fn parse_auto_filter_custom_condition() {
        let xml = br#"<autoFilter ref="A1:B10"><filterColumn colId="0">
            <customFilters><customFilter operator="greaterThan" val="10"/></customFilters>
        </filterColumn></autoFilter>"#;
        let tree = parse_xml_tree(xml).unwrap();
        let result = parse_auto_filter(&tree, "Sheet1");
        assert_eq!(result["sheet"], "Sheet1");
        assert_eq!(result["ref"], "A1:B10");
        assert_eq!(result["columns"][0]["type"], "custom");
        assert_eq!(result["columns"][0]["conditions"][0]["operator"], "greaterThan");
    }
}
