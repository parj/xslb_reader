//! spec.rs — the token-efficient structural spec extractor
//! (`xlsb-extract-spec`).
//!
//! Port of `xlsb_reader/_spec_extractor.py`'s `build_spec` pipeline (the
//! `--validate` CLI mode, `validate_spec`, is CLI-only and out of scope
//! here — see module doc history). Backs the crate's `extract_spec`
//! `#[pyfunction]` in `lib.rs`, the fast path `xlsb_reader.extract_spec`
//! uses when this extension is installed. Output must match
//! `_spec_extractor.build_spec` byte-for-byte for the same
//! `sample_rows`/`sheets` arguments, modulo the `extracted_at` timestamp
//! line (inherently non-deterministic in both implementations, and
//! rendered here in UTC via a hand-rolled calendar conversion rather than
//! Python's local time — see [`render_header`] — since matching the exact
//! clock value was never the point).

use std::collections::{BTreeSet, HashMap, HashSet};

use serde_json::{Map, Value};

use crate::common::{self, CellValue, Result, WorkbookData};
use crate::render::python_repr;
use crate::xlsb::python_float_repr;

/// `(formulas_by_sheet, dependencies)`, each a list of `(sheet_name, ...)`
/// pairs in workbook-tab order — see [`extract_formulas`]. Each sheet's
/// columns are `(col_key, entry)` pairs in *first-column-encountered*
/// order while scanning that sheet's formulas in `(row, col)` order —
/// matching the insertion order of Python's `sheet_json` dict, which
/// `render_hints`' "Key transformation" lines render sequentially (the
/// `fenced_json("formulas", ...)` output doesn't care about this order,
/// since JSON object keys get sorted at serialization either way).
type FormulasAndDeps = (
    Vec<(String, Vec<(String, Value)>)>,
    Vec<(String, Vec<String>)>,
);

const MIN_SCAN_ROWS: u32 = 500;
const MAX_SAMPLE_VALUES: usize = 5;

// ---------------------------------------------------------------------------
// Small helpers — port of the top-of-file helpers in `_spec_extractor.py`.
// ---------------------------------------------------------------------------

/// Convert Excel column letters to a 0-based index, e.g. `"A" -> 0`,
/// `"AA" -> 26`. Port of `letters_to_col`.
fn letters_to_col(letters: &str) -> i64 {
    let mut n: i64 = 0;
    for ch in letters.chars() {
        n = n * 26 + (ch as i64 - 64);
    }
    n - 1
}

fn cell_addr(row: u32, col: u32) -> String {
    format!("{}{}", common::col_to_letter(col), row + 1)
}

/// `str(value)` for a raw cell value — see `render::cellvalue_str`'s doc
/// comment for why floats specifically need `python_float_repr` rather
/// than any JSON-oriented formatting.
fn py_str(v: &CellValue) -> String {
    match v {
        CellValue::Int(i) => i.to_string(),
        CellValue::Float(f) => python_float_repr(*f),
        CellValue::Str(s) => s.clone(),
        CellValue::Bool(b) => {
            if *b {
                "True".to_string()
            } else {
                "False".to_string()
            }
        }
    }
}

/// `v is not None and v != ""` negated — matches Python's `non_null`
/// filter used throughout the schema-inference pipeline (a present cell
/// whose value is the empty string counts as "null" for these purposes,
/// but is still rendered literally, not as `"null"`, wherever raw values
/// are sampled — see [`build_column_schema`]).
fn is_null_ish(v: Option<&CellValue>) -> bool {
    match v {
        None => true,
        Some(CellValue::Str(s)) => s.is_empty(),
        _ => false,
    }
}

/// Port of `csv_field`.
fn csv_field(s: &str) -> String {
    if s.contains(',') || s.contains('"') || s.contains('\n') {
        format!("\"{}\"", s.replace('"', "\"\""))
    } else {
        s.to_string()
    }
}

/// Port of `toon_table`. `columns`/rows are pre-stringified (Python's
/// polymorphic `Any` cell values are converted to their `str()` form by
/// each caller before reaching here — see [`py_str`]/[`is_null_ish`]).
fn toon_table(name: &str, columns: &[String], rows: &[Vec<String>]) -> String {
    let mut lines = vec![format!("{name}[{}]{{{}}}:", rows.len(), columns.join(","))];
    for row in rows {
        lines.push(format!(
            "  {}",
            row.iter()
                .map(|v| csv_field(v))
                .collect::<Vec<_>>()
                .join(",")
        ));
    }
    lines.join("\n")
}

/// Port of `fenced_json`: compact (no extra whitespace), sorted-key JSON —
/// matches `json.dumps(obj, ensure_ascii=False, separators=(",", ":"),
/// sort_keys=True)`. serde_json's default `to_string` is already compact
/// with no extra whitespace, and `Map`/`Value::Object` is a `BTreeMap`
/// under the hood (the `preserve_order` feature is not enabled), so keys
/// come out sorted with no extra work.
fn fenced_json(label: &str, obj: &Value) -> String {
    let body = serde_json::to_string(obj).unwrap_or_else(|_| "null".to_string());
    format!("```json:{label}\n{body}\n```")
}

/// Port of `fenced_vba`.
fn fenced_vba(name: &str, src: &str) -> String {
    format!("```vba:{name}\n{}\n```", src.trim_end())
}

fn opt_i64_to_value(v: Option<i64>) -> Value {
    v.map(Value::from).unwrap_or(Value::Null)
}

fn opt_string_to_value(v: Option<String>) -> Value {
    v.map(Value::String).unwrap_or(Value::Null)
}

/// Python truthiness for a `serde_json::Value` (`None`, `False`, `0`,
/// `""`, `[]`, `{}` are all falsy) — used for the several `if x.get(y):`
/// / `x.get(y) or default` checks throughout the pivot/filter normalizers.
fn truthy(v: &Value) -> bool {
    match v {
        Value::Null => false,
        Value::Bool(b) => *b,
        Value::Number(n) => n.as_f64().map(|f| f != 0.0).unwrap_or(true),
        Value::String(s) => !s.is_empty(),
        Value::Array(a) => !a.is_empty(),
        Value::Object(o) => !o.is_empty(),
    }
}

// ---------------------------------------------------------------------------
// ISO-date validation — port of `_is_iso_date` (`ISO_DATE_RE` + `date.fromisoformat`).
// ---------------------------------------------------------------------------

fn is_leap_year(y: i32) -> bool {
    (y % 4 == 0 && y % 100 != 0) || y % 400 == 0
}

fn days_in_month(y: i32, m: u32) -> u32 {
    match m {
        1 | 3 | 5 | 7 | 8 | 10 | 12 => 31,
        4 | 6 | 9 | 11 => 30,
        2 => {
            if is_leap_year(y) {
                29
            } else {
                28
            }
        }
        _ => 0,
    }
}

/// Checks that `s` *starts with* a valid `YYYY-MM-DD` date (extra trailing
/// characters, e.g. a time component, are permitted — matches
/// `ISO_DATE_RE.match` + `date.fromisoformat(v[:10])`).
fn is_iso_date(s: &str) -> bool {
    let b = s.as_bytes();
    if b.len() < 10 {
        return false;
    }
    let digit = |i: usize| b[i].is_ascii_digit();
    if !(digit(0) && digit(1) && digit(2) && digit(3)) || b[4] != b'-' {
        return false;
    }
    if !(digit(5) && digit(6)) || b[7] != b'-' {
        return false;
    }
    if !(digit(8) && digit(9)) {
        return false;
    }
    let year: i32 = s[0..4].parse().unwrap_or(0);
    let month: u32 = s[5..7].parse().unwrap_or(0);
    let day: u32 = s[8..10].parse().unwrap_or(0);
    (1..=12).contains(&month) && (1..=days_in_month(year, month)).contains(&day)
}

// ---------------------------------------------------------------------------
// Step 2: per-sheet column schema
// ---------------------------------------------------------------------------

fn is_num_val(v: &CellValue) -> bool {
    matches!(v, CellValue::Int(_) | CellValue::Float(_))
}

fn is_whole_num(v: &CellValue) -> bool {
    match v {
        CellValue::Int(_) => true,
        CellValue::Float(f) => f.fract() == 0.0,
        _ => false,
    }
}

/// Port of `infer_column_type`.
fn infer_column_type(non_null: &[&CellValue]) -> &'static str {
    if non_null.is_empty() {
        return "string";
    }
    if non_null.iter().all(|v| is_num_val(v) && is_whole_num(v)) {
        return "integer";
    }
    if non_null.iter().all(|v| is_num_val(v)) {
        return "float";
    }
    if non_null
        .iter()
        .all(|v| matches!(v, CellValue::Str(s) if is_iso_date(s)))
    {
        return "date";
    }
    if non_null.iter().all(|v| matches!(v, CellValue::Bool(_))) {
        return "boolean";
    }
    "string"
}

/// Port of `sheet_dimensions`.
fn sheet_dimensions(values: &common::CellMap<CellValue>) -> (u32, u32) {
    if values.is_empty() {
        return (0, 0);
    }
    let max_row = values.keys().map(|&(r, _)| r).max().unwrap();
    let max_col = values.keys().map(|&(_, c)| c).max().unwrap();
    (max_row + 1, max_col + 1)
}

/// Port of `detect_headers`.
fn detect_headers(values: &common::CellMap<CellValue>, ncols: u32) -> (HashMap<u32, String>, u32) {
    let has_header = (0..ncols).any(|c| !is_null_ish(values.get(&(0, c))));
    let mut header_names = HashMap::new();
    for c in 0..ncols {
        let v = if has_header {
            values.get(&(0, c))
        } else {
            None
        };
        let name = match v {
            Some(cv) if !is_null_ish(Some(cv)) => py_str(cv),
            _ => format!("col_{}", common::col_to_letter(c)),
        };
        header_names.insert(c, name);
    }
    (header_names, u32::from(has_header))
}

/// A Python-numeric-tower-aware dedup key: `int`/`float`/`bool` collapse
/// onto the same key when numerically equal (mirroring `1 == 1.0 == True`
/// in Python, which a plain `set()` of raw values would also collapse),
/// while strings key separately.
#[derive(PartialEq, Eq, Hash)]
enum DistinctKey {
    Num(u64),
    Str(String),
}

fn distinct_key(v: &CellValue) -> DistinctKey {
    match v {
        CellValue::Int(i) => DistinctKey::Num((*i as f64).to_bits()),
        CellValue::Float(f) => {
            let f = if *f == 0.0 { 0.0 } else { *f };
            DistinctKey::Num(f.to_bits())
        }
        CellValue::Bool(b) => DistinctKey::Num((if *b { 1.0f64 } else { 0.0f64 }).to_bits()),
        CellValue::Str(s) => DistinctKey::Str(s.clone()),
    }
}

struct ColumnSchema {
    name: String,
    inferred_type: &'static str,
    nullable: bool,
    sample_values: String,
    notes: String,
}

/// Port of `build_column_schema`.
fn build_column_schema(
    values: &common::CellMap<CellValue>,
    ncols: u32,
    header_names: &HashMap<u32, String>,
    data_start_row: u32,
    nrows: u32,
    scan_rows: u32,
) -> Vec<ColumnSchema> {
    let window_end = nrows.min(data_start_row + scan_rows);
    let mut columns = Vec::new();
    for c in 0..ncols {
        let col_values: Vec<Option<&CellValue>> = (data_start_row..window_end)
            .map(|r| values.get(&(r, c)))
            .collect();
        let non_null_vals: Vec<&CellValue> = col_values
            .iter()
            .filter_map(|v| if is_null_ish(*v) { None } else { *v })
            .collect();
        let null_count = col_values.len() - non_null_vals.len();
        let col_type = infer_column_type(&non_null_vals);

        let mut samples: Vec<String> = Vec::new();
        for v in &col_values {
            if samples.len() >= MAX_SAMPLE_VALUES {
                break;
            }
            match v {
                None => samples.push("null".to_string()),
                Some(CellValue::Float(f)) if col_type == "integer" => {
                    samples.push((*f as i64).to_string())
                }
                Some(cv) => samples.push(py_str(cv)),
            }
        }

        let mut notes: Vec<String> = Vec::new();
        let distinct: HashSet<DistinctKey> =
            non_null_vals.iter().map(|v| distinct_key(v)).collect();
        if c == 0
            && null_count == 0
            && !non_null_vals.is_empty()
            && distinct.len() == non_null_vals.len()
        {
            notes.push("primary key candidate".to_string());
        } else if col_type == "string"
            && !distinct.is_empty()
            && distinct.len() <= 10
            && distinct.len() < non_null_vals.len()
        {
            notes.push(format!("{} distinct values seen", distinct.len()));
        }
        if (col_type == "integer" || col_type == "float")
            && header_names[&c].to_lowercase().contains("date")
        {
            notes.push("possible Excel serial date".to_string());
        }

        columns.push(ColumnSchema {
            name: header_names[&c].clone(),
            inferred_type: col_type,
            nullable: null_count > 0,
            sample_values: samples.join(","),
            notes: notes.join("; "),
        });
    }
    columns
}

/// Port of `render_sheet_schema`.
fn render_sheet_schema(
    sheet_name: &str,
    nrows: u32,
    ncols: u32,
    columns: &[ColumnSchema],
) -> String {
    let mut lines = vec![
        format!("sheet: {sheet_name}"),
        format!("dimensions: {nrows} rows x {ncols} cols"),
    ];
    let headers = [
        "name".to_string(),
        "inferred_type".to_string(),
        "nullable".to_string(),
        "sample_values".to_string(),
        "notes".to_string(),
    ];
    let rows: Vec<Vec<String>> = columns
        .iter()
        .map(|c| {
            vec![
                c.name.clone(),
                c.inferred_type.to_string(),
                if c.nullable { "true" } else { "false" }.to_string(),
                c.sample_values.clone(),
                c.notes.clone(),
            ]
        })
        .collect();
    lines.push(toon_table("columns", &headers, &rows));
    lines.join("\n")
}

/// Port of `render_sheet_sample`.
fn render_sheet_sample(
    sheet_name: &str,
    values: &common::CellMap<CellValue>,
    header_names: &HashMap<u32, String>,
    data_start_row: u32,
    nrows: u32,
    ncols: u32,
    sample_rows: u32,
) -> String {
    let n = sample_rows.min(nrows.saturating_sub(data_start_row));
    let col_order: Vec<String> = (0..ncols).map(|c| header_names[&c].clone()).collect();
    let mut rows: Vec<Vec<String>> = Vec::new();
    for r in data_start_row..(data_start_row + n) {
        let row: Vec<String> = (0..ncols)
            .map(|c| match values.get(&(r, c)) {
                None => String::new(),
                Some(v) => py_str(v),
            })
            .collect();
        rows.push(row);
    }
    let mut lines = vec![format!("sample: {sheet_name}")];
    lines.push(toon_table("rows", &col_order, &rows));
    lines.join("\n")
}

// ---------------------------------------------------------------------------
// Step 3/4: formula pattern extraction + inter-sheet dependencies
// ---------------------------------------------------------------------------

/// One match of the hand-rolled `CELL_REF_RE` equivalent (see
/// [`find_cell_refs`]'s doc comment for why this isn't the `regex` crate).
struct CellRefMatch {
    start: usize,
    end: usize,
    qsheet: Option<String>,
    sheet: Option<String>,
    col: String,
    row: String,
}

fn is_word_byte(b: u8) -> bool {
    b.is_ascii_alphanumeric() || b == b'_'
}

/// Try to match `\$?[A-Z]{1,3}\$?\d{1,7}` starting exactly at byte offset
/// `pos`, including the trailing `(?![A-Za-z0-9_(])` lookahead. Both the
/// letter-run and digit-run quantifiers are effectively unambiguous here
/// (letters and digits are disjoint character classes with no
/// interleaving), so a single greedy pass need not backtrack — see
/// [`find_cell_refs`]'s doc comment for the full argument.
fn match_col_row(bytes: &[u8], pos: usize) -> Option<(usize, String, String)> {
    let n = bytes.len();
    let mut i = pos;
    if i < n && bytes[i] == b'$' {
        i += 1;
    }
    let letters_start = i;
    while i < n && bytes[i].is_ascii_uppercase() {
        i += 1;
    }
    let letters_len = i - letters_start;
    if letters_len == 0 || letters_len > 3 {
        return None;
    }
    if i < n && bytes[i] == b'$' {
        i += 1;
    }
    let digits_start = i;
    while i < n && bytes[i].is_ascii_digit() {
        i += 1;
    }
    let digits_len = i - digits_start;
    if digits_len == 0 || digits_len > 7 {
        return None;
    }
    if i < n {
        let b = bytes[i];
        if b.is_ascii_alphanumeric() || b == b'_' || b == b'(' {
            return None;
        }
    }
    let col =
        String::from_utf8_lossy(&bytes[letters_start..letters_start + letters_len]).into_owned();
    let row = String::from_utf8_lossy(&bytes[digits_start..digits_start + digits_len]).into_owned();
    Some((i, col, row))
}

/// Hand-rolled equivalent of Python's `CELL_REF_RE`:
/// ```text
/// (?<![A-Za-z0-9_])
/// (?:(?:'(?P<qsheet>[^']+)'|(?P<sheet>[A-Za-z_][A-Za-z0-9_.]*))!)?
/// \$?(?P<col>[A-Z]{1,3})\$?(?P<row>\d{1,7})
/// (?![A-Za-z0-9_(])
/// ```
/// The `regex` crate (used elsewhere in this crate) cannot express the
/// look-behind/look-ahead assertions this pattern needs — Rust's `regex`
/// is a linear-time automaton with no backtracking/look-around support.
/// Rather than pull in `fancy-regex` (a backtracking engine, and another
/// dependency) for one pattern, this scans manually. The pattern's pieces
/// use disjoint character classes in strict sequence (letters, then
/// digits; quote or bareword-then-`!`, or neither), which makes every
/// quantifier's greedy match the *only* candidate a backtracking engine
/// could ever settle on too — see the inline comments on the trickier
/// cases (4+ letters/8+ digits; a bareword sheet-prefix candidate whose
/// tail also happens to look like a standalone reference).
fn find_cell_refs(s: &str) -> Vec<CellRefMatch> {
    let bytes = s.as_bytes();
    let n = bytes.len();
    let mut matches = Vec::new();
    let mut i = 0usize;
    while i < n {
        if i > 0 && is_word_byte(bytes[i - 1]) {
            i += 1;
            continue;
        }

        let mut qsheet: Option<String> = None;
        let mut sheet: Option<String> = None;
        let mut prefix_end: Option<usize> = None;

        if bytes[i] == b'\'' {
            if let Some(rel) = bytes[i + 1..].iter().position(|&b| b == b'\'') {
                let close = i + 1 + rel;
                if close > i + 1 && close + 1 < n && bytes[close + 1] == b'!' {
                    qsheet = Some(s[i + 1..close].to_string());
                    prefix_end = Some(close + 2);
                }
            }
        } else if bytes[i].is_ascii_alphabetic() || bytes[i] == b'_' {
            let mut j = i + 1;
            while j < n
                && (bytes[j].is_ascii_alphanumeric() || bytes[j] == b'_' || bytes[j] == b'.')
            {
                j += 1;
            }
            if j < n && bytes[j] == b'!' {
                sheet = Some(s[i..j].to_string());
                prefix_end = Some(j + 1);
            }
        }

        // Prefer the sheet-prefixed interpretation if it leads to a full
        // match; otherwise (no prefix found, or prefix found but the
        // mandatory col/row part after it doesn't match) fall back to a
        // plain reference starting at `i` itself — matches a backtracking
        // engine giving up on the optional group when it doesn't lead to
        // overall success.
        let matched = prefix_end
            .and_then(|pe| {
                match_col_row(bytes, pe)
                    .map(|(end, col, row)| (end, qsheet.clone(), sheet.clone(), col, row))
            })
            .or_else(|| match_col_row(bytes, i).map(|(end, col, row)| (end, None, None, col, row)));

        match matched {
            Some((end, qs, sh, col, row)) => {
                matches.push(CellRefMatch {
                    start: i,
                    end,
                    qsheet: qs,
                    sheet: sh,
                    col,
                    row,
                });
                i = end;
            }
            None => i += 1,
        }
    }
    matches
}

/// Port of `normalize_formula`.
fn normalize_formula(
    formula: &str,
    cur_sheet: &str,
    cur_row: i64,
    header_maps: &HashMap<String, HashMap<i64, String>>,
) -> (String, BTreeSet<String>) {
    let mut referenced_sheets = BTreeSet::new();
    let matches = find_cell_refs(formula);
    let mut out = String::with_capacity(formula.len());
    let mut last = 0usize;
    for m in &matches {
        out.push_str(&formula[last..m.start]);
        let sheet_ref = m.qsheet.clone().or_else(|| m.sheet.clone());
        let col_idx = letters_to_col(&m.col);
        let lookup_sheet = sheet_ref.clone().unwrap_or_else(|| cur_sheet.to_string());
        let name = header_maps
            .get(&lookup_sheet)
            .and_then(|hm| hm.get(&col_idx))
            .cloned();
        let token = match name {
            Some(n) => n,
            None => common::col_to_letter(col_idx.max(0) as u32),
        };
        if let Some(sref) = &sheet_ref {
            referenced_sheets.insert(sref.clone());
            out.push_str(&format!("<{sref}!{token}>"));
        } else {
            let row_num: i64 = m.row.parse().unwrap_or(1);
            let offset = (row_num - 1) - cur_row;
            if offset == 0 {
                out.push_str(&format!("<{token}>"));
            } else {
                let sign = if offset > 0 { "+" } else { "" };
                out.push_str(&format!("<{token}{sign}{offset}>"));
            }
        }
        last = m.end;
    }
    out.push_str(&formula[last..]);
    (out, referenced_sheets)
}

/// Increment (or insert) `pattern`'s count, preserving first-insertion
/// order — mirrors `collections.Counter`'s dict-based storage (needed for
/// [`most_common_first`]'s tie-breaking).
fn counter_increment(counter: &mut Vec<(String, u32)>, pattern: &str) {
    match counter.iter_mut().find(|(p, _)| p == pattern) {
        Some(entry) => entry.1 += 1,
        None => counter.push((pattern.to_string(), 1)),
    }
}

/// Port of `counter.most_common(1)[0]`: the highest count, ties broken by
/// whichever pattern was inserted first (matches `Counter.most_common`'s
/// stable-sort tie-breaking).
fn most_common_first(counter: &[(String, u32)]) -> (&str, u32) {
    let mut best = &counter[0];
    for entry in &counter[1..] {
        if entry.1 > best.1 {
            best = entry;
        }
    }
    (best.0.as_str(), best.1)
}

struct ColFormulaStats {
    col: u32,
    patterns: Vec<(String, u32)>,
    examples: HashMap<String, (String, String)>,
    parse_errors: u32,
}

/// Port of `extract_formulas`. Returns `(formulas_by_sheet,
/// dependencies)` in workbook-tab order (matching `wb.iter_formulas()`'s
/// iteration order, which Python's `formulas_by_sheet`/`dependencies`
/// dicts preserve via insertion order).
fn extract_formulas(
    data: &WorkbookData,
    header_maps: &HashMap<String, HashMap<i64, String>>,
    sheet_selected: &impl Fn(&str) -> bool,
) -> FormulasAndDeps {
    let mut formulas_by_sheet: Vec<(String, Vec<(String, Value)>)> = Vec::new();
    let mut dependencies: Vec<(String, Vec<String>)> = Vec::new();

    for (sheet_name, formulas) in &data.formulas {
        if formulas.is_empty() || !sheet_selected(sheet_name) {
            continue;
        }

        let mut col_stats: Vec<ColFormulaStats> = Vec::new();
        let mut col_index: HashMap<u32, usize> = HashMap::new();
        let mut deps: BTreeSet<String> = BTreeSet::new();

        let get_or_create = |col: u32,
                             col_stats: &mut Vec<ColFormulaStats>,
                             col_index: &mut HashMap<u32, usize>|
         -> usize {
            *col_index.entry(col).or_insert_with(|| {
                col_stats.push(ColFormulaStats {
                    col,
                    patterns: Vec::new(),
                    examples: HashMap::new(),
                    parse_errors: 0,
                });
                col_stats.len() - 1
            })
        };

        for (&(row, col), formula) in formulas.iter() {
            if formula.starts_with("=<parse_error:") {
                let idx = get_or_create(col, &mut col_stats, &mut col_index);
                col_stats[idx].parse_errors += 1;
                continue;
            }
            let (pattern, refd) = normalize_formula(formula, sheet_name, row as i64, header_maps);
            deps.extend(refd);
            let idx = get_or_create(col, &mut col_stats, &mut col_index);
            counter_increment(&mut col_stats[idx].patterns, &pattern);
            col_stats[idx]
                .examples
                .entry(pattern)
                .or_insert_with(|| (cell_addr(row, col), formula.clone()));
        }

        deps.remove(sheet_name);
        if !deps.is_empty() {
            dependencies.push((sheet_name.clone(), deps.into_iter().collect()));
        }

        let mut sheet_json: Vec<(String, Value)> = Vec::new();
        for stats in &col_stats {
            if stats.patterns.is_empty() {
                continue;
            }
            let (dominant_pattern, dominant_count) = most_common_first(&stats.patterns);
            let dominant_pattern = dominant_pattern.to_string();
            let mut entry = Map::new();
            entry.insert(
                "pattern".to_string(),
                Value::String(dominant_pattern.clone()),
            );
            let (ex_cell, ex_formula) = &stats.examples[&dominant_pattern];
            entry.insert("example_cell".to_string(), Value::String(ex_cell.clone()));
            entry.insert(
                "example_formula".to_string(),
                Value::String(ex_formula.clone()),
            );
            entry.insert("row_count".to_string(), Value::from(dominant_count));
            let other_count: u32 = stats
                .patterns
                .iter()
                .filter(|(p, _)| p != &dominant_pattern)
                .map(|(_, c)| c)
                .sum();
            if other_count > 0 {
                entry.insert("other_pattern_count".to_string(), Value::from(other_count));
            }
            if stats.parse_errors > 0 {
                entry.insert(
                    "parse_error_count".to_string(),
                    Value::from(stats.parse_errors),
                );
            }
            sheet_json.push((
                format!("col_{}", common::col_to_letter(stats.col)),
                Value::Object(entry),
            ));
        }
        for stats in &col_stats {
            if stats.patterns.is_empty() && stats.parse_errors > 0 {
                let mut entry = Map::new();
                entry.insert("parse_error".to_string(), Value::Bool(true));
                entry.insert("count".to_string(), Value::from(stats.parse_errors));
                sheet_json.push((
                    format!("col_{}", common::col_to_letter(stats.col)),
                    Value::Object(entry),
                ));
            }
        }

        if !sheet_json.is_empty() {
            formulas_by_sheet.push((sheet_name.clone(), sheet_json));
        }
    }

    (formulas_by_sheet, dependencies)
}

// ---------------------------------------------------------------------------
// Step 5/6: pivot tables + autofilters
// ---------------------------------------------------------------------------

/// Port of `normalize_pivots`.
fn normalize_pivots(data: &WorkbookData, sheet_selected: &impl Fn(&str) -> bool) -> Vec<Value> {
    let mut out = Vec::new();
    for pt in &data.pivots {
        if let Some(s) = pt.get("sheet").and_then(Value::as_str) {
            if !sheet_selected(s) {
                continue;
            }
        }

        let location_str: Option<String> =
            pt.get("location")
                .filter(|l| l.is_object())
                .and_then(|loc| {
                    let geom = loc.get("rfx_geom");
                    let tl = geom.and_then(|g| g.get("top_left")).and_then(Value::as_str);
                    let br = geom
                        .and_then(|g| g.get("bottom_right"))
                        .and_then(Value::as_str);
                    match (tl, br) {
                        (Some(tl), Some(br)) if !tl.is_empty() && !br.is_empty() => {
                            Some(format!("{tl}:{br}"))
                        }
                        _ => None,
                    }
                });

        let empty_vec: Vec<Value> = Vec::new();
        let cache_fields: &Vec<Value> = pt
            .get("cache_fields")
            .and_then(Value::as_array)
            .unwrap_or(&empty_vec);

        let field_name = |idx: Option<i64>| -> Option<String> {
            let idx = idx?;
            let field = cache_fields.get(usize::try_from(idx).ok()?)?;
            field
                .get("name")
                .and_then(Value::as_str)
                .map(str::to_string)
        };
        let field_value = |idx: Option<i64>, item_idx: Option<i64>| -> Option<Value> {
            let idx = idx?;
            let item_idx = item_idx?;
            let field = cache_fields.get(usize::try_from(idx).ok()?)?;
            let items = field.get("shared_items").and_then(Value::as_array)?;
            items.get(usize::try_from(item_idx).ok()?).cloned()
        };

        let mut filters: Vec<Value> = Vec::new();
        if let Some(sx_filters) = pt.get("sx_filters").and_then(Value::as_array) {
            for sf in sx_filters {
                if let Some(criteria) = sf.get("criteria").and_then(Value::as_array) {
                    let fidx = sf.get("field_index").and_then(Value::as_i64);
                    for crit in criteria {
                        let mut e = Map::new();
                        e.insert(
                            "kind".to_string(),
                            Value::String("value_filter".to_string()),
                        );
                        e.insert("field_index".to_string(), opt_i64_to_value(fidx));
                        e.insert(
                            "field_name".to_string(),
                            opt_string_to_value(field_name(fidx)),
                        );
                        e.insert(
                            "operator".to_string(),
                            crit.get("operator").cloned().unwrap_or(Value::Null),
                        );
                        e.insert(
                            "value".to_string(),
                            crit.get("value").cloned().unwrap_or(Value::Null),
                        );
                        filters.push(Value::Object(e));
                    }
                }
            }
        }
        if let Some(report_filters) = pt.get("report_filters").and_then(Value::as_array) {
            for rf in report_filters {
                let fidx = rf.get("fld").and_then(Value::as_i64);
                let operator = rf
                    .get("operator")
                    .filter(|v| !v.is_null())
                    .or_else(|| rf.get("type"))
                    .cloned()
                    .unwrap_or(Value::Null);
                let mut e = Map::new();
                e.insert(
                    "kind".to_string(),
                    Value::String("value_filter".to_string()),
                );
                e.insert("field_index".to_string(), opt_i64_to_value(fidx));
                e.insert(
                    "field_name".to_string(),
                    opt_string_to_value(field_name(fidx)),
                );
                e.insert("operator".to_string(), operator);
                e.insert(
                    "value".to_string(),
                    rf.get("value").cloned().unwrap_or(Value::Null),
                );
                filters.push(Value::Object(e));
            }
        }
        if let Some(field_filters) = pt.get("field_filters").and_then(Value::as_array) {
            for ff in field_filters {
                let fidx = ff.get("field_index").and_then(Value::as_i64);
                let mut excluded = Vec::new();
                if let Some(hidden) = ff.get("hidden_indices").and_then(Value::as_array) {
                    for idx_v in hidden {
                        if let Some(v) = field_value(fidx, idx_v.as_i64()) {
                            excluded.push(v);
                        }
                    }
                }
                let mut e = Map::new();
                e.insert("kind".to_string(), Value::String("item_filter".to_string()));
                e.insert("field_index".to_string(), opt_i64_to_value(fidx));
                e.insert(
                    "field_name".to_string(),
                    opt_string_to_value(field_name(fidx)),
                );
                e.insert("excluded_values".to_string(), Value::Array(excluded));
                filters.push(Value::Object(e));
            }
        }

        let mut entry = Map::new();
        entry.insert(
            "name".to_string(),
            pt.get("name").cloned().unwrap_or(Value::Null),
        );
        entry.insert(
            "sheet".to_string(),
            pt.get("sheet").cloned().unwrap_or(Value::Null),
        );
        entry.insert(
            "source_cache_id".to_string(),
            pt.get("cache_id").cloned().unwrap_or(Value::Null),
        );
        entry.insert("location".to_string(), opt_string_to_value(location_str));
        entry.insert(
            "fields".to_string(),
            pt.get("pivot_fields").cloned().unwrap_or(Value::Null),
        );
        if !filters.is_empty() {
            entry.insert("filters".to_string(), Value::Array(filters));
        }
        for key in ["data_fields", "row_fields", "col_fields"] {
            if let Some(v) = pt.get(key) {
                if truthy(v) {
                    entry.insert(key.to_string(), v.clone());
                }
            }
        }
        out.push(Value::Object(entry));
    }
    out
}

/// Port of `normalize_filters`. `data.filters` is always `(sheet_name,
/// info_or_None)` regardless of backend (see `WorkbookData::filters`'s
/// doc comment), so unlike Python this needs no `isinstance(item, tuple)`
/// branch — the sheet name always comes from our own tuple.
fn normalize_filters(
    data: &WorkbookData,
    sheet_selected: &impl Fn(&str) -> bool,
    header_maps: &HashMap<String, HashMap<i64, String>>,
) -> Vec<Value> {
    let mut out = Vec::new();
    for (sheet_name, info_opt) in &data.filters {
        let info = match info_opt {
            Some(v) => v,
            None => continue,
        };
        if !sheet_selected(sheet_name) {
            continue;
        }

        let (rng, top_left): (Option<String>, String) =
            if let Some(range_obj) = info.get("range").filter(|v| v.is_object()) {
                let tl = range_obj
                    .get("top_left")
                    .and_then(Value::as_str)
                    .unwrap_or("");
                let br = range_obj
                    .get("bottom_right")
                    .and_then(Value::as_str)
                    .unwrap_or("");
                (Some(format!("{tl}:{br}")), tl.to_string())
            } else if let Some(r) = info
                .get("ref")
                .and_then(Value::as_str)
                .filter(|s| !s.is_empty())
            {
                let tl = r.split(':').next().unwrap_or("").to_string();
                (Some(r.to_string()), tl)
            } else {
                (None, String::new())
            };
        let letters: String = top_left
            .chars()
            .take_while(|c| c.is_ascii_uppercase())
            .collect();
        let range_start_col: i64 = if letters.is_empty() {
            0
        } else {
            letters_to_col(&letters)
        };

        let empty_map = HashMap::new();
        let sheet_headers = header_maps.get(sheet_name).unwrap_or(&empty_map);

        let mut columns_out: Vec<Value> = Vec::new();
        if let Some(columns) = info.get("columns").and_then(Value::as_array) {
            for col in columns {
                let idx: Option<i64> = col
                    .get("column_index")
                    .and_then(Value::as_i64)
                    .or_else(|| col.get("col_id").and_then(Value::as_i64));
                let name: Value = match idx {
                    Some(i) => sheet_headers
                        .get(&(range_start_col + i))
                        .map(|s| Value::String(s.clone()))
                        .unwrap_or(Value::Null),
                    None => Value::Null,
                };
                let ctype = col.get("type").and_then(Value::as_str);
                let ctype_or = |default: &str| -> Value {
                    ctype
                        .map(|s| Value::String(s.to_string()))
                        .unwrap_or_else(|| Value::String(default.to_string()))
                };

                if let Some(custom) = col.get("custom_filters").filter(|v| truthy(v)) {
                    let empty_arr = Vec::new();
                    let crits = custom
                        .get("criteria")
                        .and_then(Value::as_array)
                        .unwrap_or(&empty_arr);
                    let logic = custom.get("logic");
                    for crit in crits {
                        let mut e = Map::new();
                        e.insert("index".to_string(), opt_i64_to_value(idx));
                        e.insert("name".to_string(), name.clone());
                        e.insert("type".to_string(), Value::String("custom".to_string()));
                        e.insert(
                            "operator".to_string(),
                            crit.get("operator").cloned().unwrap_or(Value::Null),
                        );
                        e.insert(
                            "value".to_string(),
                            crit.get("value").cloned().unwrap_or(Value::Null),
                        );
                        if let Some(l) = logic {
                            if truthy(l) && crits.len() > 1 {
                                e.insert("logic".to_string(), l.clone());
                            }
                        }
                        columns_out.push(Value::Object(e));
                    }
                } else if let Some(conditions) = col
                    .get("conditions")
                    .and_then(Value::as_array)
                    .filter(|a| !a.is_empty())
                {
                    for crit in conditions {
                        let mut e = Map::new();
                        e.insert("index".to_string(), opt_i64_to_value(idx));
                        e.insert("name".to_string(), name.clone());
                        e.insert("type".to_string(), ctype_or("custom"));
                        e.insert(
                            "operator".to_string(),
                            crit.get("operator").cloned().unwrap_or(Value::Null),
                        );
                        e.insert(
                            "value".to_string(),
                            crit.get("val").cloned().unwrap_or(Value::Null),
                        );
                        columns_out.push(Value::Object(e));
                    }
                } else if let Some(filters_arr) = col.get("filters").filter(|v| truthy(v)) {
                    let mut e = Map::new();
                    e.insert("index".to_string(), opt_i64_to_value(idx));
                    e.insert("name".to_string(), name.clone());
                    e.insert("type".to_string(), Value::String("discrete".to_string()));
                    e.insert("values".to_string(), filters_arr.clone());
                    columns_out.push(Value::Object(e));
                } else if let Some(values_arr) = col.get("values").filter(|v| truthy(v)) {
                    let mut e = Map::new();
                    e.insert("index".to_string(), opt_i64_to_value(idx));
                    e.insert("name".to_string(), name.clone());
                    e.insert("type".to_string(), ctype_or("discrete"));
                    e.insert("values".to_string(), values_arr.clone());
                    columns_out.push(Value::Object(e));
                } else if let Some(attrs) = col.get("attrs").filter(|v| truthy(v)) {
                    let mut e = Map::new();
                    e.insert("index".to_string(), opt_i64_to_value(idx));
                    e.insert("name".to_string(), name.clone());
                    e.insert("type".to_string(), ctype_or("unknown"));
                    e.insert("attrs".to_string(), attrs.clone());
                    columns_out.push(Value::Object(e));
                } else {
                    let mut e = Map::new();
                    e.insert("index".to_string(), opt_i64_to_value(idx));
                    e.insert("name".to_string(), name.clone());
                    e.insert("type".to_string(), ctype_or("unknown"));
                    columns_out.push(Value::Object(e));
                }
            }
        }

        let mut result = Map::new();
        result.insert("sheet".to_string(), Value::String(sheet_name.clone()));
        result.insert(
            "range".to_string(),
            rng.map(Value::String).unwrap_or(Value::Null),
        );
        result.insert("columns".to_string(), Value::Array(columns_out));
        out.push(Value::Object(result));
    }
    out
}

// ---------------------------------------------------------------------------
// Step 9: hints
// ---------------------------------------------------------------------------

/// `str(value)` for a `serde_json::Value`, Python-`repr`-style (used for
/// the pivot/filter hint lines, which embed raw JSON scalars/arrays
/// directly into an f-string).
fn json_str(v: &Value) -> String {
    match v {
        Value::String(s) => s.clone(),
        Value::Null => "None".to_string(),
        _ => python_repr(v),
    }
}

/// Port of `render_hints`.
#[allow(clippy::too_many_arguments)]
fn render_hints(
    sheet_dims: &[(String, (u32, u32))],
    formulas_by_sheet: &[(String, Vec<(String, Value)>)],
    header_maps: &HashMap<String, HashMap<i64, String>>,
    dependencies: &[(String, Vec<String>)],
    pivots: &[Value],
    filters: &[Value],
    vba_modules: &[(String, String)],
) -> String {
    let mut lines = vec!["[hints]".to_string()];

    if let Some((primary_name, (primary_rows, _))) =
        sheet_dims.iter().max_by_key(|(_, (rows, _))| *rows)
    {
        lines.push(format!(
            "- Primary table: {primary_name} ({primary_rows} rows, grain = one row per record)"
        ));
    }

    for (sheet, cols) in formulas_by_sheet {
        for (colkey, info) in cols {
            if info.get("parse_error").is_some() {
                continue;
            }
            let pattern = info.get("pattern").and_then(Value::as_str).unwrap_or("");
            let col_letters = colkey.split_once('_').map(|x| x.1).unwrap_or("");
            let col_idx = letters_to_col(col_letters);
            let colname = header_maps
                .get(sheet)
                .and_then(|hm| hm.get(&col_idx))
                .cloned()
                .unwrap_or_else(|| colkey.clone());
            let approach = if ["VLOOKUP", "INDEX", "MATCH"]
                .iter()
                .any(|f| pattern.contains(f))
            {
                "use pd.merge or dict map"
            } else {
                "vectorisable"
            };
            lines.push(format!(
                "- Key transformation: {sheet}.{colname} = {pattern} ({approach})"
            ));
        }
    }

    for (sheet, deps) in dependencies {
        lines.push(format!("- {sheet} depends on: {}", deps.join(", ")));
    }

    for pt in pivots {
        let mut f_desc = String::new();
        if let Some(pfilters) = pt
            .get("filters")
            .and_then(Value::as_array)
            .filter(|a| !a.is_empty())
        {
            let mut parts = Vec::new();
            for f in pfilters {
                let label = f
                    .get("field_name")
                    .and_then(Value::as_str)
                    .filter(|s| !s.is_empty())
                    .map(str::to_string)
                    .unwrap_or_else(|| {
                        format!(
                            "field {}",
                            json_str(f.get("field_index").unwrap_or(&Value::Null))
                        )
                    });
                if f.get("kind").and_then(Value::as_str) == Some("item_filter") {
                    let excluded = f
                        .get("excluded_values")
                        .cloned()
                        .unwrap_or(Value::Array(vec![]));
                    parts.push(format!("{label} excludes {}", python_repr(&excluded)));
                } else {
                    let operator = f.get("operator").map(json_str).unwrap_or_default();
                    let val = f.get("value").filter(|v| !v.is_null());
                    match val {
                        Some(v) => parts.push(format!("{label} {operator} {}", json_str(v))),
                        None => parts.push(format!("{label} {operator}")),
                    }
                }
            }
            f_desc = format!(" filters {}", parts.join("; "));
        }
        let name = pt.get("name").map(json_str).unwrap_or_default();
        let sheet = pt.get("sheet").map(json_str).unwrap_or_default();
        lines.push(format!(
            "- Pivot: {name} on {sheet}{f_desc} -> implement via pandas pivot_table()/groupby()"
        ));
    }

    for f in filters {
        if let Some(cols) = f.get("columns").and_then(Value::as_array) {
            for col in cols {
                let val = col
                    .get("value")
                    .filter(|v| !v.is_null())
                    .or_else(|| col.get("values"))
                    .map(json_str)
                    .unwrap_or_default();
                let label = col
                    .get("name")
                    .and_then(Value::as_str)
                    .filter(|s| !s.is_empty())
                    .map(str::to_string)
                    .unwrap_or_else(|| {
                        format!("col {}", json_str(col.get("index").unwrap_or(&Value::Null)))
                    });
                let sheet = f.get("sheet").map(json_str).unwrap_or_default();
                let operator = col.get("operator").map(json_str).unwrap_or_default();
                lines.push(format!(
                    "- AutoFilter on {sheet}.{label} {operator} {val} -> document as input filter parameter"
                ));
            }
        }
    }

    if !vba_modules.is_empty() {
        let names: Vec<&str> = vba_modules.iter().map(|(n, _)| n.as_str()).collect();
        lines.push(format!(
            "- VBA modules present: {} (see vba:... blocks for macro logic)",
            names.join(", ")
        ));
    } else {
        lines.push("- No VBA found".to_string());
    }

    lines.join("\n")
}

// ---------------------------------------------------------------------------
// Orchestration
// ---------------------------------------------------------------------------

/// Days since the Unix epoch -> proleptic Gregorian `(year, month, day)`.
/// Howard Hinnant's `civil_from_days` algorithm (public domain).
fn civil_from_days(z: i64) -> (i64, u32, u32) {
    let z = z + 719468;
    let era = if z >= 0 { z } else { z - 146096 } / 146097;
    let doe = (z - era * 146097) as u64;
    let yoe = (doe - doe / 1460 + doe / 36524 - doe / 146096) / 365;
    let y = yoe as i64 + era * 400;
    let doy = doe - (365 * yoe + yoe / 4 - yoe / 100);
    let mp = (5 * doy + 2) / 153;
    let d = (doy - (153 * mp + 2) / 5 + 1) as u32;
    let m = if mp < 10 { mp + 3 } else { mp - 9 } as u32;
    let y = if m <= 2 { y + 1 } else { y };
    (y, m, d)
}

/// `YYYY-MM-DDTHH:MM:SS`, UTC — same format as Python's
/// `datetime.now().isoformat(timespec='seconds')`, but in UTC rather than
/// local time (no timezone database is available without an extra
/// dependency, and — per the module doc comment — this line is excluded
/// from parity checks by design; only its *presence and format* matter).
fn now_iso_seconds() -> String {
    let secs = std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .map(|d| d.as_secs() as i64)
        .unwrap_or(0);
    let days = secs.div_euclid(86400);
    let secs_of_day = secs.rem_euclid(86400);
    let (y, mo, d) = civil_from_days(days);
    let h = secs_of_day / 3600;
    let mi = (secs_of_day % 3600) / 60;
    let se = secs_of_day % 60;
    format!("{y:04}-{mo:02}-{d:02}T{h:02}:{mi:02}:{se:02}")
}

/// Port of `render_header`.
fn render_header(file_name: &str, fmt: &str, sheet_names: &[String]) -> String {
    [
        "workbook:".to_string(),
        format!("  file: {file_name}"),
        format!("  format: {fmt}"),
        format!("  sheets[{}]: {}", sheet_names.len(), sheet_names.join(",")),
        format!("  extracted_at: {}", now_iso_seconds()),
    ]
    .join("\n")
}

/// Build the spec text `xlsb-extract-spec`/`xlsb_reader.extract_spec`
/// produce for one workbook. Port of `build_spec`.
///
/// `file_name` is the workbook's base file name (used in the `file:`
/// header line) and `format_ext` is its lowercase extension without the
/// leading dot (used in the `format:` header line and to gate the VBA
/// section — see `render.rs`'s `supports_vba` for the same
/// `hasattr(wb, "iter_vba_modules")` mirror).
pub fn extract_spec(
    data: &WorkbookData,
    file_name: &str,
    format_ext: &str,
    sample_rows: usize,
    sheets: &str,
) -> Result<String> {
    let selected: Option<HashSet<String>> = if sheets == "all" {
        None
    } else {
        Some(
            sheets
                .split(',')
                .map(str::trim)
                .filter(|s| !s.is_empty())
                .map(str::to_string)
                .collect(),
        )
    };
    let sheet_selected = |name: &str| -> bool {
        match &selected {
            None => true,
            Some(set) => set.contains(name),
        }
    };

    let sample_rows_u32 = sample_rows as u32;
    let scan_rows = sample_rows_u32.max(MIN_SCAN_ROWS);
    let supports_vba = format_ext == "xlsx" || format_ext == "xlsm";

    let sheet_names = data.sheet_names.clone();
    let mut header_maps: HashMap<String, HashMap<i64, String>> = HashMap::new();
    let mut sheet_dims: Vec<(String, (u32, u32))> = Vec::new();
    let mut schema_sections: Vec<String> = Vec::new();
    let mut sample_sections: Vec<String> = Vec::new();

    for (sheet_name, values) in &data.values {
        let (nrows, ncols) = sheet_dimensions(values);
        let (header_names, data_start_row) = detect_headers(values, ncols);
        let header_names_i64: HashMap<i64, String> = header_names
            .iter()
            .map(|(&k, v)| (k as i64, v.clone()))
            .collect();
        header_maps.insert(sheet_name.clone(), header_names_i64);
        sheet_dims.push((sheet_name.clone(), (nrows, ncols)));

        if !sheet_selected(sheet_name) {
            continue;
        }

        let columns = build_column_schema(
            values,
            ncols,
            &header_names,
            data_start_row,
            nrows,
            scan_rows,
        );
        schema_sections.push(render_sheet_schema(sheet_name, nrows, ncols, &columns));
        if sample_rows > 0 {
            sample_sections.push(render_sheet_sample(
                sheet_name,
                values,
                &header_names,
                data_start_row,
                nrows,
                ncols,
                sample_rows_u32,
            ));
        }
    }

    let (formulas_by_sheet, dependencies) = extract_formulas(data, &header_maps, &sheet_selected);
    let pivots = normalize_pivots(data, &sheet_selected);
    let filters = normalize_filters(data, &sheet_selected, &header_maps);
    let vba_modules: &[(String, String)] = if supports_vba { &data.vba } else { &[] };

    let mut sections: Vec<String> = vec![render_header(file_name, format_ext, &sheet_names)];
    sections.extend(schema_sections);

    let formulas_value: Value = Value::Object(
        formulas_by_sheet
            .iter()
            .map(|(k, cols)| (k.clone(), Value::Object(cols.iter().cloned().collect())))
            .collect(),
    );
    sections.push(fenced_json("formulas", &formulas_value));

    let dependencies_value: Value = Value::Object(
        dependencies
            .iter()
            .map(|(k, v)| {
                (
                    k.clone(),
                    Value::Array(v.iter().cloned().map(Value::String).collect()),
                )
            })
            .collect(),
    );
    sections.push(fenced_json("dependencies", &dependencies_value));

    sections.push(fenced_json("pivots", &Value::Array(pivots.clone())));
    sections.push(fenced_json("filters", &Value::Array(filters.clone())));

    for (name, src) in vba_modules {
        sections.push(fenced_vba(name, src));
    }

    sections.extend(sample_sections);
    sections.push(render_hints(
        &sheet_dims,
        &formulas_by_sheet,
        &header_maps,
        &dependencies,
        &pivots,
        &filters,
        vba_modules,
    ));

    Ok(format!("{}\n", sections.join("\n---\n")))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn refs(s: &str) -> Vec<(String, String)> {
        find_cell_refs(s)
            .iter()
            .map(|m| (m.col.clone(), m.row.clone()))
            .collect()
    }

    #[test]
    fn cell_ref_scanner_basic() {
        assert_eq!(
            refs("A1+B2"),
            vec![("A".into(), "1".into()), ("B".into(), "2".into())]
        );
    }

    #[test]
    fn cell_ref_scanner_matches_up_to_three_letter_columns() {
        // "TAG" is 3 uppercase letters immediately followed by digits, with
        // nothing before it to fail the lookbehind — this *is* a
        // syntactically valid match per the real Python regex (verified
        // directly against `re` — the pattern doesn't know or care about
        // Excel's actual column-count limit).
        assert_eq!(refs("TAG123"), vec![("TAG".into(), "123".into())]);
    }

    #[test]
    fn cell_ref_scanner_rejects_too_many_letters_or_digits() {
        // 4 consecutive uppercase letters before digits -> no valid split.
        assert_eq!(refs("ABCD123"), Vec::<(String, String)>::new());
        // 8 consecutive digits after valid letters -> lookahead always fails.
        assert_eq!(refs("A12345678"), Vec::<(String, String)>::new());
    }

    #[test]
    fn cell_ref_scanner_function_call_not_a_ref() {
        // Followed by '(' -> not a cell reference (it's a function/UDF call).
        assert_eq!(refs("A1(x)"), Vec::<(String, String)>::new());
    }

    #[test]
    fn cell_ref_scanner_bareword_sheet_prefix() {
        let m = &find_cell_refs("Sheet1!A1")[0];
        assert_eq!(m.sheet.as_deref(), Some("Sheet1"));
        assert_eq!((m.col.as_str(), m.row.as_str()), ("A", "1"));
    }

    #[test]
    fn cell_ref_scanner_quoted_sheet_prefix() {
        let m = &find_cell_refs("'My Sheet'!B2")[0];
        assert_eq!(m.qsheet.as_deref(), Some("My Sheet"));
        assert_eq!((m.col.as_str(), m.row.as_str()), ("B", "2"));
    }

    #[test]
    fn cell_ref_scanner_falls_back_when_prefix_has_no_valid_col_row() {
        // "Sheet1!" structurally looks like a prefix, but "XYZ" (no digits)
        // doesn't complete a col/row match — falls back to trying "S" (part
        // of "Sheet1") directly, which also fails, so nothing matches here.
        assert_eq!(refs("Sheet1!XYZ"), Vec::<(String, String)>::new());
    }

    #[test]
    fn cell_ref_scanner_dollar_signs_absolute_refs() {
        let m = &find_cell_refs("$A$1")[0];
        assert_eq!((m.col.as_str(), m.row.as_str()), ("A", "1"));
    }

    #[test]
    fn expand_normalize_formula_relative_offset() {
        let header_maps: HashMap<String, HashMap<i64, String>> = HashMap::new();
        let (pattern, refd) = normalize_formula("=A1+A2", "Sheet1", 4, &header_maps);
        assert!(refd.is_empty());
        // cur_row=4 (0-based) -> row1 (A1, 1-based row 1) offset = (1-1)-4 = -4;
        // row2 (A2) offset = (2-1)-4 = -3.
        assert_eq!(pattern, "=<A-4>+<A-3>");
    }

    #[test]
    fn normalize_formula_cross_sheet_reference() {
        let header_maps: HashMap<String, HashMap<i64, String>> = HashMap::new();
        let (pattern, refd) = normalize_formula("=Other!A1", "Sheet1", 0, &header_maps);
        assert_eq!(pattern, "=<Other!A>");
        assert!(refd.contains("Other"));
    }

    #[test]
    fn is_iso_date_validates_calendar() {
        assert!(is_iso_date("2024-01-15"));
        assert!(is_iso_date("2024-02-29")); // leap year
        assert!(!is_iso_date("2023-02-29")); // not a leap year
        assert!(!is_iso_date("2024-13-01")); // invalid month
        assert!(!is_iso_date("2024-01-32")); // invalid day
        assert!(!is_iso_date("not a date"));
        assert!(is_iso_date("2024-01-15T10:30:00")); // prefix match only
    }

    #[test]
    fn most_common_first_ties_prefer_earliest_inserted() {
        let counter = vec![("first".to_string(), 3), ("second".to_string(), 3)];
        assert_eq!(most_common_first(&counter), ("first", 3));
        let counter2 = vec![("first".to_string(), 1), ("second".to_string(), 5)];
        assert_eq!(most_common_first(&counter2), ("second", 5));
    }

    #[test]
    fn csv_field_quotes_when_needed() {
        assert_eq!(csv_field("plain"), "plain");
        assert_eq!(csv_field("a,b"), "\"a,b\"");
        assert_eq!(csv_field("has \"quote\""), "\"has \"\"quote\"\"\"");
        assert_eq!(csv_field("multi\nline"), "\"multi\nline\"");
    }

    #[test]
    fn toon_table_renders_header_and_rows() {
        let cols = vec!["a".to_string(), "b".to_string()];
        let rows = vec![vec!["1".to_string(), "2".to_string()]];
        assert_eq!(toon_table("t", &cols, &rows), "t[1]{a,b}:\n  1,2");
    }

    #[test]
    fn distinct_key_collapses_python_numeric_tower() {
        // 1 (int), 1.0 (float), and True (bool) are all == in Python and
        // would collapse into one element in a `set()`.
        let keys: HashSet<DistinctKey> = [
            CellValue::Int(1),
            CellValue::Float(1.0),
            CellValue::Bool(true),
        ]
        .iter()
        .map(distinct_key)
        .collect();
        assert_eq!(keys.len(), 1);
    }

    #[test]
    fn infer_column_type_precedence() {
        assert_eq!(
            infer_column_type(&[&CellValue::Int(1), &CellValue::Int(2)]),
            "integer"
        );
        assert_eq!(
            infer_column_type(&[&CellValue::Int(1), &CellValue::Float(2.5)]),
            "float"
        );
        assert_eq!(
            infer_column_type(&[&CellValue::Str("2024-01-01".to_string())]),
            "date"
        );
        assert_eq!(infer_column_type(&[&CellValue::Bool(true)]), "boolean");
        assert_eq!(
            infer_column_type(&[&CellValue::Str("hello".to_string())]),
            "string"
        );
        assert_eq!(infer_column_type(&[]), "string");
    }

    #[test]
    fn letters_to_col_matches_col_to_letter_roundtrip() {
        for col in [0u32, 1, 25, 26, 27, 701, 702] {
            let letters = common::col_to_letter(col);
            assert_eq!(letters_to_col(&letters), col as i64);
        }
    }
}
