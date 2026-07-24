//! common.rs — shared foundation for the `xlsb_reader_rs` crate.
//!
//! This module is the CONTRACT that every other module (`xlsb`, `xlsx`,
//! `vba`, `render`, `spec`) builds on. It provides:
//!
//!   - [`col_to_letter`] — 0-based column index -> Excel column letters.
//!   - [`CellValue`] — the small set of value types a worksheet cell can
//!     hold, mirroring exactly what the pure-Python `iter_values()`
//!     generators yield (`int`, `float`, `str`, `bool`).
//!   - [`WorkbookData`] — the parsed, backend-agnostic representation of a
//!     whole workbook. Parsers (Wave 1A `.xlsb`, Wave 1B `.xlsx`/`.xlsm`,
//!     Wave 1C VBA) fill this in; renderers (Wave 1D) and the PyO3
//!     `#[pyclass]`es consume it.
//!   - [`Reader`] — a little-endian byte-cursor helper for binary (BIFF12)
//!     parsing, porting the read primitives used throughout
//!     `xlsb_reader/_reader.py` (`_Decompiler._u8/_u16/_i32/_u32/_f64/_wstr16`
//!     and `RecordReader._varint`).
//!   - [`XlsbError`] / [`Result`] — the crate's error type, convertible to
//!     `pyo3::PyErr` so it can be returned directly from `#[pyfunction]`s
//!     and `#[pymethods]`.
//!   - PyO3 conversion helpers (`*_to_py*`) that turn the data model above
//!     into the *exact* Python objects the pure-Python `iter_*()` methods
//!     yield today, so the Rust backend is a drop-in replacement.
//!
//! # Design note: `FilterInfo` / `PivotTable` as `serde_json::Value`
//!
//! `iter_filters()` and `iter_pivot_tables()` in the pure-Python readers
//! return irregular, deeply-nested dicts whose exact shape differs subtly
//! between the `.xlsb` binary parser (`_parse_worksheet_filters`,
//! `_parse_pivot_table_part` in `_reader.py`) and the `.xlsx` XML parser
//! (`_parse_auto_filter`, `_parse_pivot_table_xml` in `_xlsx_reader.py`).
//! Modelling every field as a dedicated Rust struct would require the two
//! parser waves (1A/1B) to agree on a shared schema up front and would
//! make it easy to silently drop fields the Python side still emits.
//!
//! Instead, `FilterInfo` and `PivotTable` are both `serde_json::Value`:
//! parsers build them directly as JSON objects/arrays, and
//! [`json_value_to_py`] converts a `Value` into the equivalent Python
//! object (`dict`/`list`/`str`/`int`/`float`/`bool`/`None`). This keeps the
//! Rust model flexible and trivially JSON-parity-friendly (the `to_json`
//! one-shot function can serialize these sections directly with
//! `serde_json::to_string`), at the cost of losing compile-time field
//! checking for those two sections specifically.

use std::collections::BTreeMap;
use std::fmt;

use pyo3::exceptions::{PyIOError, PyValueError};
use pyo3::prelude::*;
use pyo3::types::{PyDict, PyList, PyTuple};

// ---------------------------------------------------------------------------
// col_to_letter
// ---------------------------------------------------------------------------

/// Convert a 0-based column index to Excel column letter(s), e.g. `0 -> "A"`,
/// `26 -> "AA"`.
///
/// Direct port of `xlsb_reader/_reader.py::col_to_letter`:
///
/// ```python
/// def col_to_letter(col: int) -> str:
///     result = ""
///     col += 1
///     while col:
///         col, rem = divmod(col - 1, 26)
///         result = chr(65 + rem) + result
///     return result
/// ```
pub fn col_to_letter(col: u32) -> String {
    let mut result = String::new();
    let mut n: i64 = col as i64 + 1;
    while n > 0 {
        let rem = (n - 1) % 26;
        n = (n - 1) / 26;
        result.insert(0, (b'A' + rem as u8) as char);
    }
    result
}

// ---------------------------------------------------------------------------
// CellValue
// ---------------------------------------------------------------------------

/// A worksheet cell's cached/constant value.
///
/// Mirrors the small set of Python types `iter_values()` yields: `int`,
/// `float`, `str`, and `bool` (see `_parse_worksheet_values` in
/// `_reader.py` / `_xlsx_reader.py` — RK numbers and whole `REAL`s become
/// `int`, everything else numeric becomes `float`, `BOOL` cells become
/// `bool`, strings/errors become `str`).
#[derive(Debug, Clone, PartialEq)]
pub enum CellValue {
    Int(i64),
    Float(f64),
    Str(String),
    Bool(bool),
}

/// A single sheet's `{(row, col): value}` map (0-based keys), for whichever
/// per-cell payload `T` a parser produces. Named alias purely to keep
/// `WorkbookData` and the conversion helpers below under clippy's
/// `type_complexity` threshold — structurally identical to writing out
/// `BTreeMap<(u32, u32), T>` everywhere.
pub type CellMap<T> = BTreeMap<(u32, u32), T>;

impl CellValue {
    /// Convert to the exact Python object `iter_values()` would yield for
    /// this cell (`int`/`float`/`str`/`bool`).
    pub fn to_py(&self, py: Python<'_>) -> PyResult<Py<PyAny>> {
        Ok(match self {
            // NOTE: match Bool before Int is irrelevant here since CellValue
            // is an explicit enum (no accidental bool/int aliasing like in
            // dynamically-typed Python) — each variant maps 1:1.
            CellValue::Int(v) => v.into_pyobject(py)?.into_any().unbind(),
            CellValue::Float(v) => v.into_pyobject(py)?.into_any().unbind(),
            CellValue::Str(v) => v.into_pyobject(py)?.into_any().unbind(),
            CellValue::Bool(v) => v.into_pyobject(py)?.to_owned().into_any().unbind(),
        })
    }
}

// ---------------------------------------------------------------------------
// FilterInfo / PivotTable — see module-level design note above.
// ---------------------------------------------------------------------------

/// Auto-filter metadata for one sheet's `<AutoFilter>`/`BrtBeginAFilter`
/// block. See `_parse_worksheet_filters` (`_reader.py`) and
/// `_parse_auto_filter` (`_xlsx_reader.py`) for the shapes this must match.
pub type FilterInfo = serde_json::Value;

/// PivotTable metadata for one pivot table. See `_parse_pivot_table_part`
/// (`_reader.py`) and `_parse_pivot_table_xml` (`_xlsx_reader.py`) for the
/// shapes this must match.
pub type PivotTable = serde_json::Value;

// ---------------------------------------------------------------------------
// WorkbookData — the shared parse result
// ---------------------------------------------------------------------------

/// Backend-agnostic parsed representation of a whole workbook.
///
/// Wave 1A (`xlsb.rs`, `.xlsb` binary format) and Wave 1B (`xlsx.rs`,
/// `.xlsx`/`.xlsm` OOXML format) are each responsible for producing one of
/// these from a file path. Wave 1C (`vba.rs`) fills in `vba`. Wave 1D
/// (`render.rs`, `spec.rs`) and the `#[pyclass]` wrappers in `xlsb.rs`/
/// `xlsx.rs` consume it to answer `iter_formulas()`, `iter_values()`,
/// `iter_filters()`, `iter_pivot_tables()`, `iter_vba_modules()`, and the
/// one-shot `to_dict`/`to_json`/`to_markdown`/`extract_spec` functions.
///
/// Field shapes deliberately mirror the pure-Python `iter_*()` generators
/// exactly (0-based `(row, col)` keys, per-sheet ordering preserved as a
/// `Vec` rather than a `HashMap` so sheet order == workbook tab order):
///
///   - `formulas` <-> `XlsbWorkbook.iter_formulas()` /
///     `XlsxWorkbook.iter_formulas()`: `Iterator[Tuple[str, Dict[(int,int), str]]]`
///   - `values` <-> `iter_values()`: `Iterator[Tuple[str, Dict[(int,int), object]]]`
///   - `filters` <-> `XlsbWorkbook.iter_filters()`:
///     `Iterator[Tuple[str, Optional[Dict]]]` (the `.xlsb` reader yields a
///     `(sheet_name, filter_info_or_None)` pair for *every* sheet; the
///     `.xlsx` reader's `iter_filters()` only yields sheets that actually
///     have an AutoFilter — Wave 1B should populate `filters` with an entry
///     for every sheet, `None` where absent, and have the xlsx pyclass
///     filter out the `None`s when converting to match that asymmetry, see
///     `xlsx.rs`).
///   - `pivots` <-> `iter_pivot_tables()`: `Iterator[Dict[str, object]]`
///     (flat list, each `Value` already carries its own `"sheet"` key).
///   - `vba` <-> `iter_vba_modules()`: `Dict[str, str]` (module name -> VBA
///     source).
#[derive(Debug, Clone, Default)]
pub struct WorkbookData {
    pub sheet_names: Vec<String>,
    pub formulas: Vec<(String, CellMap<String>)>,
    pub values: Vec<(String, CellMap<CellValue>)>,
    pub filters: Vec<(String, Option<FilterInfo>)>,
    pub pivots: Vec<PivotTable>,
    /// `(module_name, source)` pairs in the VBA project's own `dir`-stream
    /// declaration order — a `Vec` rather than a `BTreeMap` specifically so
    /// this order (which Python's `iter_vba_modules()` dict preserves) isn't
    /// silently alphabetized. See `vba::read_vba_modules`.
    pub vba: Vec<(String, String)>,
}

// ---------------------------------------------------------------------------
// Errors
// ---------------------------------------------------------------------------

/// Crate-wide error type. Converts into `pyo3::PyErr` via `From` so it can
/// be returned directly as the `Err` variant of any `PyResult<T>`-returning
/// `#[pyfunction]`/`#[pymethods]` body (including through the `?` operator).
#[derive(Debug)]
pub enum XlsbError {
    /// Ran out of input while reading a fixed-size or length-prefixed field.
    Eof {
        what: &'static str,
        needed: usize,
        available: usize,
    },
    /// Generic malformed-data condition with a human-readable description
    /// (mirrors the pure-Python readers' habit of returning partial/empty
    /// results rather than raising, but Wave 1 parsers may choose to
    /// surface hard failures this way where the Python side would have
    /// silently produced garbage).
    Parse(String),
    Io(std::io::Error),
    Zip(zip::result::ZipError),
    Xml(quick_xml::Error),
}

impl fmt::Display for XlsbError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            XlsbError::Eof {
                what,
                needed,
                available,
            } => write!(
                f,
                "unexpected end of data reading {what}: needed {needed} byte(s), \
                 {available} available"
            ),
            XlsbError::Parse(msg) => write!(f, "parse error: {msg}"),
            XlsbError::Io(e) => write!(f, "I/O error: {e}"),
            XlsbError::Zip(e) => write!(f, "zip error: {e}"),
            XlsbError::Xml(e) => write!(f, "XML error: {e}"),
        }
    }
}

impl std::error::Error for XlsbError {}

impl From<std::io::Error> for XlsbError {
    fn from(e: std::io::Error) -> Self {
        XlsbError::Io(e)
    }
}

impl From<zip::result::ZipError> for XlsbError {
    fn from(e: zip::result::ZipError) -> Self {
        XlsbError::Zip(e)
    }
}

impl From<quick_xml::Error> for XlsbError {
    fn from(e: quick_xml::Error) -> Self {
        XlsbError::Xml(e)
    }
}

impl From<XlsbError> for PyErr {
    fn from(err: XlsbError) -> PyErr {
        match &err {
            XlsbError::Io(_) | XlsbError::Zip(_) => PyIOError::new_err(err.to_string()),
            XlsbError::Eof { .. } | XlsbError::Parse(_) | XlsbError::Xml(_) => {
                PyValueError::new_err(err.to_string())
            }
        }
    }
}

/// Crate-wide result alias.
pub type Result<T> = std::result::Result<T, XlsbError>;

// ---------------------------------------------------------------------------
// Reader — little-endian byte-cursor helper for BIFF12 / binary parsing
// ---------------------------------------------------------------------------

/// A minimal little-endian byte-cursor, matching the read primitives used
/// throughout `_reader.py`:
///
///   - `Reader::u8/u16/u32/i32/f64`  <-> `_Decompiler._u8/_u16/_u32/_i32/_f64`
///   - `Reader::wstr16`              <-> `_Decompiler._wstr16`
///     (`u16` character count, then `cch * 2` bytes of UTF-16LE)
///   - `Reader::varint`              <-> `RecordReader._varint`
///     (BIFF12 record type/size varint: up to 4 bytes, 7 data bits each,
///     high bit = "more bytes follow")
///   - `Reader::take_bytes`/`remaining` <-> raw slicing done ad-hoc via
///     `io.BytesIO` in the Python parsers.
///
/// Every read method returns [`Result`] instead of panicking on truncated
/// input, since untrusted workbook bytes are the norm here.
pub struct Reader<'a> {
    buf: &'a [u8],
    pos: usize,
}

impl<'a> Reader<'a> {
    pub fn new(buf: &'a [u8]) -> Self {
        Reader { buf, pos: 0 }
    }

    /// Current read offset from the start of the buffer.
    ///
    /// Not currently called by any parser (each wrote its own equivalent
    /// bookkeeping inline), but kept as part of this documented byte-cursor
    /// contract alongside the other primitives.
    #[allow(dead_code)]
    pub fn pos(&self) -> usize {
        self.pos
    }

    /// Number of unread bytes remaining.
    pub fn remaining(&self) -> usize {
        self.buf.len() - self.pos
    }

    /// True once every byte has been consumed.
    pub fn is_empty(&self) -> bool {
        self.pos >= self.buf.len()
    }

    /// Take and return the next `n` bytes without copying, advancing the
    /// cursor. Errors if fewer than `n` bytes remain.
    pub fn take_bytes(&mut self, n: usize) -> Result<&'a [u8]> {
        if self.remaining() < n {
            return Err(XlsbError::Eof {
                what: "take_bytes",
                needed: n,
                available: self.remaining(),
            });
        }
        let out = &self.buf[self.pos..self.pos + n];
        self.pos += n;
        Ok(out)
    }

    /// Skip `n` bytes without returning them (mirrors the many
    /// `buf.read(n)` calls in the Python parsers used purely to advance
    /// past reserved/ignored fields).
    pub fn skip(&mut self, n: usize) -> Result<()> {
        self.take_bytes(n).map(|_| ())
    }

    pub fn u8(&mut self) -> Result<u8> {
        Ok(self.take_bytes(1)?[0])
    }

    pub fn u16(&mut self) -> Result<u16> {
        let b = self.take_bytes(2)?;
        Ok(u16::from_le_bytes([b[0], b[1]]))
    }

    pub fn u32(&mut self) -> Result<u32> {
        let b = self.take_bytes(4)?;
        Ok(u32::from_le_bytes([b[0], b[1], b[2], b[3]]))
    }

    pub fn i32(&mut self) -> Result<i32> {
        let b = self.take_bytes(4)?;
        Ok(i32::from_le_bytes([b[0], b[1], b[2], b[3]]))
    }

    pub fn f64(&mut self) -> Result<f64> {
        let b = self.take_bytes(8)?;
        Ok(f64::from_le_bytes([
            b[0], b[1], b[2], b[3], b[4], b[5], b[6], b[7],
        ]))
    }

    /// Read a `u16`-length-prefixed UTF-16LE string: `cch: u16` followed by
    /// `cch * 2` bytes. Matches `_Decompiler._wstr16` (used for `PtgStr`
    /// operands, i.e. formula string literals).
    pub fn wstr16(&mut self) -> Result<String> {
        let cch = self.u16()? as usize;
        let bytes = self.take_bytes(cch * 2)?;
        Ok(utf16le_lossy(bytes))
    }

    /// Read a `u32`-length-prefixed UTF-16LE string: `cch: u32` followed by
    /// `cch * 2` bytes. Matches the `XLWideString` pattern used for cell
    /// text/SST entries/etc. (`_read_xlwide_from` / inline reads in
    /// `_read_sst`, `_read_workbook`, `_read_defined_names`).
    ///
    /// Not currently called: `xlsb.rs`'s own `read_xlwide_from` inlines the
    /// same read against its own `RecordReader`-scoped buffers instead of
    /// going through this shared `Reader`. Kept as part of the documented
    /// contract (see module doc comment) for API completeness.
    #[allow(dead_code)]
    pub fn xlwstr32(&mut self) -> Result<String> {
        let cch = self.u32()? as usize;
        let bytes = self.take_bytes(cch * 2)?;
        Ok(utf16le_lossy(bytes))
    }

    /// Read a BIFF12 varint: up to 4 bytes, 7 payload bits per byte
    /// (little-endian bit order), continuation signalled by the top bit.
    /// Matches `RecordReader._varint` exactly, including its 4-byte cap.
    ///
    /// Not currently called: `xlsb.rs`'s own `RecordReader` has an
    /// equivalent private method operating on its own cursor state instead
    /// of going through this shared `Reader`. Kept as part of the
    /// documented contract for API completeness.
    #[allow(dead_code)]
    pub fn varint(&mut self) -> Result<u32> {
        let mut result: u32 = 0;
        let mut shift: u32 = 0;
        for _ in 0..4 {
            let b = self.u8()?;
            result |= ((b & 0x7F) as u32) << shift;
            shift += 7;
            if b & 0x80 == 0 {
                break;
            }
        }
        Ok(result)
    }
}

/// Decode UTF-16LE bytes, replacing unpaired/invalid code units with U+FFFD
/// — matches Python's `.decode("utf-16-le", errors="replace")`.
fn utf16le_lossy(bytes: &[u8]) -> String {
    let units: Vec<u16> = bytes
        .chunks_exact(2)
        .map(|c| u16::from_le_bytes([c[0], c[1]]))
        .collect();
    char::decode_utf16(units)
        .map(|r| r.unwrap_or('\u{FFFD}'))
        .collect()
}

// ---------------------------------------------------------------------------
// PyO3 conversion helpers
// ---------------------------------------------------------------------------
//
// These turn `WorkbookData` (or slices of it) into the *exact* Python
// objects the pure-Python `iter_*()` generators yield, so callers on the
// Python side cannot tell which backend produced them.

/// Convert a `serde_json::Value` into the equivalent Python object:
/// `Null -> None`, `Bool -> bool`, integral `Number -> int`, other
/// `Number -> float`, `String -> str`, `Array -> list`, `Object -> dict`
/// (with `str` keys, insertion order preserved).
pub fn json_value_to_py(py: Python<'_>, value: &serde_json::Value) -> PyResult<Py<PyAny>> {
    use serde_json::Value;
    Ok(match value {
        Value::Null => py.None(),
        Value::Bool(b) => b.into_pyobject(py)?.to_owned().into_any().unbind(),
        Value::Number(n) => {
            if let Some(i) = n.as_i64() {
                i.into_pyobject(py)?.into_any().unbind()
            } else if let Some(u) = n.as_u64() {
                u.into_pyobject(py)?.into_any().unbind()
            } else {
                n.as_f64()
                    .unwrap_or(0.0)
                    .into_pyobject(py)?
                    .into_any()
                    .unbind()
            }
        }
        Value::String(s) => s.into_pyobject(py)?.into_any().unbind(),
        Value::Array(items) => {
            let list = PyList::empty(py);
            for item in items {
                list.append(json_value_to_py(py, item)?)?;
            }
            list.into_any().unbind()
        }
        Value::Object(map) => {
            let dict = PyDict::new(py);
            for (k, v) in map {
                dict.set_item(k, json_value_to_py(py, v)?)?;
            }
            dict.into_any().unbind()
        }
    })
}

/// Build a Python `dict` from a `{(row, col): formula}` map, with `(row,
/// col)` as a Python `(int, int)` tuple key — matches every pure-Python
/// `iter_formulas()`/`iter_values()` generator's per-sheet payload.
/// `BTreeMap` iteration order (row-major, then column) is preserved, same
/// as the `sorted(formulas.items())` the CLI/render layer applies before
/// display (though callers should not rely on dict ordering for
/// correctness — only content — same as today).
pub fn cell_formula_map_to_py(
    py: Python<'_>,
    map: &CellMap<String>,
) -> PyResult<Py<PyDict>> {
    let dict = PyDict::new(py);
    for (&(row, col), formula) in map.iter() {
        let key = PyTuple::new(py, [row, col])?;
        dict.set_item(key, formula)?;
    }
    Ok(dict.into())
}

/// Same as [`cell_formula_map_to_py`] but for cached/constant cell values.
pub fn cell_value_map_to_py(
    py: Python<'_>,
    map: &CellMap<CellValue>,
) -> PyResult<Py<PyDict>> {
    let dict = PyDict::new(py);
    for (&(row, col), value) in map.iter() {
        let key = PyTuple::new(py, [row, col])?;
        dict.set_item(key, value.to_py(py)?)?;
    }
    Ok(dict.into())
}

/// Build the list `XlsbWorkbook.iter_formulas()`/`XlsxWorkbook.iter_formulas()`
/// yield: `[(sheet_name, {(row, col): formula, ...}), ...]`.
pub fn formulas_to_py(
    py: Python<'_>,
    formulas: &[(String, CellMap<String>)],
) -> PyResult<Py<PyList>> {
    let list = PyList::empty(py);
    for (sheet, map) in formulas {
        let dict = cell_formula_map_to_py(py, map)?;
        let tuple = PyTuple::new(py, [sheet.into_pyobject(py)?.into_any(), dict.into_bound(py).into_any()])?;
        list.append(tuple)?;
    }
    Ok(list.into())
}

/// Build the list `iter_values()` yields:
/// `[(sheet_name, {(row, col): value, ...}), ...]`.
pub fn values_to_py(
    py: Python<'_>,
    values: &[(String, CellMap<CellValue>)],
) -> PyResult<Py<PyList>> {
    let list = PyList::empty(py);
    for (sheet, map) in values {
        let dict = cell_value_map_to_py(py, map)?;
        let tuple = PyTuple::new(py, [sheet.into_pyobject(py)?.into_any(), dict.into_bound(py).into_any()])?;
        list.append(tuple)?;
    }
    Ok(list.into())
}

/// Build the list `XlsbWorkbook.iter_filters()` yields:
/// `[(sheet_name, filter_info_or_None), ...]`, one entry per sheet.
///
/// Note: `XlsxWorkbook.iter_filters()` instead yields *plain filter dicts*
/// (each already containing a `"sheet"` key) only for sheets that have an
/// AutoFilter — see the doc comment on [`WorkbookData::filters`] for how
/// `xlsx.rs`'s pyclass should adapt this same underlying data to that
/// shape.
pub fn filters_to_py(
    py: Python<'_>,
    filters: &[(String, Option<FilterInfo>)],
) -> PyResult<Py<PyList>> {
    let list = PyList::empty(py);
    for (sheet, info) in filters {
        let info_obj = match info {
            Some(v) => json_value_to_py(py, v)?,
            None => py.None(),
        };
        let tuple = PyTuple::new(py, [sheet.into_pyobject(py)?.into_any(), info_obj.into_bound(py)])?;
        list.append(tuple)?;
    }
    Ok(list.into())
}

/// Build the flat list `iter_pivot_tables()` yields: `[pivot_dict, ...]`.
pub fn pivots_to_py(py: Python<'_>, pivots: &[PivotTable]) -> PyResult<Py<PyList>> {
    let list = PyList::empty(py);
    for pt in pivots {
        list.append(json_value_to_py(py, pt)?)?;
    }
    Ok(list.into())
}

/// Build the dict `iter_vba_modules()` returns: `{module_name: source}`.
pub fn vba_to_py(py: Python<'_>, vba: &[(String, String)]) -> PyResult<Py<PyDict>> {
    let dict = PyDict::new(py);
    for (name, src) in vba {
        dict.set_item(name, src)?;
    }
    Ok(dict.into())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn col_to_letter_matches_python() {
        assert_eq!(col_to_letter(0), "A");
        assert_eq!(col_to_letter(1), "B");
        assert_eq!(col_to_letter(25), "Z");
        assert_eq!(col_to_letter(26), "AA");
        assert_eq!(col_to_letter(27), "AB");
        assert_eq!(col_to_letter(51), "AZ");
        assert_eq!(col_to_letter(52), "BA");
        assert_eq!(col_to_letter(701), "ZZ");
        assert_eq!(col_to_letter(702), "AAA");
    }

    #[test]
    fn reader_primitives() {
        let data = [0x01, 0x02, 0x03, 0x04, 0x00, 0x00, 0x00, 0x00];
        let mut r = Reader::new(&data);
        assert_eq!(r.u8().unwrap(), 0x01);
        assert_eq!(r.remaining(), 7);
        let mut r2 = Reader::new(&data);
        assert_eq!(r2.u32().unwrap(), 0x04030201);
    }

    #[test]
    fn varint_single_byte() {
        let data = [0x7F];
        let mut r = Reader::new(&data);
        assert_eq!(r.varint().unwrap(), 0x7F);
    }

    #[test]
    fn varint_two_bytes() {
        // 0x80 | 0x01, 0x01 -> (0x01) | (0x01 << 7) = 0x81
        let data = [0x81, 0x01];
        let mut r = Reader::new(&data);
        assert_eq!(r.varint().unwrap(), 0x81);
    }

    #[test]
    fn eof_reports_short_read() {
        let data = [0x01];
        let mut r = Reader::new(&data);
        assert!(r.u32().is_err());
    }
}
