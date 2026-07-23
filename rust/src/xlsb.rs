//! xlsb.rs — `.xlsb` (Excel Binary Workbook) support.
//!
//! TODO(Wave 1A): Port `xlsb_reader/_reader.py` in full. In particular:
//!   - `RecordReader` (BIFF12 variable-length record iterator: `[type varint]
//!     [size varint] [data]`) — build this on top of `common::Reader::varint`.
//!   - `_Decompiler` (stack-based RPN Ptg token decompiler -> infix formula
//!     strings) — use `common::Reader` for all token field reads.
//!   - `_read_sst`, `_read_workbook`, `_read_defined_names`, `_read_rels`,
//!     `_resolve_rel_target` (OPC container plumbing).
//!   - `_parse_worksheet` / `_parse_worksheet_values` (per-sheet formula and
//!     value extraction, including shared/array formula resolution).
//!   - `_parse_worksheet_filters` (AutoFilter -> [`common::FilterInfo`] JSON
//!     matching the shape documented in `_parse_worksheet_filters`'s
//!     docstring).
//!   - `_parse_pivot_table_part` / `_parse_pivot_cache_fields` (PivotTable ->
//!     [`common::PivotTable`] JSON).
//!
//! The `.xlsb` container itself is a plain ZIP (see `zipfile.ZipFile` in
//! Python) — use the `zip` crate to open it and read parts by name, same
//! part paths as Python (`xl/workbook.bin`, `xl/worksheets/sheetN.bin`,
//! `xl/sharedStrings.bin`, `xl/_rels/workbook.bin.rels`, etc).
//!
//! This file currently only wires up the `XlsbWorkbook` PyO3 class with
//! stub parsing (always returns an empty [`WorkbookData`]) so the crate
//! compiles and the Python-side backend dispatch in Wave 0 can be
//! exercised end-to-end.

use std::path::{Path, PathBuf};
use std::sync::Mutex;

use pyo3::prelude::*;
use pyo3::types::{PyDict, PyList, PyTuple};

use crate::common::{self, Result, WorkbookData};

/// Parse a `.xlsb` file at `path` into a [`WorkbookData`].
///
/// TODO(Wave 1A): implement — see module-level doc comment above.
pub fn parse_xlsb(_path: &Path) -> Result<WorkbookData> {
    // STUB: Wave 1A replaces this with the real BIFF12 parser.
    Ok(WorkbookData::default())
}

/// Mirrors `xlsb_reader._reader.XlsbWorkbook`.
///
/// Holds the source file path and lazily parses it into a [`WorkbookData`]
/// on first use (cached for the lifetime of the object — note this is a
/// deliberate behavioural improvement over the pure-Python reader, which
/// re-parses from scratch on every `iter_*()` call; Wave 1A should preserve
/// this caching).
#[pyclass(module = "xlsb_reader_rs")]
pub struct XlsbWorkbook {
    path: PathBuf,
    data: Mutex<Option<WorkbookData>>,
}

impl XlsbWorkbook {
    fn ensure_parsed(&self) -> Result<()> {
        let mut guard = self.data.lock().expect("XlsbWorkbook mutex poisoned");
        if guard.is_none() {
            *guard = Some(parse_xlsb(&self.path)?);
        }
        Ok(())
    }
}

#[pymethods]
impl XlsbWorkbook {
    /// `path` accepts anything `os.PathLike` (`str` or `pathlib.Path`),
    /// same as `XlsbWorkbook.__init__(self, path: "os.PathLike")` in
    /// `_reader.py` — PyO3's `PathBuf` extraction handles both via
    /// `os.fspath()`.
    #[new]
    fn new(path: PathBuf) -> Self {
        XlsbWorkbook {
            path,
            data: Mutex::new(None),
        }
    }

    /// Ordered list of worksheet names. Matches
    /// `XlsbWorkbook.sheet_names` (a `@property` in Python).
    #[getter]
    fn sheet_names(&self) -> PyResult<Vec<String>> {
        self.ensure_parsed()?;
        let guard = self.data.lock().expect("XlsbWorkbook mutex poisoned");
        Ok(guard.as_ref().expect("just parsed").sheet_names.clone())
    }

    /// `Iterator[Tuple[str, Dict[(int,int), str]]]` — see
    /// `XlsbWorkbook.iter_formulas` in `_reader.py`.
    fn iter_formulas(&self, py: Python<'_>) -> PyResult<Py<PyList>> {
        self.ensure_parsed()?;
        let guard = self.data.lock().expect("XlsbWorkbook mutex poisoned");
        common::formulas_to_py(py, &guard.as_ref().expect("just parsed").formulas)
    }

    /// `Iterator[Tuple[str, Dict[(int,int), object]]]` — see
    /// `XlsbWorkbook.iter_values` in `_reader.py`.
    fn iter_values(&self, py: Python<'_>) -> PyResult<Py<PyList>> {
        self.ensure_parsed()?;
        let guard = self.data.lock().expect("XlsbWorkbook mutex poisoned");
        common::values_to_py(py, &guard.as_ref().expect("just parsed").values)
    }

    /// `Iterator[Tuple[str, Optional[Dict[str, object]]]]` — see
    /// `XlsbWorkbook.iter_filters` in `_reader.py` (yields one entry per
    /// sheet, `None` where there is no AutoFilter).
    fn iter_filters(&self, py: Python<'_>) -> PyResult<Py<PyList>> {
        self.ensure_parsed()?;
        let guard = self.data.lock().expect("XlsbWorkbook mutex poisoned");
        common::filters_to_py(py, &guard.as_ref().expect("just parsed").filters)
    }

    /// `Iterator[Dict[str, object]]` — see `XlsbWorkbook.iter_pivot_tables`
    /// in `_reader.py`.
    fn iter_pivot_tables(&self, py: Python<'_>) -> PyResult<Py<PyList>> {
        self.ensure_parsed()?;
        let guard = self.data.lock().expect("XlsbWorkbook mutex poisoned");
        common::pivots_to_py(py, &guard.as_ref().expect("just parsed").pivots)
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
