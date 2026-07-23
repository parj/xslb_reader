//! xlsx.rs — `.xlsx`/`.xlsm` (Excel Open XML) support.
//!
//! TODO(Wave 1B): Port `xlsb_reader/_xlsx_reader.py` in full. In particular:
//!   - `_parse_rels` / `_resolve_rel_target` (OPC relationship plumbing —
//!     use `quick-xml` instead of `xml.etree.ElementTree`).
//!   - `_parse_shared_strings` (`xl/sharedStrings.xml` -> `Vec<String>`,
//!     handling both simple `<t>` and rich-text `<r><t>` runs).
//!   - `_parse_workbook` (`xl/workbook.xml` -> sheet list + defined names).
//!   - `_parse_worksheet_formulas` (normal/shared/array formulas, including
//!     `_expand_shared_formula`'s cell-reference-shifting regex — the
//!     `regex` crate covers this).
//!   - `_parse_worksheet_values` (typed cell values: `s`/`b`/`e`/`str`/
//!     inline-string/numeric).
//!   - `_parse_pivot_table_xml` + `_parse_pivot_cache_fields_xml` ->
//!     [`common::PivotTable`] JSON matching the same shape as the `.xlsb`
//!     side (see `_spec_extractor.normalize_pivots`, which already treats
//!     both shapes uniformly — match that superset shape).
//!   - `_parse_auto_filter` -> [`common::FilterInfo`] JSON. Note this is
//!     merged with the sheet name into one object per the asymmetry
//!     documented on [`crate::common::WorkbookData::filters`].
//!   - `iter_vba_modules` — delegates to `vba::read_vba_modules` (Wave 1C)
//!     on the embedded `xl/vbaProject.bin` part, if present.
//!
//! The `.xlsx`/`.xlsm` container is a plain ZIP of XML parts — use the
//! `zip` crate to open it and `quick-xml` to parse each part, same part
//! paths as Python.
//!
//! This file currently only wires up the `XlsxWorkbook` PyO3 class with
//! stub parsing (always returns an empty [`WorkbookData`]) so the crate
//! compiles and the Python-side backend dispatch in Wave 0 can be
//! exercised end-to-end.

use std::path::{Path, PathBuf};
use std::sync::Mutex;

use pyo3::prelude::*;
use pyo3::types::{PyDict, PyList, PyTuple};

use crate::common::{self, Result, WorkbookData};

/// Parse a `.xlsx`/`.xlsm` file at `path` into a [`WorkbookData`].
///
/// TODO(Wave 1B): implement — see module-level doc comment above. Should
/// also populate `WorkbookData::vba` via [`crate::vba::read_vba_modules`]
/// when an `xl/vbaProject.bin` part is present (mirrors
/// `XlsxWorkbook.iter_vba_modules`'s lazy, on-demand read in Python — Wave
/// 1B may choose to parse VBA eagerly here instead, since `WorkbookData` is
/// cached as a whole).
pub fn parse_xlsx(_path: &Path) -> Result<WorkbookData> {
    // STUB: Wave 1B replaces this with the real OOXML parser.
    Ok(WorkbookData::default())
}

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
