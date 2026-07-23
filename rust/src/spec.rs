//! spec.rs — the token-efficient structural spec extractor (`xlsb-extract-spec`).
//!
//! TODO(Wave 1D): Port `xlsb_reader/_spec_extractor.py` in full. This is the
//! most involved port in the crate — it's a full pipeline, not just a
//! parser:
//!   - Schema inference per sheet: `infer_column_type`, `detect_headers`,
//!     `build_column_schema`, `sheet_dimensions`.
//!   - Formula pattern extraction + inter-sheet dependency detection:
//!     `normalize_formula` (cell-reference regex substitution — see
//!     `CELL_REF_RE`, `ISO_DATE_RE`; the `regex` crate is a dependency of
//!     this crate specifically for this), `extract_formulas`.
//!   - PivotTable / AutoFilter normalization into one common
//!     representation across the `.xlsb` binary and `.xlsx` XML shapes:
//!     `normalize_pivots`, `normalize_filters`.
//!   - Rendering: `toon_table` (compact CSV-like tables), `fenced_json`,
//!     `fenced_vba`, `render_sheet_schema`, `render_sheet_sample`,
//!     `render_header`, `render_hints`.
//!   - Orchestration + final assembly: `build_spec` (section ordering,
//!     `"\n---\n".join(...)`).
//!
//! Also note `validate_spec` (the `--validate` CLI mode) is *not* part of
//! the one-shot `extract_spec` Python function `xlsb_reader.extract_spec`
//! dispatches to (see `__init__.py`) — it is CLI-only and out of scope for
//! this crate for now.
//!
//! Backs the crate's `extract_spec` `#[pyfunction]` in `lib.rs` (the "fast
//! path" `xlsb_reader.extract_spec` in `__init__.py` uses when this
//! extension is installed). Output MUST match
//! `xlsb_reader._spec_extractor.build_spec` byte-for-byte for the same
//! `sample_rows`/`sheets` arguments (modulo the `extracted_at` timestamp
//! line, which is inherently non-deterministic in both implementations).

use crate::common::{Result, WorkbookData};

/// Build the spec text `xlsb-extract-spec`/`xlsb_reader.extract_spec`
/// produce for one workbook.
///
/// `file_name` is the workbook's base file name (used in the `file:`
/// header line) and `format_ext` is its lowercase extension without the
/// leading dot (used in the `format:` header line) — see
/// `render_header`/`build_spec` in `_spec_extractor.py`.
///
/// TODO(Wave 1D): implement for real — see module-level doc comment above.
/// Current stub always returns an empty string.
pub fn extract_spec(
    _data: &WorkbookData,
    _file_name: &str,
    _format_ext: &str,
    _sample_rows: usize,
    _sheets: &str,
) -> Result<String> {
    // STUB: Wave 1D replaces this with the real spec builder.
    Ok(String::new())
}
