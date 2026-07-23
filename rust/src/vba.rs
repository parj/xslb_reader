//! vba.rs — VBA macro extraction from embedded OLE (MS-CFB) containers.
//!
//! TODO(Wave 1C): Port `xlsb_reader/_vba_reader.py` in full:
//!   - MS-CFB compound-file reading (`_cfb_read_streams`) — prefer the
//!     `cfb` crate's reader over hand-rolling FAT/MiniFAT walking, but
//!     double check it exposes stream paths the same way (uppercase,
//!     `/`-joined, e.g. `"VBA/DIR"`, `"VBA/MODULE1"`, `"PROJECT"`) or adapt
//!     the lookup accordingly.
//!   - the MS-OVBA RLE decompressor (`_decompress`, §2.4.1.3.1) — this is
//!     pure bit-twiddling, safe to port near-verbatim.
//!   - `dir` stream parsing (§2.3.4.2) to recover the module name -> stream
//!     offset table.
//!   - module stream extraction (§2.3.4.3): skip the compiled P-code
//!     performance cache up to `text_offset`, decompress the remainder.
//!
//! Consumed by `xlsx::parse_xlsx` (Wave 1B) for `.xlsx`/`.xlsm` workbooks
//! that embed an `xl/vbaProject.bin` part (mirrors
//! `XlsxWorkbook.iter_vba_modules` in `_xlsx_reader.py`, which imports and
//! calls `xlsb_reader._vba_reader.read_vba_modules` lazily).
//!
//! Note: `.xlsb` workbooks can carry the same `xl/vbaProject.bin` part, but
//! `XlsbWorkbook` in `_reader.py` does not currently expose an
//! `iter_vba_modules` method. If a later wave adds that for parity, this
//! function is what would back it too.

use std::collections::BTreeMap;

use crate::common::Result;

/// Extract VBA module source code from an embedded `vbaProject.bin` OLE
/// container.
///
/// Must match `read_vba_modules(cfb_data: bytes) -> Dict[str, str]` in
/// `_vba_reader.py`: returns `{module_name: plain_text_source}`, and
/// returns an **empty map** (not an error) when the input is not a valid
/// VBA project — see the broad `except Exception: return {}` around the
/// call site in `XlsxWorkbook.iter_vba_modules`.
///
/// TODO(Wave 1C): implement — see module-level doc comment above.
pub fn read_vba_modules(_data: &[u8]) -> Result<BTreeMap<String, String>> {
    // STUB: Wave 1C replaces this with the real MS-CFB + MS-OVBA parser.
    Ok(BTreeMap::new())
}
