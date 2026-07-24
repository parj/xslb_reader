//! vba.rs — VBA macro extraction from embedded OLE (MS-CFB) containers.
//!
//! Port of `xlsb_reader/_vba_reader.py`. Implements:
//!   - MS-CFB compound-file reading, via the `cfb` crate (see
//!     [`cfb_read_streams`]) instead of hand-rolling FAT/MiniFAT walking.
//!     `cfb::CompoundFile::walk()` yields `Entry`s whose `path()` is an
//!     absolute `/`-joined `Path` (e.g. `/VBA/dir`); [`path_to_key`] turns
//!     that into the same `UPPERCASE/JOINED` key shape the Python
//!     `_cfb_read_streams` produces (e.g. `"VBA/DIR"`), so the stream
//!     lookups below (`dir`, `VBA/<module>`, case-insensitive fallback) are
//!     byte-for-byte the same logic as `read_vba_modules` in Python.
//!   - the MS-OVBA RLE decompressor ([`decompress`], §2.4.1.3.1).
//!   - `dir` stream parsing ([`parse_dir`], §2.3.4.2) to recover the module
//!     name -> stream name/offset table.
//!   - module stream extraction ([`extract_module_source`], §2.3.4.3): skip
//!     the compiled P-code performance cache up to `text_offset`, decompress
//!     the remainder, decode as Latin-1 (matches Python's
//!     `.decode("latin-1")` — a safe byte<->codepoint superset for the MBCS
//!     VBA source Excel writes).
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
//!
//! # Error handling contract
//!
//! [`read_vba_modules`] never returns `Err` — any failure anywhere in the
//! CFB/dir-stream parse (not a CFB file, missing `VBA/dir` stream, corrupt
//! MS-OVBA data, truncated `dir` stream records, ...) is swallowed and an
//! **empty map** is returned instead. This matches the pure-Python
//! contract: `read_vba_modules` itself can raise `ValueError`, but every
//! caller (`XlsxWorkbook.iter_vba_modules`) wraps the call in a broad
//! `except Exception: return {}`, so from the perspective of anything that
//! calls into VBA extraction, malformed input always yields `{}`. Failures
//! while decompressing/decoding one *specific* module's stream only drop
//! that module (mirrors the `try/except ... pass` around
//! `_extract_module_source` inside `read_vba_modules` itself) — the rest of
//! the modules are still returned.

use std::collections::BTreeMap;
use std::io::Read;
use std::path::{Component, Path, PathBuf};

use crate::common::{Result, XlsbError};

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

/// Extract VBA module source code from an embedded `vbaProject.bin` OLE
/// container.
///
/// Matches `read_vba_modules(cfb_data: bytes) -> Dict[str, str]` in
/// `_vba_reader.py`: returns `[(module_name, plain_text_source), ...]` in
/// the VBA project's own `dir`-stream declaration order (matching the
/// iteration order of Python's `dict`, built by inserting modules in that
/// same order), and returns an **empty list** (not an error) when the
/// input is not a valid VBA project — see the module-level doc comment
/// above.
pub fn read_vba_modules(data: &[u8]) -> Result<Vec<(String, String)>> {
    Ok(try_read_vba_modules(data).unwrap_or_default())
}

/// Fallible core of [`read_vba_modules`]. Any `Err` returned here is
/// equivalent to the pure-Python `read_vba_modules` raising an exception
/// that its caller catches broadly and turns into `{}`.
fn try_read_vba_modules(data: &[u8]) -> Result<Vec<(String, String)>> {
    let streams = cfb_read_streams(data)?;

    // Locate VBA/dir stream (case-insensitive), same rule as Python:
    // `k.endswith("/DIR") or k == "DIR"`.
    let dir_key = streams
        .keys()
        .find(|k| k.ends_with("/DIR") || k.as_str() == "DIR")
        .ok_or_else(|| XlsbError::Parse("VBA/dir stream not found in CFB container".into()))?
        .clone();

    let dir_data = decompress(&streams[&dir_key])?;
    let modules_meta = parse_dir(&dir_data)?;

    let mut result = Vec::new();
    for module in modules_meta {
        if module.stream.is_empty() {
            continue;
        }
        if let Some(stream_data) = find_module_stream(&streams, &module.stream) {
            if let Ok(source) = extract_module_source(stream_data, module.offset) {
                result.push((module.name, source));
            }
            // Modules that fail to decompress are silently skipped, same as
            // the `except Exception: pass` in `read_vba_modules`.
        }
    }

    Ok(result)
}

// ---------------------------------------------------------------------------
// CFB (OLE Compound File Binary) reading — via the `cfb` crate
// ---------------------------------------------------------------------------

/// Read every stream in the compound file at `data`, keyed by its
/// `/`-joined, UPPERCASE path relative to the root storage (e.g.
/// `"VBA/DIR"`, `"VBA/MODULE1"`, `"PROJECT"`) — same shape as Python's
/// `_cfb_read_streams`.
fn cfb_read_streams(data: &[u8]) -> Result<BTreeMap<String, Vec<u8>>> {
    let cursor = std::io::Cursor::new(data);
    let mut comp = cfb::CompoundFile::open(cursor)?;

    // Collect paths first: `walk()` borrows `comp` immutably, while
    // `open_stream()` below needs `&mut comp`.
    let stream_paths: Vec<PathBuf> = comp
        .walk()
        .filter(|e| e.is_stream())
        .map(|e| e.path().to_path_buf())
        .collect();

    let mut streams = BTreeMap::new();
    for path in stream_paths {
        let mut stream = comp.open_stream(&path)?;
        let mut buf = Vec::new();
        stream.read_to_end(&mut buf)?;
        streams.insert(path_to_key(&path), buf);
    }
    Ok(streams)
}

/// Turn a `cfb` entry path (e.g. `/VBA/dir`, absolute, `/`-separated) into
/// the same key shape Python's `_cfb_read_streams` produces: path
/// components joined with `/` (no leading slash), uppercased.
fn path_to_key(path: &Path) -> String {
    path.components()
        .filter_map(|c| match c {
            Component::Normal(s) => s.to_str(),
            _ => None,
        })
        .collect::<Vec<_>>()
        .join("/")
        .to_uppercase()
}

/// Find the stream backing a module's `MODULESTREAMNAME`, mirroring the two
/// lookups `read_vba_modules` performs in Python: first `VBA/<stream>`
/// (uppercased), then a case-insensitive scan for any key ending with
/// `/<STREAM>` or equal to `<STREAM>`.
fn find_module_stream<'a>(
    streams: &'a BTreeMap<String, Vec<u8>>,
    stream_name: &str,
) -> Option<&'a Vec<u8>> {
    let direct_key = format!("VBA/{stream_name}").to_uppercase();
    if let Some(v) = streams.get(&direct_key) {
        return Some(v);
    }
    let upper_name = stream_name.to_uppercase();
    let suffix = format!("/{upper_name}");
    streams
        .iter()
        .find(|(k, _)| k.ends_with(&suffix) || **k == upper_name)
        .map(|(_, v)| v)
}

// ---------------------------------------------------------------------------
// MS-OVBA Decompressor (§2.4.1.3.1)
// ---------------------------------------------------------------------------

/// Decompress an MS-OVBA `CompressedContainer`.
///
/// `CompressedContainer = SignatureByte(0x01) + CompressedChunk*`. Each
/// `CompressedChunk` is a 2-byte header followed by either 4096 raw bytes or
/// a sequence of literal/copy tokens. Direct port of `_decompress` in
/// `_vba_reader.py`, with out-of-range `CopyToken` offsets/lengths turned
/// into an [`XlsbError::Parse`] instead of Python's bytearray
/// negative-indexing/`IndexError` behavior (both are "this data doesn't
/// decompress cleanly", which is exactly the state this function should
/// surface as an error either way — see the module-level error-handling
/// contract doc comment).
fn decompress(compressed: &[u8]) -> Result<Vec<u8>> {
    if compressed.is_empty() || compressed[0] != 0x01 {
        let got = compressed.first().copied().unwrap_or(0);
        return Err(XlsbError::Parse(format!(
            "Invalid OVBA compressed data: expected SignatureByte 0x01, got 0x{got:02x}"
        )));
    }

    let mut out: Vec<u8> = Vec::new();
    let mut pos: usize = 1; // skip SignatureByte
    let n = compressed.len();

    while pos < n {
        if pos + 2 > n {
            break; // truncated
        }

        // --- CompressedChunkHeader (§2.4.1.1.5) ---
        let header = u16::from_le_bytes([compressed[pos], compressed[pos + 1]]);
        let chunk_size = ((header & 0x0FFF) as usize) + 3; // CompressedChunkSize
        let compressed_flag = (header >> 15) & 1; // 1=compressed 0=raw

        let chunk_start = pos;
        let decomp_chunk_st = out.len();
        let chunk_end = n.min(chunk_start + chunk_size);
        pos += 2; // past 2-byte header

        if compressed_flag == 0 {
            // Decompressing a RawChunk (§2.4.1.3.3): copy 4096 verbatim bytes.
            let end = n.min(pos + 4096);
            out.extend_from_slice(&compressed[pos..end]);
            pos += 4096;
        } else {
            // Decompressing via TokenSequences (§2.4.1.3.4).
            while pos < chunk_end {
                let flag_byte = compressed[pos];
                pos += 1;

                for bit_idx in 0..8u32 {
                    if pos >= chunk_end {
                        break;
                    }
                    let flag = (flag_byte >> bit_idx) & 1;

                    if flag == 0 {
                        // LiteralToken: copy one byte.
                        out.push(compressed[pos]);
                        pos += 1;
                    } else {
                        // CopyToken (§2.4.1.1.8).
                        if pos + 2 > chunk_end {
                            break;
                        }
                        let token = u16::from_le_bytes([compressed[pos], compressed[pos + 1]]);
                        pos += 2;

                        // CopyToken Help (§2.4.1.3.19.1) — derive bit masks.
                        let difference = out.len() - decomp_chunk_st;
                        let raw_bit_count = if difference > 1 {
                            (difference as f64).log2().ceil() as u32
                        } else {
                            1
                        };
                        let bit_count = raw_bit_count.max(4);
                        if bit_count >= 16 {
                            return Err(XlsbError::Parse(
                                "OVBA CopyToken: bit_count out of range".to_string(),
                            ));
                        }
                        let length_mask: u16 = 0xFFFFu16 >> bit_count;
                        let offset_mask: u16 = !length_mask;

                        // Unpack CopyToken (§2.4.1.3.19.2).
                        let length = ((token & length_mask) as usize) + 3;
                        let temp1 = token & offset_mask;
                        let temp2 = 16 - bit_count;
                        let offset = ((temp1 >> temp2) as usize) + 1;

                        // Byte Copy (§2.4.1.3.11) — source may overlap dest.
                        if offset > out.len() {
                            return Err(XlsbError::Parse(
                                "OVBA CopyToken: offset exceeds decompressed length".to_string(),
                            ));
                        }
                        for copy_src in (out.len() - offset..).take(length) {
                            out.push(out[copy_src]);
                        }
                    }
                }
            }
        }
    }

    Ok(out)
}

// ---------------------------------------------------------------------------
// dir Stream Parser (§2.3.4.2)
// ---------------------------------------------------------------------------

/// One module's metadata as recovered from the `dir` stream: its display
/// name (`MODULENAME`), backing stream name (`MODULESTREAMNAME`), and the
/// byte offset in that stream where the decompressed source text begins
/// (`MODULEOFFSET`).
///
/// The Python `_parse_dir` also tracks a `'procedural'`/`'class'` module
/// `type` (from `MODULETYPE` records `0x0021`/`0x0022`), but that field is
/// never read anywhere in `read_vba_modules` — only `name`/`stream`/`offset`
/// feed into the returned `{name: source}` map — so it is intentionally not
/// modelled here.
struct ModuleMeta {
    name: String,
    stream: String,
    offset: u32,
}

/// Parse a decompressed `dir` stream and return module descriptors in
/// declaration order. Direct port of `_parse_dir` in `_vba_reader.py`.
fn parse_dir(data: &[u8]) -> Result<Vec<ModuleMeta>> {
    let n = data.len();
    let mut pos: usize = 0;
    let mut modules: Vec<ModuleMeta> = Vec::new();
    let mut cur: Option<ModuleMeta> = None;

    while pos + 6 <= n {
        let rec_id = ru16(data, &mut pos)?;
        let rec_size = ru32(data, &mut pos)? as usize;

        if rec_size > n - pos {
            break;
        }

        match rec_id {
            // ── Simple pass-through records (generic Id + Size + Data) ────
            0x0001 // PROJECTSYSKIND
            | 0x004A // PROJECTCOMPATVERSION
            | 0x0002 // PROJECTLCID
            | 0x0014 // PROJECTLCIDINVOKE
            | 0x0003 // PROJECTCODEPAGE
            | 0x0004 // PROJECTNAME
            | 0x0007 // PROJECTHELPCONTEXT
            | 0x0008 // PROJECTLIBFLAGS
            | 0x000F // PROJECTMODULES Count
            | 0x0013 // PROJECTCOOKIE
            | 0x001E // MODULEHELPCONTEXT
            | 0x002C // MODULECOOKIE
            | 0x0047 // MODULENAMEUNICODE
            | 0x0025 // MODULEREADONLY
            | 0x0028 // MODULEPRIVATE
            => {
                skip(data, &mut pos, rec_size)?;
            }

            0x0009 => {
                // PROJECTVERSION: rec_size covers MajorVersion(4);
                // MinorVersion(2) follows.
                skip(data, &mut pos, rec_size)?;
                skip(data, &mut pos, 2)?;
            }

            0x0005 // PROJECTDOCSTRING
            | 0x0006 // PROJECTHELPFILEPATH
            | 0x000C // PROJECTCONSTANTS
            | 0x0016 // REFERENCENAME
            => {
                // MBCS payload + unicode companion.
                skip(data, &mut pos, rec_size)?;
                skip_unicode_pair(data, &mut pos)?;
            }

            0x000D | 0x000E => {
                // REFERENCEREGISTERED / REFERENCEPROJECT: Size covers the
                // entire payload.
                skip(data, &mut pos, rec_size)?;
            }

            0x002F => {
                // REFERENCECONTROL (complex multi-part record).
                // Part 1: SizeTwiddled bytes.
                skip(data, &mut pos, rec_size)?;
                // Part 2: optional NameRecordExtended (REFERENCENAME 0x0016).
                if peek_u16(data, pos) == Some(0x0016) {
                    skip(data, &mut pos, 2)?; // Id
                    let ns = ru32(data, &mut pos)? as usize;
                    skip(data, &mut pos, ns)?; // Name (MBCS)
                    skip_unicode_pair(data, &mut pos)?; // Reserved + SizeUnicode + Unicode
                }
                // Part 3: Reserved3(2) + SizeExtended(4) + SizeExtended bytes.
                if pos + 2 <= n {
                    skip(data, &mut pos, 2)?; // Reserved3
                }
                if pos + 4 <= n {
                    let ext_size = ru32(data, &mut pos)? as usize;
                    skip(data, &mut pos, ext_size)?;
                }
            }

            // ── MODULE records ─────────────────────────────────────────────
            0x0019 => {
                // MODULENAME — starts a new module.
                let payload_start = pos;
                skip(data, &mut pos, rec_size)?;
                let name = latin1_decode(&data[payload_start..pos]);
                if let Some(prev) = cur.take() {
                    if !prev.stream.is_empty() {
                        modules.push(prev);
                    }
                }
                cur = Some(ModuleMeta { name, stream: String::new(), offset: 0 });
            }

            0x001A => {
                // MODULESTREAMNAME: MBCS payload + unicode companion.
                let payload_start = pos;
                skip(data, &mut pos, rec_size)?;
                if let Some(m) = cur.as_mut() {
                    m.stream = latin1_decode(&data[payload_start..pos]);
                }
                skip_unicode_pair(data, &mut pos)?;
            }

            0x001C => {
                // MODULEDOCSTRING: MBCS payload + unicode companion.
                skip(data, &mut pos, rec_size)?;
                skip_unicode_pair(data, &mut pos)?;
            }

            0x0031 => {
                // MODULEOFFSET.
                let value = ru32(data, &mut pos)?;
                if let Some(m) = cur.as_mut() {
                    m.offset = value;
                }
                if rec_size > 4 {
                    skip(data, &mut pos, rec_size - 4)?;
                }
            }

            0x0021 | 0x0022 => {
                // MODULETYPE (procedural / class-or-document-or-designer).
                // Not needed for the returned {name: source} map — see the
                // doc comment on `ModuleMeta`. `rec_size` is always 0 here.
                skip(data, &mut pos, rec_size)?;
            }

            0x002B => {
                // MODULE Terminator.
                if let Some(prev) = cur.take() {
                    if !prev.stream.is_empty() {
                        modules.push(prev);
                    }
                }
                skip(data, &mut pos, rec_size)?; // rec_size == 0
                // Note: the OVBA spec defines 4 Reserved bytes here, but
                // Excel does not write them in practice — do not skip them.
            }

            _ => {
                // Unknown record — skip payload using Size field.
                skip(data, &mut pos, rec_size)?;
            }
        }
    }

    if let Some(prev) = cur {
        if !prev.stream.is_empty() {
            modules.push(prev);
        }
    }

    Ok(modules)
}

// ---------------------------------------------------------------------------
// dir stream cursor helpers
// ---------------------------------------------------------------------------
//
// A small hand-rolled position cursor (rather than `common::Reader`) because
// `parse_dir`/the REFERENCECONTROL record specifically need a non-consuming
// `peek_u16` to decide whether an optional sub-record is present, which
// `common::Reader` does not expose. Errors here (insufficient bytes) mirror
// the uncaught `struct.error`/slicing exceptions the Python implementation
// would raise in the same truncated-data situations.

fn ru16(data: &[u8], pos: &mut usize) -> Result<u16> {
    if *pos + 2 > data.len() {
        return Err(XlsbError::Parse(
            "dir stream: truncated u16 read".to_string(),
        ));
    }
    let v = u16::from_le_bytes([data[*pos], data[*pos + 1]]);
    *pos += 2;
    Ok(v)
}

fn ru32(data: &[u8], pos: &mut usize) -> Result<u32> {
    if *pos + 4 > data.len() {
        return Err(XlsbError::Parse(
            "dir stream: truncated u32 read".to_string(),
        ));
    }
    let v = u32::from_le_bytes([data[*pos], data[*pos + 1], data[*pos + 2], data[*pos + 3]]);
    *pos += 4;
    Ok(v)
}

fn skip(data: &[u8], pos: &mut usize, n: usize) -> Result<()> {
    if *pos + n > data.len() {
        return Err(XlsbError::Parse("dir stream: truncated skip".to_string()));
    }
    *pos += n;
    Ok(())
}

/// Peek a little-endian `u16` at `pos` without advancing, or `None` if fewer
/// than 2 bytes remain (matches Python's `if pos + 2 <= n and
/// struct.unpack_from("<H", data, pos)[0] == ...`).
fn peek_u16(data: &[u8], pos: usize) -> Option<u16> {
    if pos + 2 <= data.len() {
        Some(u16::from_le_bytes([data[pos], data[pos + 1]]))
    } else {
        None
    }
}

/// Skip the Unicode companion that follows some MBCS record payloads.
///
/// Layout: `Reserved(2 bytes) + SizeOfUnicode(4 bytes) + Unicode(Size bytes)`.
fn skip_unicode_pair(data: &[u8], pos: &mut usize) -> Result<()> {
    skip(data, pos, 2)?; // Reserved
    let uni_size = ru32(data, pos)? as usize;
    skip(data, pos, uni_size)?;
    Ok(())
}

// ---------------------------------------------------------------------------
// Module Stream Extractor (§2.3.4.3)
// ---------------------------------------------------------------------------

/// Decompress and decode the VBA source from a module stream.
///
/// Module stream layout:
///   - `PerformanceCache` (`text_offset` bytes — ignored on read)
///   - `CompressedSourceCode` (remainder — decompress to get VBA text)
fn extract_module_source(stream_data: &[u8], text_offset: u32) -> Result<String> {
    let text_offset = text_offset as usize;
    if text_offset > stream_data.len() {
        return Err(XlsbError::Parse(format!(
            "MODULEOFFSET {text_offset} exceeds stream size {}",
            stream_data.len()
        )));
    }
    let compressed = &stream_data[text_offset..];
    let decompressed = decompress(compressed)?;
    // VBA source is MBCS; latin-1 is a safe superset for ASCII VBA code.
    Ok(latin1_decode(&decompressed))
}

/// Decode bytes as Latin-1 (ISO-8859-1): every byte maps 1:1 to the Unicode
/// code point of the same value. Matches Python's `bytes.decode("latin-1")`
/// exactly (that codec never fails and never remaps).
fn latin1_decode(bytes: &[u8]) -> String {
    bytes.iter().map(|&b| b as char).collect()
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

#[cfg(test)]
mod tests {
    use super::*;

    /// Build a single raw (uncompressed) MS-OVBA chunk: header with the
    /// compressed flag cleared, then exactly 4096 payload bytes (padded with
    /// zeros), matching `_decompress`'s RawChunk handling.
    fn raw_chunk(payload: &[u8]) -> Vec<u8> {
        assert!(payload.len() <= 4096);
        let mut chunk = Vec::new();
        // CompressedChunkSize field is (header & 0x0FFF) + 3; for a raw
        // chunk of the max 4096-byte payload plus 2-byte header, the size
        // field is conventionally 4098 -> encoded value 0x0FFF (all raw
        // chunks are fixed 4096 bytes regardless of real payload length).
        let header: u16 = 0x0FFF; // compressed_flag bit (0x8000) = 0 => raw
        chunk.extend_from_slice(&header.to_le_bytes());
        let mut data = payload.to_vec();
        data.resize(4096, 0);
        chunk.extend_from_slice(&data);
        chunk
    }

    #[test]
    fn decompress_rejects_bad_signature() {
        let err = decompress(&[0x00, 0x01, 0x02]).unwrap_err();
        assert!(matches!(err, XlsbError::Parse(_)));
    }

    #[test]
    fn decompress_rejects_empty_input() {
        assert!(decompress(&[]).is_err());
    }

    #[test]
    fn decompress_raw_chunk_roundtrip() {
        let mut payload = b"Attribute VB_Name = \"Module1\"\r\n".to_vec();
        payload.resize(4096, 0);
        let mut compressed = vec![0x01u8]; // SignatureByte
        compressed.extend_from_slice(&raw_chunk(&payload));

        let out = decompress(&compressed).unwrap();
        assert_eq!(out, payload);
    }

    #[test]
    fn decompress_literal_only_token_sequence() {
        // One compressed chunk containing a single flag byte with all
        // literal (flag bit 0) tokens: "Sub Foo".
        let text = b"Sub Foo";
        let mut token_bytes = Vec::new();
        token_bytes.push(0x00u8); // flag byte: all 8 bits = LiteralToken
        token_bytes.extend_from_slice(text);
        // Pad remaining literal bits (8 total) with extra literal bytes so
        // the flag byte's 8 bits all correspond to real data; simplest is
        // to use exactly 8 literal bytes.
        while token_bytes.len() < 1 + 8 {
            token_bytes.push(b'!');
        }

        // header: chunk_size = token_bytes.len() + 2 - 3... actually encode
        // chunk_size as (len(token_bytes)) + 2 header bytes total content,
        // per spec ChunkSize = 4098 normally, but for a chunk shorter than
        // 4096 uncompressed we still set the 12-bit size field to
        // (compressed_chunk_data_len + 2 - 3) i.e. total_chunk_len - 3,
        // where total_chunk_len includes the 2-byte header.
        let total_chunk_len = 2 + token_bytes.len();
        let size_field = (total_chunk_len - 3) as u16 & 0x0FFF;
        let header: u16 = 0x8000 | size_field; // compressed_flag = 1

        let mut compressed = vec![0x01u8];
        compressed.extend_from_slice(&header.to_le_bytes());
        compressed.extend_from_slice(&token_bytes);

        let out = decompress(&compressed).unwrap();
        assert_eq!(out.len(), 8);
        assert_eq!(&out[..text.len()], text.as_slice());
        assert_eq!(out[7], b'!');
    }

    #[test]
    fn decompress_copy_token_backreference() {
        // Encode "AAAAAAAAAA" (10 'A's) as one literal 'A' followed by a
        // CopyToken that copies 9 more bytes from offset 1.
        //
        // After the first literal, decomp_chunk_st = 0, out.len() = 1, so
        // difference = 1 -> bit_count = max(1, 4) = 4.
        // length_mask = 0xFFFF >> 4 = 0x0FFF, offset_mask = 0xF000.
        // We want length = 9 -> (token & 0x0FFF) = 9 - 3 = 6.
        // We want offset = 1 -> temp1 = (offset - 1) << (16 - 4) = 0 << 12 = 0.
        // token = 0x0000 | 0x0006 = 0x0006.
        let flag_byte: u8 = 0b0000_0010; // bit0 = literal 'A', bit1 = CopyToken
        let mut token_bytes = vec![flag_byte, b'A'];
        let token: u16 = 0x0006;
        token_bytes.extend_from_slice(&token.to_le_bytes());

        let total_chunk_len = 2 + token_bytes.len();
        let size_field = (total_chunk_len - 3) as u16 & 0x0FFF;
        let header: u16 = 0x8000 | size_field;

        let mut compressed = vec![0x01u8];
        compressed.extend_from_slice(&header.to_le_bytes());
        compressed.extend_from_slice(&token_bytes);

        let out = decompress(&compressed).unwrap();
        assert_eq!(out, b"AAAAAAAAAA".to_vec());
    }

    #[test]
    fn decompress_copy_token_rejects_offset_beyond_output() {
        // Same shape as the backreference test, but request an offset that
        // cannot possibly be satisfied (larger than anything decompressed
        // so far), which must produce an error rather than panicking.
        let flag_byte: u8 = 0b0000_0010;
        let mut token_bytes = vec![flag_byte, b'A'];
        // temp1 = 0xF000 (offset field maxed out) -> offset = (0xF000 >> 12) + 1 = 16
        // out.len() at that point is only 1, so offset (16) > out.len() (1).
        let token: u16 = 0xF000;
        token_bytes.extend_from_slice(&token.to_le_bytes());

        let total_chunk_len = 2 + token_bytes.len();
        let size_field = (total_chunk_len - 3) as u16 & 0x0FFF;
        let header: u16 = 0x8000 | size_field;

        let mut compressed = vec![0x01u8];
        compressed.extend_from_slice(&header.to_le_bytes());
        compressed.extend_from_slice(&token_bytes);

        assert!(decompress(&compressed).is_err());
    }

    #[test]
    fn latin1_decode_is_byte_identity() {
        let bytes: Vec<u8> = (0u8..=255).collect();
        let s = latin1_decode(&bytes);
        assert_eq!(s.chars().count(), 256);
        for (i, c) in s.chars().enumerate() {
            assert_eq!(c as u32, i as u32);
        }
    }

    #[test]
    fn parse_dir_recovers_module_offset_and_stream() {
        // Build a minimal synthetic `dir` stream: MODULENAME "Module1",
        // MODULESTREAMNAME "Module1", MODULEOFFSET 42, MODULE terminator.
        let mut dir = Vec::new();

        // MODULENAME (0x0019): payload "Module1"
        let name = b"Module1";
        dir.extend_from_slice(&0x0019u16.to_le_bytes());
        dir.extend_from_slice(&(name.len() as u32).to_le_bytes());
        dir.extend_from_slice(name);

        // MODULESTREAMNAME (0x001A): payload "Module1" + unicode companion
        // (Reserved(2) + SizeOfUnicode(4)=0 + no unicode bytes)
        dir.extend_from_slice(&0x001Au16.to_le_bytes());
        dir.extend_from_slice(&(name.len() as u32).to_le_bytes());
        dir.extend_from_slice(name);
        dir.extend_from_slice(&0u16.to_le_bytes()); // Reserved
        dir.extend_from_slice(&0u32.to_le_bytes()); // SizeOfUnicode = 0

        // MODULEOFFSET (0x0031): size 4, value 42
        dir.extend_from_slice(&0x0031u16.to_le_bytes());
        dir.extend_from_slice(&4u32.to_le_bytes());
        dir.extend_from_slice(&42u32.to_le_bytes());

        // MODULE Terminator (0x002B): size 0
        dir.extend_from_slice(&0x002Bu16.to_le_bytes());
        dir.extend_from_slice(&0u32.to_le_bytes());

        let modules = parse_dir(&dir).unwrap();
        assert_eq!(modules.len(), 1);
        assert_eq!(modules[0].name, "Module1");
        assert_eq!(modules[0].stream, "Module1");
        assert_eq!(modules[0].offset, 42);
    }

    #[test]
    fn parse_dir_drops_module_with_no_stream_name() {
        // MODULENAME without a following MODULESTREAMNAME should not be
        // emitted (matches Python's `cur.get("stream")` truthiness check).
        let mut dir = Vec::new();
        let name = b"Orphan";
        dir.extend_from_slice(&0x0019u16.to_le_bytes());
        dir.extend_from_slice(&(name.len() as u32).to_le_bytes());
        dir.extend_from_slice(name);
        dir.extend_from_slice(&0x002Bu16.to_le_bytes());
        dir.extend_from_slice(&0u32.to_le_bytes());

        let modules = parse_dir(&dir).unwrap();
        assert!(modules.is_empty());
    }

    #[test]
    fn extract_module_source_skips_performance_cache() {
        // PerformanceCache of 5 arbitrary bytes, then a raw (uncompressed)
        // MS-OVBA chunk holding "Hi".
        let mut stream = vec![0xAA; 5];
        let mut payload = b"Hi".to_vec();
        payload.resize(4096, 0);
        stream.push(0x01); // SignatureByte
        stream.extend_from_slice(&raw_chunk(&payload));

        let source = extract_module_source(&stream, 5).unwrap();
        assert!(source.starts_with("Hi"));
    }

    #[test]
    fn extract_module_source_rejects_offset_past_end() {
        let stream = vec![0x01, 0x00, 0x00];
        assert!(extract_module_source(&stream, 100).is_err());
    }

    #[test]
    fn read_vba_modules_returns_empty_map_for_garbage_input() {
        let result = read_vba_modules(b"not a cfb file at all").unwrap();
        assert!(result.is_empty());
    }

    #[test]
    fn read_vba_modules_returns_empty_map_for_empty_input() {
        let result = read_vba_modules(&[]).unwrap();
        assert!(result.is_empty());
    }
}
