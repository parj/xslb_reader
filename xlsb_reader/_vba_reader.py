"""
_vba_reader.py
==============
Pure-Python VBA macro extractor for Excel workbooks.

Implements:
  - OLE Compound File Binary (CFB/OLE2) reader — MS-CFB
  - MS-OVBA decompressor — §2.4.1.3.1
  - dir stream parser — §2.3.4.2
  - Module stream extractor — §2.3.4.3

Public API:
    read_vba_modules(cfb_data: bytes) -> Dict[str, str]
        Returns {module_name: vba_source_text}.

No third-party dependencies — stdlib only.
"""

from __future__ import annotations

import math
import struct
from typing import Dict, List, Optional

# ---------------------------------------------------------------------------
# CFB (OLE Compound File Binary) constants
# ---------------------------------------------------------------------------

_FREESECT = 0xFFFFFFFF
_ENDOFCHAIN = 0xFFFFFFFE
_FATSECT = 0xFFFFFFFD
_DIFSECT = 0xFFFFFFFC
_NOSTREAM = 0xFFFFFFFF

_CFB_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"

# ---------------------------------------------------------------------------
# CFB reader
# ---------------------------------------------------------------------------


def _cfb_read_streams(data: bytes) -> Dict[str, bytes]:
    """
    Parse an OLE Compound File Binary and return all streams as
    {UPPERCASE_PATH: bytes}.

    Paths use '/' separator relative to the root storage.
    Examples: 'VBA/DIR', 'VBA/MODULE1', 'PROJECT'
    """
    if len(data) < 512 or data[:8] != _CFB_MAGIC:
        raise ValueError("Not a valid CFB/OLE2 file (bad magic)")

    # Header fields (all little-endian)
    sec_size_exp = struct.unpack_from("<H", data, 30)[0]
    mini_sec_size_exp = struct.unpack_from("<H", data, 32)[0]
    first_dir_sec = struct.unpack_from("<I", data, 48)[0]
    mini_cutoff = struct.unpack_from("<I", data, 56)[0]
    first_minifat_sec = struct.unpack_from("<I", data, 60)[0]

    sec_size = 1 << sec_size_exp  # 512 for major version 3
    mini_size = 1 << mini_sec_size_exp  # typically 64

    def sec_offset(sec_id: int) -> int:
        return 512 + sec_id * sec_size  # header occupies first 512 bytes

    # Build FAT from DIFAT entries embedded in the header (up to 109 sectors)
    difat = struct.unpack_from("<109I", data, 76)
    fat_bytes = bytearray()
    for sec_id in difat:
        if sec_id in (_FREESECT, _ENDOFCHAIN, _DIFSECT):
            break
        off = sec_offset(sec_id)
        fat_bytes.extend(data[off : off + sec_size])

    n_fat = len(fat_bytes) // 4
    fat = struct.unpack_from(f"<{n_fat}I", fat_bytes) if n_fat else ()

    def read_chain(first_sec: int) -> bytes:
        """Follow FAT chain and concatenate sector contents."""
        chunks: List[bytes] = []
        sec = first_sec
        while sec not in (_ENDOFCHAIN, _FREESECT) and sec < n_fat:
            off = sec_offset(sec)
            chunks.append(data[off : off + sec_size])
            sec = fat[sec]
        return b"".join(chunks)

    # Directory
    dir_bytes = read_chain(first_dir_sec)
    n_entries = len(dir_bytes) // 128

    entries: Dict[int, Optional[dict]] = {}
    for i in range(n_entries):
        e = dir_bytes[i * 128 : (i + 1) * 128]
        name_len = struct.unpack_from("<H", e, 64)[0]
        if name_len < 2:
            entries[i] = None
            continue
        name = e[: name_len - 2].decode("utf-16-le", errors="replace")
        obj_type = e[66]  # 0=empty 1=storage 2=stream 5=root
        left_id = struct.unpack_from("<I", e, 68)[0]
        right_id = struct.unpack_from("<I", e, 72)[0]
        child_id = struct.unpack_from("<I", e, 76)[0]
        start = struct.unpack_from("<I", e, 116)[0]
        size = struct.unpack_from("<I", e, 120)[0]
        entries[i] = {
            "name": name,
            "type": obj_type,
            "left": left_id,
            "right": right_id,
            "child": child_id,
            "start": start,
            "size": size,
        }

    # Mini FAT
    if first_minifat_sec not in (_ENDOFCHAIN, _FREESECT):
        mf_bytes = read_chain(first_minifat_sec)
        n_mf = len(mf_bytes) // 4
        mini_fat: tuple = struct.unpack_from(f"<{n_mf}I", mf_bytes) if n_mf else ()
    else:
        mini_fat = ()

    # Mini stream lives in root entry's (SID 0) regular-sector chain
    root = entries.get(0)
    mini_stream = b""
    if root and root["start"] not in (_ENDOFCHAIN, _FREESECT):
        mini_stream = read_chain(root["start"])

    def read_mini_chain(first_sec: int, size: int) -> bytes:
        chunks: List[bytes] = []
        sec = first_sec
        while sec not in (_ENDOFCHAIN, _FREESECT) and sec < len(mini_fat):
            off = sec * mini_size
            chunks.append(mini_stream[off : off + mini_size])
            sec = mini_fat[sec]
        return b"".join(chunks)[:size]

    def read_stream(entry: dict) -> bytes:
        # Streams smaller than mini_cutoff live in the mini stream
        if (
            entry["size"] < mini_cutoff
            and root
            and root["start"] not in (_ENDOFCHAIN, _FREESECT)
            and mini_fat
        ):
            return read_mini_chain(entry["start"], entry["size"])
        return read_chain(entry["start"])[: entry["size"]]

    # Walk the red-black tree rooted at a given SID
    def walk_rb(sid: int):
        if sid == _NOSTREAM or sid >= len(entries):
            return
        e = entries.get(sid)
        if e is None:
            return
        yield sid
        yield from walk_rb(e["left"])
        yield from walk_rb(e["right"])

    # DFS over the directory tree collecting all streams with paths
    def collect(parent_sid: int, path_prefix: str, result: Dict[str, bytes]) -> None:
        e = entries.get(parent_sid)
        if e is None:
            return
        for sid in walk_rb(e["child"]):
            child = entries.get(sid)
            if child is None:
                continue
            child_path = (
                f"{path_prefix}/{child['name']}" if path_prefix else child["name"]
            )
            if child["type"] == 2:  # stream
                result[child_path.upper()] = read_stream(child)
            elif child["type"] in (1, 5):  # storage
                collect(sid, child_path, result)

    streams: Dict[str, bytes] = {}
    collect(0, "", streams)
    return streams


# ---------------------------------------------------------------------------
# MS-OVBA Decompressor (§2.4.1.3.1)
# ---------------------------------------------------------------------------


def _decompress(compressed: bytes) -> bytes:
    """Decompress an MS-OVBA CompressedContainer.

    CompressedContainer = SignatureByte(0x01) + CompressedChunk*
    Each CompressedChunk: 2-byte header + (raw 4096 bytes  OR  token sequences)
    """
    if not compressed or compressed[0] != 0x01:
        raise ValueError(
            f"Invalid OVBA compressed data: expected SignatureByte 0x01, "
            f"got 0x{compressed[0]:02x}"
        )

    out = bytearray()
    pos = 1  # skip SignatureByte
    n = len(compressed)

    while pos < n:
        if pos + 2 > n:
            break  # truncated

        # --- CompressedChunkHeader (§2.4.1.1.5) ---
        header = struct.unpack_from("<H", compressed, pos)[0]
        chunk_size = (header & 0x0FFF) + 3  # Extract CompressedChunkSize
        compressed_flag = (header >> 15) & 1  # 1=compressed  0=raw

        chunk_start = pos
        decomp_chunk_st = len(out)
        chunk_end = min(n, chunk_start + chunk_size)
        pos += 2  # past 2-byte header

        if compressed_flag == 0:
            # Decompressing a RawChunk (§2.4.1.3.3): copy 4096 verbatim bytes
            out.extend(compressed[pos : pos + 4096])
            pos += 4096
        else:
            # Decompressing via TokenSequences (§2.4.1.3.4)
            while pos < chunk_end:
                flag_byte = compressed[pos]
                pos += 1

                for bit_idx in range(8):
                    if pos >= chunk_end:
                        break
                    flag = (flag_byte >> bit_idx) & 1

                    if flag == 0:
                        # LiteralToken: copy one byte
                        out.append(compressed[pos])
                        pos += 1
                    else:
                        # CopyToken (§2.4.1.1.8)
                        if pos + 2 > chunk_end:
                            break
                        token = struct.unpack_from("<H", compressed, pos)[0]
                        pos += 2

                        # CopyToken Help (§2.4.1.3.19.1) — derive bit masks
                        difference = len(out) - decomp_chunk_st
                        bit_count = max(
                            math.ceil(math.log2(difference)) if difference > 1 else 1, 4
                        )
                        length_mask = 0xFFFF >> bit_count
                        offset_mask = (~length_mask) & 0xFFFF

                        # Unpack CopyToken (§2.4.1.3.19.2)
                        length = (token & length_mask) + 3
                        temp1 = token & offset_mask
                        temp2 = 16 - bit_count
                        offset = (temp1 >> temp2) + 1

                        # Byte Copy (§2.4.1.3.11) — source may overlap dest
                        copy_src = len(out) - offset
                        for _ in range(length):
                            out.append(out[copy_src])
                            copy_src += 1

    return bytes(out)


# ---------------------------------------------------------------------------
# dir Stream Parser (§2.3.4.2)
# ---------------------------------------------------------------------------


def _parse_dir(data: bytes) -> List[dict]:
    """
    Parse a decompressed dir stream and return a list of module descriptors:
      [{'name': str, 'stream': str, 'offset': int, 'type': str}, ...]

    'type' is 'procedural' or 'class'.
    """
    modules: List[dict] = []
    cur: Optional[dict] = None
    pos = 0
    n = len(data)

    def ru16() -> int:
        nonlocal pos
        v = struct.unpack_from("<H", data, pos)[0]
        pos += 2
        return v

    def ru32() -> int:
        nonlocal pos
        v = struct.unpack_from("<I", data, pos)[0]
        pos += 4
        return v

    def skip_unicode_pair() -> None:
        """Skip the Unicode companion that follows some MBCS record payloads.

        Layout: Reserved(2 bytes) + SizeOfUnicode(4 bytes) + Unicode(Size bytes)
        """
        nonlocal pos
        pos += 2  # Reserved
        uni_size = ru32()
        pos += uni_size

    while pos + 6 <= n:
        rec_id = ru16()
        rec_size = ru32()

        if pos + rec_size > n:
            break

        # ── Simple pass-through records (generic Id + Size + Data) ────────
        if rec_id in (
            0x0001,  # PROJECTSYSKIND
            0x004A,  # PROJECTCOMPATVERSION
            0x0002,  # PROJECTLCID
            0x0014,  # PROJECTLCIDINVOKE
            0x0003,  # PROJECTCODEPAGE
            0x0004,  # PROJECTNAME
            0x0007,  # PROJECTHELPCONTEXT
            0x0008,  # PROJECTLIBFLAGS
            0x000F,  # PROJECTMODULES Count
            0x0013,  # PROJECTCOOKIE
            0x001E,  # MODULEHELPCONTEXT
            0x002C,  # MODULECOOKIE
            0x0047,  # MODULENAMEUNICODE
            0x0025,  # MODULEREADONLY
            0x0028,  # MODULEPRIVATE
        ):
            pos += rec_size

        elif rec_id == 0x0009:
            # PROJECTVERSION: rec_size covers MajorVersion(4); MinorVersion(2) follows
            pos += rec_size  # MajorVersion
            pos += 2  # MinorVersion

        elif rec_id in (0x0005, 0x0006, 0x000C):
            # PROJECTDOCSTRING / PROJECTHELPFILEPATH / PROJECTCONSTANTS
            # MBCS payload + unicode companion
            pos += rec_size
            skip_unicode_pair()

        # ── PROJECTREFERENCES records ─────────────────────────────────────
        elif rec_id == 0x0016:
            # REFERENCENAME: MBCS payload + unicode companion
            pos += rec_size
            skip_unicode_pair()

        elif rec_id in (0x000D, 0x000E):
            # REFERENCEREGISTERED / REFERENCEPROJECT
            # Size covers the entire payload
            pos += rec_size

        elif rec_id == 0x002F:
            # REFERENCECONTROL (complex multi-part record)
            # Part 1: SizeTwiddled bytes
            pos += rec_size
            # Part 2: optional NameRecordExtended (REFERENCENAME 0x0016)
            if pos + 2 <= n and struct.unpack_from("<H", data, pos)[0] == 0x0016:
                pos += 2  # Id
                ns = ru32()
                pos += ns  # Name (MBCS)
                skip_unicode_pair()  # Reserved + SizeUnicode + Unicode
            # Part 3: Reserved3(2) + SizeExtended(4) + SizeExtended bytes
            if pos + 2 <= n:
                pos += 2  # Reserved3
            if pos + 4 <= n:
                ext_size = ru32()
                pos += ext_size

        # ── MODULE records ────────────────────────────────────────────────
        elif rec_id == 0x0019:
            # MODULENAME — starts a new module
            if cur is not None and cur.get("stream"):
                modules.append(cur)
            cur = {
                "name": data[pos : pos + rec_size].decode("latin-1"),
                "stream": "",
                "offset": 0,
                "type": "procedural",
            }
            pos += rec_size

        elif rec_id == 0x001A:
            # MODULESTREAMNAME: MBCS payload + unicode companion
            if cur is not None:
                cur["stream"] = data[pos : pos + rec_size].decode("latin-1")
            pos += rec_size
            skip_unicode_pair()

        elif rec_id == 0x001C:
            # MODULEDOCSTRING: MBCS payload + unicode companion
            pos += rec_size
            skip_unicode_pair()

        elif rec_id == 0x0031:
            # MODULEOFFSET
            if cur is not None:
                cur["offset"] = struct.unpack_from("<I", data, pos)[0]
            pos += rec_size

        elif rec_id == 0x0021:
            # MODULETYPE: procedural module
            if cur is not None:
                cur["type"] = "procedural"
            pos += rec_size  # rec_size == 0

        elif rec_id == 0x0022:
            # MODULETYPE: class / document / designer module
            if cur is not None:
                cur["type"] = "class"
            pos += rec_size  # rec_size == 0

        elif rec_id == 0x002B:
            # MODULE Terminator
            if cur is not None and cur.get("stream"):
                modules.append(cur)
                cur = None
            pos += rec_size  # rec_size == 0
            # Note: the OVBA spec defines 4 Reserved bytes here, but Excel
            # does not write them in practice — do not skip them.

        else:
            # Unknown record — skip payload using Size field
            pos += rec_size

    if cur is not None and cur.get("stream"):
        modules.append(cur)

    return modules


# ---------------------------------------------------------------------------
# Module Stream Extractor (§2.3.4.3)
# ---------------------------------------------------------------------------


def _extract_module_source(stream_data: bytes, text_offset: int) -> str:
    """Decompress and decode the VBA source from a module stream.

    Module stream layout:
        PerformanceCache  (TextOffset bytes — ignored on read)
        CompressedSourceCode (remainder — decompress to get VBA text)
    """
    if text_offset > len(stream_data):
        raise ValueError(
            f"MODULEOFFSET {text_offset} exceeds stream size {len(stream_data)}"
        )
    compressed = stream_data[text_offset:]
    decompressed = _decompress(compressed)
    # VBA source is MBCS; latin-1 is a safe superset for ASCII VBA code
    return decompressed.decode("latin-1")


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------


def read_vba_modules(cfb_data: bytes) -> Dict[str, str]:
    """Extract VBA module source code from vbaProject.bin data.

    Args:
        cfb_data: Raw bytes of an OLE Compound File Binary (vbaProject.bin).

    Returns:
        Dict mapping module name → plain-text VBA source (including Attribute
        lines as stored in the file).

    Raises:
        ValueError: If the data is not a valid CFB or the VBA/dir stream is
            missing.
    """
    streams = _cfb_read_streams(cfb_data)

    # Locate VBA/dir stream (case-insensitive)
    dir_key = next((k for k in streams if k.endswith("/DIR") or k == "DIR"), None)
    if dir_key is None:
        raise ValueError("VBA/dir stream not found in CFB container")

    dir_data = _decompress(streams[dir_key])
    mods_meta = _parse_dir(dir_data)

    result: Dict[str, str] = {}
    for mod in mods_meta:
        stream_name = mod.get("stream", "")
        if not stream_name:
            continue

        # Find the module stream: look under VBA/ prefix (case-insensitive)
        stream_key = f"VBA/{stream_name}".upper()
        stream_data = streams.get(stream_key)

        if stream_data is None:
            # Fallback: search all keys ending with the stream name
            upper_name = stream_name.upper()
            for k, v in streams.items():
                if k.endswith(f"/{upper_name}") or k == upper_name:
                    stream_data = v
                    break

        if stream_data is not None:
            try:
                source = _extract_module_source(stream_data, mod["offset"])
                result[mod["name"]] = source
            except Exception:
                pass  # skip modules that fail to decompress

    return result
