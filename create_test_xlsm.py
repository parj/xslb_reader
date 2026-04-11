#!/usr/bin/env python3
"""
create_test_xlsm.py
====================
Generate test-data/Finance_Ledger_VBA.xlsm containing known VBA modules.

The file is a zip archive (OOXML) whose xl/vbaProject.bin is a hand-crafted
OLE Compound File Binary (CFB/OLE2) containing:

    Root Entry (SID 0, storage)
    ├── VBA (SID 1, storage)
    │   ├── _VBA_PROJECT (SID 2, stream)  — version header, 7 bytes
    │   ├── dir          (SID 3, stream)  — compressed dir stream
    │   ├── Module1      (SID 4, stream)  — compressed VBA source
    │   └── Module2      (SID 5, stream)  — compressed VBA source
    └── PROJECT          (SID 6, stream)  — plain-text project properties

All data streams are short (< 4096 bytes) and live in the mini stream.

No third-party libraries required — stdlib only.
"""

from __future__ import annotations

import io
import struct
import zipfile
from pathlib import Path
from typing import List, Tuple

# ---------------------------------------------------------------------------
# MS-OVBA literals-only compressor (§2.4.1.3)
# ---------------------------------------------------------------------------


def _compress(data: bytes) -> bytes:
    """Compress using CompressedFlag=1 with only literal tokens (no back-refs).

    Valid per MS-OVBA spec; not space-optimal but simple and correct.

    Format:
        SignatureByte (0x01)
        For each up-to-4096-byte chunk:
            CompressedChunkHeader (2 bytes, CompressedFlag=1)
            Groups of up to 8 bytes, each preceded by FlagByte=0x00
    """
    out = bytearray([0x01])  # SignatureByte
    pos = 0
    n = len(data)

    while pos < n or pos == 0:
        chunk = data[pos : pos + 4096]
        pos += 4096

        hdr_idx = len(out)
        out.extend(b"\x00\x00")  # header placeholder

        # Encode chunk as literal tokens grouped 8 per FlagByte
        cp = 0
        while cp < len(chunk):
            out.append(0x00)  # FlagByte: all 8 bits = 0 → all literals
            for _ in range(8):
                if cp >= len(chunk):
                    break
                out.append(chunk[cp])
                cp += 1

        chunk_size = len(out) - hdr_idx
        header = ((chunk_size - 3) & 0x0FFF) | (0b011 << 12) | (1 << 15)
        struct.pack_into("<H", out, hdr_idx, header)

        if pos >= n:
            break

    return bytes(out)


# ---------------------------------------------------------------------------
# dir stream builder (MS-OVBA §2.3.4.2)
# ---------------------------------------------------------------------------


def _r(rec_id: int, payload: bytes) -> bytes:
    """Pack a standard record: Id(2) + Size(4) + payload."""
    return struct.pack("<HI", rec_id, len(payload)) + payload


def _build_dir_stream(mods: List[Tuple[str, str, int]]) -> bytes:
    """Build an uncompressed dir stream.

    mods: [(module_name, stream_name, text_offset), ...]
    """
    d = bytearray()

    # ── Project Information (§2.3.4.2.1) ─────────────────────────────────
    d += _r(0x0001, struct.pack("<I", 0x00000001))  # PROJECTSYSKIND = Win32
    d += _r(0x0002, struct.pack("<I", 0x00000409))  # PROJECTLCID
    d += _r(0x0014, struct.pack("<I", 0x00000409))  # PROJECTLCIDINVOKE
    d += _r(0x0003, struct.pack("<H", 0x04E4))  # PROJECTCODEPAGE = 1252
    d += _r(0x0004, b"VBAProject")  # PROJECTNAME

    # PROJECTDOCSTRING: MBCS(empty) + Reserved(0x0040) + Unicode(empty)
    d += struct.pack("<HI", 0x0005, 0)
    d += struct.pack("<HI", 0x0040, 0)

    # PROJECTHELPFILEPATH: MBCS(empty) + Reserved(0x003D) + HelpFile2(empty)
    d += struct.pack("<HI", 0x0006, 0)
    d += struct.pack("<HI", 0x003D, 0)

    d += _r(0x0007, struct.pack("<I", 0x00000000))  # PROJECTHELPCONTEXT
    d += _r(0x0008, struct.pack("<I", 0x00000000))  # PROJECTLIBFLAGS

    # PROJECTVERSION: Id(2) + Size=4(4) + MajorVersion(4) + MinorVersion(2)
    d += struct.pack("<HIIH", 0x0009, 0x00000004, 0x49B5196B, 0x0006)

    # PROJECTCONSTANTS: MBCS(empty) + Reserved(0x003C) + Unicode(empty)
    d += struct.pack("<HI", 0x000C, 0)
    d += struct.pack("<HI", 0x003C, 0)

    # ── No external references ────────────────────────────────────────────

    # ── Project Modules (§2.3.4.2.3) ─────────────────────────────────────
    d += _r(0x000F, struct.pack("<H", len(mods)))  # Count
    d += _r(0x0013, struct.pack("<H", 0xFFFF))  # PROJECTCOOKIE

    for mod_name, stream_name, text_offset in mods:
        mb_name = mod_name.encode("latin-1")
        uc_name = mod_name.encode("utf-16-le")
        mb_str = stream_name.encode("latin-1")
        uc_str = stream_name.encode("utf-16-le")

        d += _r(0x0019, mb_name)  # MODULENAME
        d += _r(0x0047, uc_name)  # MODULENAMEUNICODE

        # MODULESTREAMNAME: MBCS + Reserved(0x0032) + Unicode
        d += struct.pack("<HI", 0x001A, len(mb_str)) + mb_str
        d += struct.pack("<HI", 0x0032, len(uc_str)) + uc_str

        # MODULEDOCSTRING: MBCS(empty) + Reserved(0x0048) + Unicode(empty)
        d += struct.pack("<HI", 0x001C, 0)
        d += struct.pack("<HI", 0x0048, 0)

        d += _r(0x0031, struct.pack("<I", text_offset))  # MODULEOFFSET
        d += _r(0x001E, struct.pack("<I", 0x00000000))  # MODULEHELPCONTEXT
        d += _r(0x002C, struct.pack("<H", 0xFFFF))  # MODULECOOKIE

        # MODULETYPE: procedural = 0x0021, Reserved=0
        d += struct.pack("<HI", 0x0021, 0x00000000)

        # MODULE Terminator (0x002B) + Reserved(4 bytes)
        d += struct.pack("<HII", 0x002B, 0x00000000, 0x00000000)

    return bytes(d)


# ---------------------------------------------------------------------------
# CFB (OLE Compound File Binary) writer (minimal, version 3)
# ---------------------------------------------------------------------------

_SEC = 512  # regular sector size (version 3)
_MSZ = 64  # mini sector size
_CUT = 4096  # mini stream cutoff

_ENDOFCHAIN = 0xFFFFFFFE
_FREESECT = 0xFFFFFFFF
_FATSECT = 0xFFFFFFFD
_NOSTREAM = 0xFFFFFFFF


def _dir_entry(
    name: str,
    obj_type: int,
    *,
    left: int = _NOSTREAM,
    right: int = _NOSTREAM,
    child: int = _NOSTREAM,
    start: int = _ENDOFCHAIN,
    size: int = 0,
    clsid: bytes = b"\x00" * 16,
) -> bytes:
    """Build a 128-byte CFB directory entry."""
    enc = name.encode("utf-16-le")
    nlen = len(enc) + 2  # including null terminator
    pad = enc.ljust(64, b"\x00")[:64]

    return (
        pad
        + struct.pack("<H", nlen if name else 0)
        + struct.pack("B", obj_type)  # ObjectType
        + struct.pack("B", 1)  # ColorFlag = black
        + struct.pack("<I", left)
        + struct.pack("<I", right)
        + struct.pack("<I", child)
        + clsid[:16].ljust(16, b"\x00")
        + struct.pack("<I", 0)  # StateBits
        + b"\x00" * 8  # Created
        + b"\x00" * 8  # Modified
        + struct.pack("<I", start)
        + struct.pack("<I", size)
        + struct.pack("<I", 0)  # SizeHigh (v3 = 0)
    )


def _build_vba_project_bin(modules: List[Tuple[str, str, bytes]]) -> bytes:
    """Build a complete vbaProject.bin CFB file.

    modules: [(module_name, stream_name, vba_source_bytes), ...]
    """

    # ── Build raw stream contents ─────────────────────────────────────────

    # _VBA_PROJECT stream (7 bytes, §3.1.1, Version=0xFFFF for interoperability)
    vba_proj_data = bytes([0xCC, 0x61, 0xFF, 0xFF, 0x00, 0x01, 0x00])

    # dir stream
    mod_meta = [(name, sname, 0) for name, sname, _ in modules]
    dir_data = _compress(_build_dir_stream(mod_meta))

    # Module streams: PerformanceCache(0 bytes, TextOffset=0) + CompressedSource
    mod_data_list: List[bytes] = [_compress(src) for _, _, src in modules]

    # PROJECT stream (plain text, §2.3.1)
    proj_lines = [
        'ID="{00000000-0000-0000-0000-000000000000}"',
        *[f"Module={name}" for name, _, _ in modules],
        'Name="VBAProject"',
        'HelpContextID="0"',
        'VersionCompatible32="393222000"',
        'CMG="CACACACACACA"',
        'DPB="DADA"',
        'GC="GCGCGC"',
        "",
        "[Host Extender Info]",
        "",
    ]
    proj_data = "\r\n".join(proj_lines).encode("latin-1")

    # ── Assign mini-sector chains ─────────────────────────────────────────
    # Ordered list of (cfb_stream_name, raw_data) in the CFB
    # SIDs: 0=Root, 1=VBA storage, 2=_VBA_PROJECT, 3=dir,
    #        4..4+n_mods-1=modules, 4+n_mods=PROJECT
    streams_in_vba: List[Tuple[str, bytes]] = [
        ("_VBA_PROJECT", vba_proj_data),
        ("dir", dir_data),
        *[(sname, md) for (_, sname, _), md in zip(modules, mod_data_list)],
    ]
    proj_stream = ("PROJECT", proj_data)

    all_mini: List[Tuple[str, bytes]] = streams_in_vba + [proj_stream]

    def msecs(size: int) -> int:
        return max(1, (size + _MSZ - 1) // _MSZ)

    mini_starts: dict = {}
    mini_fat: List[int] = []
    cur_ms = 0
    for sname, sdata in all_mini:
        n = msecs(len(sdata))
        mini_starts[sname] = cur_ms
        for i in range(n - 1):
            mini_fat.append(cur_ms + 1)
            cur_ms += 1
        mini_fat.append(_ENDOFCHAIN)
        cur_ms += 1

    # Build mini stream blob
    ms_blob = bytearray()
    for sname, sdata in all_mini:
        n = msecs(len(sdata))
        ms_blob.extend(sdata)
        ms_blob.extend(b"\x00" * (n * _MSZ - len(sdata)))

    # ── Regular sector layout ─────────────────────────────────────────────
    n_dir_ent = 2 + len(all_mini)  # Root + VBA + streams
    n_dir_sec = max(1, (n_dir_ent + 3) // 4)
    n_mf_sec = max(1, (len(mini_fat) * 4 + _SEC - 1) // _SEC)
    n_ms_sec = (len(ms_blob) + _SEC - 1) // _SEC

    # Sector indices
    I_FAT = 0
    I_DIR = 1  # 1 .. n_dir_sec
    I_MF = I_DIR + n_dir_sec  # mini FAT sectors
    I_MS = I_MF + n_mf_sec  # mini stream sectors
    N_TOTAL = I_MS + n_ms_sec

    # ── Build FAT table ───────────────────────────────────────────────────
    fat = [_FREESECT] * N_TOTAL
    fat[I_FAT] = _FATSECT
    for i in range(I_DIR, I_MF - 1):
        fat[i] = i + 1
    fat[I_MF - 1] = _ENDOFCHAIN
    for i in range(I_MF, I_MS - 1):
        fat[i] = i + 1
    fat[I_MS - 1] = _ENDOFCHAIN
    for i in range(I_MS, N_TOTAL - 1):
        fat[i] = i + 1
    fat[N_TOTAL - 1] = _ENDOFCHAIN

    # ── Build directory entries ───────────────────────────────────────────
    # SID map: 0=Root, 1=VBA, 2=_VBA_PROJECT, 3=dir, 4..=modules, last=PROJECT
    sid_vba = 1
    vba_sid0 = 2  # first child of VBA (_VBA_PROJECT)
    vba_sids = list(range(2, 2 + len(streams_in_vba)))  # SIDs of VBA's children
    sid_proj = 2 + len(streams_in_vba)  # PROJECT SID (right of VBA in Root tree)

    dir_bytes = bytearray()

    # Root (SID 0): type=5, child=VBA(1), data=mini stream
    dir_bytes += _dir_entry(
        "Root Entry",
        5,
        child=sid_vba,
        start=I_MS,
        size=len(ms_blob),
        clsid=bytes(
            [
                0x06,
                0x09,
                0x02,
                0x00,
                0x00,
                0x00,
                0x00,
                0x00,
                0xC0,
                0x00,
                0x00,
                0x00,
                0x00,
                0x00,
                0x00,
                0x46,
            ]
        ),
    )

    # VBA storage (SID 1): right sibling = PROJECT in Root's child tree
    dir_bytes += _dir_entry(
        "VBA",
        1,
        right=sid_proj,
        child=vba_sid0,
    )

    # VBA children: chained via right siblings
    for idx, (sname, sdata) in enumerate(streams_in_vba):
        right = vba_sids[idx + 1] if idx + 1 < len(vba_sids) else _NOSTREAM
        dir_bytes += _dir_entry(
            sname,
            2,
            right=right,
            start=mini_starts[sname],
            size=len(sdata),
        )

    # PROJECT stream (child of Root via VBA.right)
    dir_bytes += _dir_entry(
        "PROJECT",
        2,
        start=mini_starts["PROJECT"],
        size=len(proj_data),
    )

    # Pad to full directory sectors
    while len(dir_bytes) < n_dir_sec * _SEC:
        dir_bytes += _dir_entry("", 0)

    # ── Build mini FAT sector ─────────────────────────────────────────────
    mf_bytes = bytearray()
    for entry in mini_fat:
        mf_bytes += struct.pack("<I", entry)
    while len(mf_bytes) < n_mf_sec * _SEC:
        mf_bytes += struct.pack("<I", _FREESECT)

    # ── Build FAT sector ──────────────────────────────────────────────────
    fat_bytes = bytearray()
    for entry in fat:
        fat_bytes += struct.pack("<I", entry)
    while len(fat_bytes) < _SEC:
        fat_bytes += struct.pack("<I", _FREESECT)

    # ── CFB header (512 bytes) ────────────────────────────────────────────
    difat = struct.pack("<I", I_FAT) + struct.pack("<I", _FREESECT) * 108

    header = (
        b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"  # Magic
        + b"\x00" * 16  # UID
        + struct.pack("<H", 0x003E)  # MinorVersion
        + struct.pack("<H", 0x0003)  # MajorVersion = 3
        + struct.pack("<H", 0xFFFE)  # ByteOrder (LE)
        + struct.pack("<H", 9)  # SectorSizeExp (2^9=512)
        + struct.pack("<H", 6)  # MiniSectorSizeExp (2^6=64)
        + b"\x00" * 6  # Reserved
        + struct.pack("<I", 0)  # DirectorySectorsNum (0 for v3)
        + struct.pack("<I", 1)  # TotalFATSectors
        + struct.pack("<I", I_DIR)  # FirstDirSectorLoc
        + struct.pack("<I", 0)  # TransactionSigNum
        + struct.pack("<I", _CUT)  # MiniStreamCutoff
        + struct.pack("<I", I_MF)  # FirstMiniFATSectorLoc
        + struct.pack("<I", n_mf_sec)  # TotalMiniFATSectors
        + struct.pack("<I", _ENDOFCHAIN)  # FirstDIFATSectorLoc (none)
        + struct.pack("<I", 0)  # TotalDIFATSectors
        + difat  # DIFAT[0..108]
    )
    assert len(header) == 512, f"Header {len(header)} != 512"

    # ── Assemble CFB ──────────────────────────────────────────────────────
    ms_padded = bytes(ms_blob).ljust(n_ms_sec * _SEC, b"\x00")

    cfb = header + bytes(fat_bytes) + bytes(dir_bytes) + bytes(mf_bytes) + ms_padded
    expected = 512 + N_TOTAL * _SEC
    assert len(cfb) == expected, f"CFB size {len(cfb)} != expected {expected}"

    return cfb


# ---------------------------------------------------------------------------
# XLSM packager
# ---------------------------------------------------------------------------

_WORKBOOK_XML = """\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"""

_SHEET1_XML = """\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
</worksheet>"""

_WORKBOOK_RELS = """\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
    Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vbaProject"
    Target="vbaProject.bin"/>
</Relationships>"""

_RELS = """\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    Target="xl/workbook.xml"/>
</Relationships>"""

_CONTENT_TYPES = """\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels"
    ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml"
    ContentType="application/vnd.ms-excel.sheet.macroEnabled.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml"
    ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/vbaProject.bin"
    ContentType="application/vnd.ms-office.activeX+xml"/>
</Types>"""


def create_xlsm(
    output_path: str | Path,
    modules: List[Tuple[str, str, str]],
) -> None:
    """Create a macro-enabled .xlsm file.

    Args:
        output_path: destination path.
        modules:     [(module_name, stream_name, vba_source_text), ...]
    """
    mods_bytes = [
        (name, stream, src.encode("latin-1")) for name, stream, src in modules
    ]
    vba_bin = _build_vba_project_bin(mods_bytes)

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", _CONTENT_TYPES)
        zf.writestr("_rels/.rels", _RELS)
        zf.writestr("xl/workbook.xml", _WORKBOOK_XML)
        zf.writestr("xl/_rels/workbook.xml.rels", _WORKBOOK_RELS)
        zf.writestr("xl/worksheets/sheet1.xml", _SHEET1_XML)
        zf.writestr("xl/vbaProject.bin", vba_bin)

    Path(output_path).write_bytes(buf.getvalue())
    print(f"Created {output_path}  ({Path(output_path).stat().st_size} bytes)")


# ---------------------------------------------------------------------------
# VBA source fixtures
# ---------------------------------------------------------------------------

VBA_MODULE1 = """\
Attribute VB_Name = "Module1"
Sub HelloWorld()
    MsgBox "Hello, World!"
End Sub

Function AddNumbers(a As Long, b As Long) As Long
    AddNumbers = a + b
End Function

Sub LoopExample()
    Dim i As Integer
    For i = 1 To 10
        Debug.Print i
    Next i
End Sub
"""

VBA_MODULE2 = """\
Attribute VB_Name = "Module2"
Function MultiplyNumbers(x As Double, y As Double) As Double
    MultiplyNumbers = x * y
End Function

Sub StringExample()
    Dim s As String
    s = "Test string"
    MsgBox Len(s)
End Sub
"""

if __name__ == "__main__":
    out = Path(__file__).parent / "test-data" / "Finance_Ledger_VBA.xlsm"
    create_xlsm(
        out,
        modules=[
            ("Module1", "Module1", VBA_MODULE1),
            ("Module2", "Module2", VBA_MODULE2),
        ],
    )
