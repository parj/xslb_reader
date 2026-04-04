"""
reader.py
=========
Pure-Python reader for Excel Binary Workbook (.xlsb) files.
Supports extracting cell formulas, cached/constant cell values, and PivotTable metadata.
No third-party dependencies — only the Python standard library.

Implemented from:
  • [MS-XLSB] Excel (.xlsb) Binary File Format
    https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-xlsb/
  • Ptg token table from MS-XLSB specification §2.5.98.16

Dispatch rules for Ptg bytes
-----------------------------
Tokens 0x01–0x1F  : exact byte match (operators, literals, control tokens).
Tokens 0x20–0x7F  : three class variants per token type:
    REF class   = base_value          (e.g. PtgRef = 0x24)
    VAL class   = base_value + 0x20   (e.g. PtgRef = 0x44)
    ARR class   = base_value + 0x40   (e.g. PtgRef = 0x64)
  Dispatch on  ptg & 0x1F  (strips the two class bits) — but only when
  ptg >= 0x20, because the same low-5-bit values are reused by operators
  in the 0x03–0x14 range.

Usage
-----
    python reader.py workbook.xlsb [sheet_name]

Or programmatically:

    from reader import XlsbWorkbook, col_to_letter
    with XlsbWorkbook("workbook.xlsb") as wb:
        for sheet_name, formulas in wb.iter_formulas():
            for (row, col), formula in sorted(formulas.items()):
                print(f"{sheet_name}!{col_to_letter(col)}{row+1}: {formula}")
"""

import io
import json
import os
import pprint
import re
import struct
import zipfile
from typing import Dict, Iterator, List, Optional, Tuple


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def col_to_letter(col: int) -> str:
    """Convert 0-based column index to Excel letter(s), e.g. 0→A, 26→AA."""
    result = ""
    col += 1
    while col:
        col, rem = divmod(col - 1, 26)
        result = chr(65 + rem) + result
    return result


def cell_ref(row: int, col: int) -> str:
    return f"{col_to_letter(col)}{row + 1}"


def _fmt_num(d: float) -> str:
    """Format a double like Excel (no unnecessary decimals)."""
    if d == int(d) and abs(d) < 1e15:
        return str(int(d))
    return repr(d)


# ---------------------------------------------------------------------------
# BIFF12 record stream reader
# ---------------------------------------------------------------------------

class RecordReader:
    """
    Iterates over BIFF12 variable-length records.

    Each record: [type: 1–2 bytes varint] [size: 1–4 bytes varint] [data: size bytes]

    Yields (record_type: int, data: bytes).
    """

    def __init__(self, data: bytes):
        self._buf = data
        self._pos = 0

    def _varint(self) -> int:
        result = 0
        shift = 0
        for _ in range(4):
            if self._pos >= len(self._buf):
                raise EOFError
            b = self._buf[self._pos]; self._pos += 1
            result |= (b & 0x7F) << shift
            shift += 7
            if not (b & 0x80):
                break
        return result

    def __iter__(self):
        return self

    def __next__(self) -> Tuple[int, bytes]:
        if self._pos >= len(self._buf):
            raise StopIteration
        # record type
        b0 = self._buf[self._pos]; self._pos += 1
        if b0 & 0x80:
            if self._pos >= len(self._buf):
                raise StopIteration
            b1 = self._buf[self._pos]; self._pos += 1
            rec_type = (b0 & 0x7F) | ((b1 & 0x7F) << 7)
        else:
            rec_type = b0
        # record size
        size = self._varint()
        # data
        data = self._buf[self._pos: self._pos + size]
        self._pos += size
        return rec_type, data


# ---------------------------------------------------------------------------
# Record type IDs (subset)
# ---------------------------------------------------------------------------

BRT_ROW_HDR       = 0x0000
BRT_CELL_BLANK    = 0x0001
BRT_CELL_RK       = 0x0002
BRT_CELL_ERROR    = 0x0003
BRT_CELL_BOOL     = 0x0004
BRT_CELL_REAL     = 0x0005
BRT_CELL_ST       = 0x0006
BRT_CELL_ISST     = 0x0007
BRT_FMLA_STRING   = 0x0008
BRT_FMLA_NUM      = 0x0009
BRT_FMLA_BOOL     = 0x000A
BRT_FMLA_ERROR    = 0x000B
BRT_ARR_FMLA      = 0x01AA  # BrtArrFmla
BRT_SHR_FMLA      = 0x01AB  # BrtShrFmla
BRT_NAME          = 0x0027  # BrtName
BRT_SST_ITEM      = 0x0013
BRT_BEGIN_SST     = 0x009F
BRT_END_SST       = 0x00A0
BRT_BEGIN_SHEET   = 0x0081
BRT_END_SHEET     = 0x0082
BRT_BUNDLE_SH     = 0x009C
BRT_BEGIN_SX_VIEW = 0x0118  # BrtBeginSXView
BRT_BEGIN_SX_LOCATION = 0x013A  # BrtBeginSXLocation
BRT_BEGIN_SXVD = 0x011D  # BrtBeginSXVD
BRT_BEGIN_SXVI = 0x011A  # BrtBeginSXVI

# ---------------------------------------------------------------------------
# Ptg exact-byte token constants  (ptg < 0x20)
# ---------------------------------------------------------------------------

PTG_EXP        = 0x01
PTG_TBL        = 0x02
PTG_ADD        = 0x03
PTG_SUB        = 0x04
PTG_MUL        = 0x05
PTG_DIV        = 0x06
PTG_POWER      = 0x07
PTG_CONCAT     = 0x08
PTG_LT         = 0x09
PTG_LE         = 0x0A
PTG_EQ         = 0x0B
PTG_GE         = 0x0C
PTG_GT         = 0x0D
PTG_NE         = 0x0E
PTG_ISECT      = 0x0F
PTG_UNION      = 0x10
PTG_RANGE      = 0x11
PTG_UPLUS      = 0x12
PTG_UMINUS     = 0x13
PTG_PERCENT    = 0x14
PTG_PAREN      = 0x15
PTG_MISSARG    = 0x16
PTG_STR        = 0x17
PTG_LIST       = 0x18   # structured table reference (PtgList/PtgSxName)
PTG_ATTR       = 0x19
PTG_ERR        = 0x1C
PTG_BOOL       = 0x1D
PTG_INT        = 0x1E
PTG_NUM        = 0x1F

# ---------------------------------------------------------------------------
# Class-variant base values  (ptg & 0x1F when ptg >= 0x20)
#
# Each token exists in three variants:
#   REF  = _B_xxx           (e.g. PtgRef  = 0x24, _B_REF = 0x04)
#   VAL  = _B_xxx + 0x20    (e.g. PtgRefV = 0x44)
#   ARR  = _B_xxx + 0x40    (e.g. PtgRefA = 0x64)
# ---------------------------------------------------------------------------

_B_ARRAY     = 0x00   # 0x20/0x40/0x60
_B_FUNC      = 0x01   # 0x21/0x41/0x61
_B_FUNC_VAR  = 0x02   # 0x22/0x42/0x62
_B_NAME      = 0x03   # 0x23/0x43/0x63
_B_REF       = 0x04   # 0x24/0x44/0x64
_B_AREA      = 0x05   # 0x25/0x45/0x65
_B_MEM_AREA  = 0x06   # 0x26/0x46/0x66
_B_MEM_ERR   = 0x07   # 0x27/0x47/0x67
_B_MEM_NOMEM = 0x08   # 0x28/0x48/0x68
_B_MEM_FUNC  = 0x09   # 0x29/0x49/0x69
_B_REF_ERR   = 0x0A   # 0x2A/0x4A/0x6A
_B_AREA_ERR  = 0x0B   # 0x2B/0x4B/0x6B
_B_REF_N     = 0x0C   # 0x2C/0x4C/0x6C
_B_AREA_N    = 0x0D   # 0x2D/0x4D/0x6D
_B_NAME_X    = 0x19   # 0x39/0x59/0x79
_B_REF_3D    = 0x1A   # 0x3A/0x5A/0x7A
_B_AREA_3D   = 0x1B   # 0x3B/0x5B/0x7B
_B_REF_ERR3D = 0x1C   # 0x3C/0x5C/0x7C
_B_AREA_ERR3D= 0x1D   # 0x3D/0x5D/0x7D

# PtgAttr sub-type flags
_ATTR_SEMI   = 0x01
_ATTR_IF     = 0x02
_ATTR_CHOOSE = 0x04
_ATTR_GOTO   = 0x08
_ATTR_SUM    = 0x10
_ATTR_BAXCEL = 0x20
_ATTR_SPACE  = 0x40

# Error code → string
ERR_CODES = {
    0x00: "#NULL!", 0x07: "#DIV/0!", 0x0F: "#VALUE!",
    0x17: "#REF!",  0x1D: "#NAME?",  0x24: "#NUM!",
    0x2A: "#N/A",   0x2B: "#GETTING_DATA",
}


def _rk_to_number(raw: int):
    """Decode RkNumber (2.5.123)."""
    fx100 = raw & 0x1
    f_int = raw & 0x2
    num30 = raw >> 2
    if f_int:
        # signed 30-bit integer
        if num30 & (1 << 29):
            value = num30 - (1 << 30)
        else:
            value = num30
    else:
        bits64 = num30 << 34
        value = struct.unpack("<d", struct.pack("<Q", bits64))[0]
    if fx100:
        value = value / 100
    return value

# ---------------------------------------------------------------------------
# Built-in function table  iftab → name
# (Covers all common functions; from MS-XLSB §2.5.97.4)
# ---------------------------------------------------------------------------

_FUNC: Dict[int, str] = {
    0x0000:"COUNT",    0x0001:"IF",        0x0002:"ISNA",
    0x0003:"ISERROR",  0x0004:"SUM",       0x0005:"AVERAGE",
    0x0006:"MIN",      0x0007:"MAX",       0x0008:"ROW",
    0x0009:"COLUMN",   0x000A:"NA",        0x000B:"NPV",
    0x000C:"STDEV",    0x000D:"DOLLAR",    0x000E:"FIXED",
    0x000F:"SIN",      0x0010:"COS",       0x0011:"TAN",
    0x0012:"ATAN",     0x0013:"PI",        0x0014:"SQRT",
    0x0015:"EXP",      0x0016:"LN",        0x0017:"LOG10",
    0x0018:"ABS",      0x0019:"INT",       0x001A:"SIGN",
    0x001B:"ROUND",    0x001C:"LOOKUP",    0x001D:"INDEX",
    0x001E:"REPT",     0x001F:"MID",       0x0020:"LEN",
    0x0021:"VALUE",    0x0022:"TRUE",      0x0023:"FALSE",
    0x0024:"AND",      0x0025:"OR",        0x0026:"NOT",
    0x0027:"MOD",      0x0028:"DCOUNT",    0x0029:"DSUM",
    0x002A:"DAVERAGE", 0x002B:"DMIN",      0x002C:"DMAX",
    0x002D:"DSTDEV",   0x002E:"VAR",       0x002F:"DVAR",
    0x0030:"TEXT",     0x0031:"LINEST",    0x0032:"TREND",
    0x0033:"LOGEST",   0x0034:"GROWTH",    0x0038:"PV",
    0x0039:"FV",       0x003A:"NPER",      0x003B:"PMT",
    0x003C:"RATE",     0x003D:"MIRR",      0x003E:"IRR",
    0x003F:"RAND",     0x0040:"MATCH",     0x0041:"DATE",
    0x0042:"TIME",     0x0043:"DAY",       0x0044:"MONTH",
    0x0045:"YEAR",     0x0046:"WEEKDAY",   0x0047:"HOUR",
    0x0048:"MINUTE",   0x0049:"SECOND",    0x004A:"NOW",
    0x004B:"AREAS",    0x004C:"ROWS",      0x004D:"COLUMNS",
    0x004E:"OFFSET",   0x0052:"SEARCH",    0x0053:"TRANSPOSE",
    0x0056:"TYPE",     0x0061:"ATAN2",     0x0062:"ASIN",
    0x0063:"ACOS",     0x0064:"CHOOSE",    0x0065:"HLOOKUP",
    0x0066:"VLOOKUP",  0x0069:"ISREF",     0x006D:"LOG",
    0x006F:"CHAR",     0x0070:"LOWER",     0x0071:"UPPER",
    0x0072:"PROPER",   0x0073:"LEFT",      0x0074:"RIGHT",
    0x0075:"EXACT",    0x0076:"TRIM",      0x0077:"REPLACE",
    0x0078:"SUBSTITUTE",0x0079:"CODE",     0x007C:"FIND",
    0x007D:"CELL",     0x007E:"ISERR",     0x007F:"ISTEXT",
    0x0080:"ISNUMBER", 0x0081:"ISBLANK",   0x0082:"T",
    0x0083:"N",        0x008C:"DATEVALUE", 0x008D:"TIMEVALUE",
    0x008E:"SLN",      0x008F:"SYD",       0x0090:"DDB",
    0x0094:"INDIRECT", 0x00BF:"STDEVP",    0x00C0:"VARP",
    0x00C1:"DSTDEVP",  0x00C2:"DVARP",     0x00C3:"TRUNC",
    0x00A2:"CLEAN",    0x00A9:"COUNTA",    0x00C4:"ISLOGICAL",
    0x00C5:"DCOUNTA",  0x00C6:"CLEAN",
    0x00C7:"MDETERM",  0x00C8:"MINVERSE",  0x00C9:"MMULT",
    0x00CB:"IPMT",     0x00CC:"PPMT",      0x00CD:"COUNTA",
    0x00D7:"PRODUCT",  0x00D8:"FACT",      0x00DB:"DPRODUCT",
    0x00DC:"ISNONTEXT",0x00DD:"GETPIVOTDATA",0x00DE:"MEDIAN",
    0x00DF:"SUMPRODUCT",0x00E0:"SINH",     0x00E1:"COSH",
    0x00E2:"TANH",     0x00E3:"ASINH",     0x00E4:"ACOSH",
    0x00E5:"ATANH",    0x00E6:"DGET",      0x00E8:"ASINH",
    0x00E9:"ACOSH",    0x00EA:"ATANH",
    0x00EF:"INFO",
    0x00F2:"DB",       0x00F9:"FREQUENCY", 0x0100:"DAYS360",
    0x0101:"TODAY",    0x0102:"VDB",       0x0109:"ERRORTYPE",
    0x010B:"WORKDAY",  0x010C:"NETWORKDAYS",0x010D:"WEEKNUM",
    0x010E:"FLOOR",    0x010F:"CEILING",   0x0110:"ISEVEN",
    0x0111:"ISODD",    0x0112:"MROUND",    0x0113:"QUOTIENT",
    0x0114:"GCD",      0x0115:"LCM",       0x0116:"MULTINOMIAL",
    0x0117:"COMBIN",   0x0118:"PERMUT",    0x011A:"CONFIDENCE",
    0x011C:"EVEN",     0x011D:"EXPONDIST", 0x011E:"FDIST",
    0x011F:"FINV",     0x0120:"FISHER",    0x0121:"FISHERINV",
    0x0123:"GAMMADIST",0x0124:"GAMMAINV",  0x0126:"HYPGEOMDIST",
    0x0127:"LOGINV",   0x0128:"LOGNORMDIST",0x0129:"NEGBINOMDIST",
    0x012A:"NORMDIST", 0x012B:"NORMSDIST", 0x012C:"NORMINV",
    0x012D:"NORMSINV", 0x012E:"STANDARDIZE",0x012F:"ODD",
    0x0131:"POISSON",  0x0132:"TDIST",     0x0133:"WEIBULL",
    0x0134:"SUMXMY2",  0x0135:"SUMX2MY2",  0x0136:"SUMX2PY2",
    0x0137:"CHITEST",  0x0138:"CORREL",    0x0139:"COVAR",
    0x013A:"FORECAST", 0x013B:"FTEST",     0x013C:"INTERCEPT",
    0x013D:"PEARSON",  0x013E:"RSQ",       0x013F:"STEYX",
    0x0140:"SLOPE",    0x0141:"TTEST",     0x0142:"PROB",
    0x0143:"DEVSQ",    0x0144:"GEOMEAN",   0x0145:"HARMEAN",
    0x0146:"SUMSQ",    0x0147:"KURT",      0x0148:"SKEW",
    0x0149:"ZTEST",    0x014A:"LARGE",     0x014B:"SMALL",
    0x014C:"QUARTILE", 0x014D:"PERCENTILE",0x014E:"PERCENTRANK",
    0x014F:"MODE",     0x0150:"TRIMMEAN",  0x0151:"TINV",
    0x0157:"CONCATENATE",0x0158:"POWER",   0x015A:"RADIANS",
    0x015B:"DEGREES",  0x015C:"SUBTOTAL",  0x015D:"SUMIF",
    0x015E:"COUNTIF",  0x015F:"COUNTBLANK",0x0162:"ROMAN",
    0x0163:"HYPERLINK",0x0164:"AVERAGEA",  0x0165:"MAXA",
    0x0166:"MINA",     0x0167:"STDEVPA",   0x0168:"VARPA",
    0x0169:"STDEVA",   0x016A:"VARA",      0x0176:"RTD",
    0x017B:"HEX2BIN",  0x017C:"HEX2DEC",  0x017D:"HEX2OCT",
    0x017E:"DEC2BIN",  0x017F:"DEC2HEX",  0x0180:"DEC2OCT",
    0x0181:"OCT2BIN",  0x0182:"OCT2DEC",  0x0183:"OCT2HEX",
    0x0184:"BIN2DEC",  0x0185:"BIN2OCT",  0x0186:"BIN2HEX",
    0x0187:"IMSUB",    0x0188:"IMDIV",     0x0189:"IMPOWER",
    0x018A:"IMABS",    0x018B:"IMSQRT",    0x018C:"IMLN",
    0x018D:"IMLOG2",   0x018E:"IMLOG10",   0x018F:"IMSIN",
    0x0190:"IMCOS",    0x0191:"IMEXP",     0x0192:"IMARGUMENT",
    0x0193:"IMCONJUGATE",0x0194:"IMAGINARY",0x0195:"IMREAL",
    0x0196:"COMPLEX",  0x0197:"IMSUM",     0x0198:"IMPRODUCT",
    0x0199:"SERIESSUM",0x019A:"FACTDOUBLE",0x019B:"SQRTPI",
    0x019D:"DELTA",    0x019E:"GESTEP",    0x01A2:"ERF",
    0x01A3:"ERFC",     0x01A4:"BESSELJ",   0x01A5:"BESSELK",
    0x01A6:"BESSELY",  0x01A7:"BESSELI",   0x01A8:"XIRR",
    0x01A9:"XNPV",     0x01AA:"PRICEMAT",  0x01AB:"YIELDMAT",
    0x01AC:"INTRATE",  0x01AD:"RECEIVED",  0x01AE:"DISC",
    0x01AF:"PRICEDISC",0x01B0:"YIELDDISC", 0x01B1:"TBILLEQ",
    0x01B2:"TBILLPRICE",0x01B3:"TBILLYIELD",0x01B4:"PRICE",
    0x01B5:"YIELD",    0x01B6:"DOLLARDE",  0x01B7:"DOLLARFR",
    0x01B8:"NOMINAL",  0x01B9:"EFFECT",    0x01BA:"CUMPRINC",
    0x01BB:"CUMIPMT",  0x01BC:"EDATE",     0x01BD:"EOMONTH",
    0x01BE:"YEARFRAC", 0x01BF:"COUPDAYBS", 0x01C0:"COUPDAYS",
    0x01C1:"COUPDAYSNC",0x01C2:"COUPNCD",  0x01C3:"COUPNUM",
    0x01C4:"COUPPCD",  0x01C5:"DURATION",  0x01C6:"MDURATION",
    0x01C7:"ODDLPRICE",0x01C8:"ODDLYIELD", 0x01C9:"ODDFPRICE",
    0x01CA:"ODDFYIELD",0x01CB:"RANDBETWEEN",0x01CC:"WEEKNUM",
    0x01CD:"AMORDEGRC",0x01CE:"AMORLINC",  0x01CF:"CONVERT",
    0x01D0:"ACCRINT",  0x01D1:"ACCRINTM",  0x01D2:"WORKDAY",
    0x01D3:"NETWORKDAYS",0x01D4:"GCD",     0x01D5:"MULTINOMIAL",
    0x01D6:"LCM",      0x01D7:"FVSCHEDULE",0x01D8:"CUBESETCOUNT",
    0x01D9:"CUBESET",  0x01DA:"IFERROR",   0x01DB:"COUNTIFS",
    0x01DC:"SUMIFS",   0x01DD:"AVERAGEIF", 0x01DE:"AVERAGEIFS",
    0x01DF:"AGGREGATE",0x01E3:"AVERAGEIF", 0x0213:"IFNA",
    0x0214:"IFS",
    0x0215:"SWITCH",   0x0216:"MAXIFS",    0x0217:"MINIFS",
    0x0241:"XLOOKUP",  0x0244:"XMATCH",    0x0246:"FILTER",
    0x0247:"SORT",     0x0248:"SORTBY",    0x0249:"UNIQUE",
    0x024A:"SEQUENCE", 0x024B:"RANDARRAY", 0x024C:"LET",
}


def _load_ftab_from_local_spec() -> Dict[int, str]:
    """
    Load Ftab function id->name mapping from local xslb-file-format.md if present.
    This avoids drift in manually maintained _FUNC tables.
    """
    md_path = os.path.join(os.path.dirname(__file__), "xslb-file-format.md")
    if not os.path.exists(md_path):
        return {}
    try:
        with open(md_path, "r", encoding="utf-8", errors="replace") as f:
            text = f.read()
    except OSError:
        return {}

    m_start = re.search(r"(?m)^2\.5\.98\.10\s+Ftab\s*$", text)
    if not m_start:
        return {}
    m_end = re.search(r"(?m)^2\.5\.98\.11\s+", text[m_start.end():])
    section = (text[m_start.end(): m_start.end() + m_end.start()]
               if m_end else text[m_start.end():])

    out: Dict[int, str] = {}
    for hx, name in re.findall(
        r"(?m)^\s*0x([0-9A-Fa-f]{4})\s+([A-Za-z0-9_.\-]+)\b", section
    ):
        fid = int(hx, 16)
        # Normalize known special row.
        if fid == 0x00FF:
            out[fid] = "UDF"
        else:
            out[fid] = name
    return out


# Prefer authoritative IDs from local spec when available.
_FUNC.update(_load_ftab_from_local_spec())


def _func_name(iftab: int) -> str:
    return _FUNC.get(iftab, f"_func{iftab:#06x}")


# Fixed arity for PtgFunc tokens
_FIXED_ARITY: Dict[int, int] = {
    0x0003:1, 0x000A:0, 0x000F:1, 0x0010:1, 0x0011:1, 0x0012:1,
    0x0013:0,  # PI
    0x0014:1, 0x0015:1, 0x0016:1, 0x0017:1, 0x0018:1, 0x0019:1,
    0x001A:1, 0x001B:2, 0x0022:0, 0x0023:0, 0x0026:1, 0x0027:2,
    0x003F:0, 0x0041:3, 0x0042:3, 0x0043:1, 0x0044:1, 0x0045:1,
    0x0046:1, 0x0047:1, 0x0048:1, 0x0049:1, 0x004A:0, 0x004B:1,
    0x004C:1, 0x004D:1, 0x0061:2, 0x0062:1, 0x0063:1, 0x006D:2,
    0x006F:1, 0x0070:1, 0x0071:1, 0x0072:1, 0x0075:2, 0x0076:1,
    0x007E:1, 0x007F:1, 0x0080:1, 0x0081:1, 0x0082:1, 0x0083:1,
    0x00C3:1, 0x00C4:1, 0x00C6:1, 0x00D8:1, 0x0100:2, 0x0101:0,
    0x010E:2, 0x010F:2, 0x011C:1, 0x012F:1, 0x0157:2, 0x0158:2,
    0x015A:1, 0x015B:1,
}


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _sheet_prefix(sheet_names: List[str], ixti: int) -> str:
    if 0 <= ixti < len(sheet_names):
        name = sheet_names[ixti]
        if any(c in name for c in " ![]'"):
            name = f"'{name}'"
        return f"{name}!"
    return f"[sheet{ixti}]!"


def _read_rgce_loc_raw(buf: io.BytesIO) -> Tuple[int, int]:
    """Read raw RgceLoc/RgceLocRel fields: row (u32) + col_flags (u16)."""
    row_raw = struct.unpack("<I", buf.read(4))[0]
    col_raw = struct.unpack("<H", buf.read(2))[0]
    return row_raw, col_raw


def _read_rgce_area_raw(buf: io.BytesIO) -> Tuple[int, int, int, int]:
    """Read raw RgceArea/RgceAreaRel fields."""
    row_first = struct.unpack("<I", buf.read(4))[0]
    row_last = struct.unpack("<I", buf.read(4))[0]
    col_first = struct.unpack("<H", buf.read(2))[0]
    col_last = struct.unpack("<H", buf.read(2))[0]
    return row_first, row_last, col_first, col_last


def _resolve_loc(
    row_raw: int,
    col_raw: int,
    base_row: Optional[int] = None,
    base_col: Optional[int] = None,
    apply_relative_offsets: bool = False,
) -> Tuple[int, int, bool, bool]:
    """
    Resolve RgceLoc/RgceLocRel into a concrete row/col.
    Returns (row, col, row_rel, col_rel) where row/col are resolved coordinates.
    """
    col = col_raw & 0x3FFF
    col_rel = bool(col_raw & 0x4000)
    row_rel = bool(col_raw & 0x8000)
    row = row_raw

    if apply_relative_offsets and row_rel and base_row is not None:
        # 20-bit signed row offset, modulo 1,048,576 rows.
        off = row_raw & 0x000FFFFF
        if off >= 0x00080000:
            off -= 0x00100000
        row = (base_row + off) % 0x00100000

    if apply_relative_offsets and col_rel and base_col is not None:
        # 14-bit signed col offset, modulo 16,384 columns.
        off = col if col < 0x2000 else (col - 0x4000)
        col = (base_col + off) % 0x4000

    return row, col, row_rel, col_rel


def _loc_str(row: int, col: int, row_rel: bool, col_rel: bool) -> str:
    c = ("" if col_rel else "$") + col_to_letter(col)
    r = ("" if row_rel else "$") + str(row + 1)
    return c + r


# ---------------------------------------------------------------------------
# Ptg decompiler
# ---------------------------------------------------------------------------

class _Decompiler:
    """
    Stack-based RPN → infix formula decompiler.

    Two-phase dispatch:
      ptg < 0x20  → exact-byte (operators, literals, control tokens)
      ptg >= 0x20 → operand/function tokens, dispatched on ptg & 0x1F
    """

    def __init__(self, rgce: bytes,
                 sheet_names: Optional[List[str]] = None,
                 defined_names: Optional[Dict[int, str]] = None,
                 base_row: Optional[int] = None,
                 base_col: Optional[int] = None,
                 rgcb: bytes = b""):
        self._buf = io.BytesIO(rgce)
        self._end = len(rgce)
        self._sheets = sheet_names or []
        self._dnames = defined_names or {}
        self._base_row = base_row
        self._base_col = base_col
        self._rgcb = rgcb

    def _u8(self)  -> int:   return struct.unpack("<B", self._buf.read(1))[0]
    def _u16(self) -> int:   return struct.unpack("<H", self._buf.read(2))[0]
    def _i32(self) -> int:   return struct.unpack("<i", self._buf.read(4))[0]
    def _u32(self) -> int:   return struct.unpack("<I", self._buf.read(4))[0]
    def _f64(self) -> float: return struct.unpack("<d", self._buf.read(8))[0]

    def _wstr16(self) -> str:
        cch = self._u16()
        return self._buf.read(cch * 2).decode("utf-16-le", errors="replace")

    @staticmethod
    def _pop(stack: List[str]) -> str:
        return stack.pop() if stack else "?"

    @staticmethod
    def _pop_n(stack: List[str], n: int) -> List[str]:
        args: List[str] = []
        for _ in range(n):
            args.insert(0, stack.pop() if stack else "?")
        return args

    def decompile(self) -> str:
        stack: List[str] = []

        while self._buf.tell() < self._end:
            ptg = self._u8()

            # ================================================================
            # Phase 1: exact-byte tokens  (ptg < 0x20)
            # ================================================================
            if ptg < 0x20:
                _BINOP = {
                    PTG_ADD:"+", PTG_SUB:"-", PTG_MUL:"*", PTG_DIV:"/",
                    PTG_POWER:"^", PTG_CONCAT:"&",
                    PTG_LT:"<",  PTG_LE:"<=", PTG_EQ:"=",
                    PTG_GE:">=", PTG_GT:">",  PTG_NE:"<>",
                    PTG_RANGE:":",
                }
                if ptg in _BINOP:
                    right = self._pop(stack); left = self._pop(stack)
                    stack.append(f"{left}{_BINOP[ptg]}{right}")

                elif ptg == PTG_ISECT:
                    right = self._pop(stack); left = self._pop(stack)
                    stack.append(f"{left} {right}")

                elif ptg == PTG_UNION:
                    right = self._pop(stack); left = self._pop(stack)
                    stack.append(f"({left},{right})")

                elif ptg == PTG_UPLUS:
                    stack.append(f"+{self._pop(stack)}")

                elif ptg == PTG_UMINUS:
                    stack.append(f"-{self._pop(stack)}")

                elif ptg == PTG_PERCENT:
                    stack.append(f"{self._pop(stack)}%")

                elif ptg == PTG_PAREN:
                    stack.append(f"({self._pop(stack)})")

                elif ptg == PTG_MISSARG:
                    stack.append("")

                elif ptg == PTG_STR:
                    stack.append(f'"{self._wstr16()}"')

                elif ptg == PTG_LIST:
                    # Either PtgList (eptg=0x19) or PtgSxName (eptg=0x1D).
                    eptg = self._u8()
                    if eptg == 0x19:
                        self._u16()  # ixti
                        f1 = self._u8()
                        f2 = self._u8()
                        columns = f1 & 0x03
                        invalid = bool(f2 & 0x10)
                        nonresident = bool(f2 & 0x20)
                        list_index = self._u32()
                        col_first = self._u16()
                        col_last = self._u16()
                        if invalid:
                            stack.append("#REF!")
                        elif nonresident:
                            stack.append(f"Table{list_index}[?]")
                        elif columns == 0:
                            stack.append(f"Table{list_index}")
                        elif columns == 1:
                            stack.append(f"Table{list_index}[col{col_first + 1}]")
                        else:
                            stack.append(
                                f"Table{list_index}[col{col_first + 1}:col{col_last + 1}]"
                            )
                    elif eptg == 0x1D:
                        sx_index = self._u32()
                        stack.append(f"PivotName{sx_index}")
                    else:
                        stack.append(f"?ptglist{eptg:#04x}")

                elif ptg == PTG_ATTR:
                    atype = self._u8()
                    if atype & _ATTR_SUM:
                        self._u16()              # skip size word
                        stack.append(f"SUM({self._pop(stack)})")
                    elif atype & _ATTR_CHOOSE:
                        ncases = self._u16()
                        for _ in range(ncases + 1):
                            self._u16()          # offset array
                    elif atype & _ATTR_SPACE:
                        self._u8(); self._u8()   # space type + count
                    else:
                        self._u16()              # GOTO/IF/SEMI/BAXCEL offset

                elif ptg == PTG_ERR:
                    stack.append(ERR_CODES.get(self._u8(), "#ERR"))

                elif ptg == PTG_BOOL:
                    stack.append("TRUE" if self._u8() else "FALSE")

                elif ptg == PTG_INT:
                    stack.append(str(self._u16()))

                elif ptg == PTG_NUM:
                    stack.append(_fmt_num(self._f64()))

                elif ptg == PTG_EXP:
                    # XLSB PtgExp stores row only; column lives in rgcb as PtgExtraCol.
                    row = self._u32()
                    col = (struct.unpack("<I", self._rgcb[:4])[0]
                           if len(self._rgcb) >= 4 else 0)
                    stack.append(f"{{array@{cell_ref(row,col)}}}")

                elif ptg == PTG_TBL:
                    row = self._i32(); col = self._u16()
                    stack.append(f"TABLE({cell_ref(row,col)})")

                else:
                    stack.append(f"?ptg{ptg:#04x}")

            # ================================================================
            # Phase 2: class-variant tokens  (ptg >= 0x20)
            # Dispatch on ptg & 0x1F
            # ================================================================
            else:
                base = ptg & 0x1F

                if base == _B_ARRAY:
                    self._buf.read(7)          # 7 placeholder bytes
                    stack.append("{array}")

                elif base == _B_FUNC:
                    iftab = self._u16()
                    name  = _func_name(iftab)
                    arity = _FIXED_ARITY.get(iftab)
                    args  = (self._pop_n(stack, arity) if arity is not None
                             else [self._pop(stack)])
                    stack.append(f"{name}({','.join(args)})")

                elif base == _B_FUNC_VAR:
                    cparams_raw = self._u8()
                    is_ce  = bool(cparams_raw & 0x80)
                    cparams = cparams_raw & 0x7F
                    iftab = self._u16()
                    args = self._pop_n(stack, cparams)
                    if iftab == 0x00FF and args:
                        # Future function/UDF call pattern:
                        # first arg is function name from a PtgName/PtgNameX token.
                        raw_name = args[0].strip()
                        call_name = raw_name.replace("_xlfn.", "")
                        if call_name and not call_name.startswith("_"):
                            call_name = f"@{call_name}"
                        stack.append(f"{call_name}({','.join(args[1:])})")
                    else:
                        name  = (_func_name(iftab) if not is_ce
                                 else f"_xlfn.{_func_name(iftab)}")
                        stack.append(f"{name}({','.join(args)})")

                elif base == _B_NAME:
                    nameindex = self._u32()
                    stack.append(self._dnames.get(nameindex,
                                                  f"name{nameindex}"))

                elif base == _B_REF:
                    row_raw, col_raw = _read_rgce_loc_raw(self._buf)
                    row, col, rr, cr = _resolve_loc(
                        row_raw, col_raw, self._base_row, self._base_col,
                        apply_relative_offsets=False
                    )
                    stack.append(_loc_str(row, col, rr, cr))

                elif base == _B_AREA:
                    r1_raw, r2_raw, c1_raw, c2_raw = _read_rgce_area_raw(self._buf)
                    r1,c1,r1r,c1r = _resolve_loc(
                        r1_raw, c1_raw, self._base_row, self._base_col,
                        apply_relative_offsets=False
                    )
                    r2,c2,r2r,c2r = _resolve_loc(
                        r2_raw, c2_raw, self._base_row, self._base_col,
                        apply_relative_offsets=False
                    )
                    stack.append(
                        f"{_loc_str(r1,c1,r1r,c1r)}:{_loc_str(r2,c2,r2r,c2r)}"
                    )

                elif base in (_B_MEM_AREA, _B_MEM_ERR,
                              _B_MEM_NOMEM, _B_MEM_FUNC):
                    self._buf.read(4); self._u16()  # 4 reserved + cce

                elif base == _B_REF_ERR:
                    self._buf.read(6)
                    stack.append("#REF!")

                elif base == _B_AREA_ERR:
                    self._buf.read(12)
                    stack.append("#REF!:#REF!")

                elif base == _B_REF_N:
                    row_raw, col_raw = _read_rgce_loc_raw(self._buf)
                    row, col, rr, cr = _resolve_loc(
                        row_raw, col_raw, self._base_row, self._base_col,
                        apply_relative_offsets=True
                    )
                    stack.append(_loc_str(row, col, rr, cr))

                elif base == _B_AREA_N:
                    r1_raw, r2_raw, c1_raw, c2_raw = _read_rgce_area_raw(self._buf)
                    r1,c1,r1r,c1r = _resolve_loc(
                        r1_raw, c1_raw, self._base_row, self._base_col,
                        apply_relative_offsets=True
                    )
                    r2,c2,r2r,c2r = _resolve_loc(
                        r2_raw, c2_raw, self._base_row, self._base_col,
                        apply_relative_offsets=True
                    )
                    stack.append(
                        f"{_loc_str(r1,c1,r1r,c1r)}:{_loc_str(r2,c2,r2r,c2r)}"
                    )

                elif base == _B_NAME_X:
                    self._u16()              # ixti
                    nameindex = self._u32()
                    stack.append(self._dnames.get(nameindex,
                                                  f"name{nameindex}"))

                elif base == _B_REF_3D:
                    ixti = self._u16()
                    row_raw, col_raw = _read_rgce_loc_raw(self._buf)
                    row, col, rr, cr = _resolve_loc(
                        row_raw, col_raw, self._base_row, self._base_col,
                        apply_relative_offsets=False
                    )
                    pfx = _sheet_prefix(self._sheets, ixti)
                    stack.append(f"{pfx}{_loc_str(row, col, rr, cr)}")

                elif base == _B_AREA_3D:
                    ixti = self._u16()
                    r1_raw, r2_raw, c1_raw, c2_raw = _read_rgce_area_raw(self._buf)
                    r1,c1,r1r,c1r = _resolve_loc(
                        r1_raw, c1_raw, self._base_row, self._base_col,
                        apply_relative_offsets=False
                    )
                    r2,c2,r2r,c2r = _resolve_loc(
                        r2_raw, c2_raw, self._base_row, self._base_col,
                        apply_relative_offsets=False
                    )
                    pfx = _sheet_prefix(self._sheets, ixti)
                    stack.append(
                        f"{pfx}"
                        f"{_loc_str(r1,c1,r1r,c1r)}:{_loc_str(r2,c2,r2r,c2r)}"
                    )

                elif base == _B_REF_ERR3D:
                    self._buf.read(2 + 6)
                    stack.append("#REF!")

                elif base == _B_AREA_ERR3D:
                    self._buf.read(2 + 12)
                    stack.append("#REF!")

                else:
                    stack.append(f"?ptg{ptg:#04x}")
                    break

        return stack[-1] if stack else ""


# ---------------------------------------------------------------------------
# Shared-string table
# ---------------------------------------------------------------------------

def _read_sst(data: bytes) -> List[str]:
    strings: List[str] = []
    for rec_type, payload in RecordReader(data):
        if rec_type == BRT_SST_ITEM:
            if len(payload) < 5:
                strings.append(""); continue
            buf = io.BytesIO(payload)
            buf.read(1)                          # flags byte
            cch = struct.unpack("<I", buf.read(4))[0]
            strings.append(
                buf.read(cch * 2).decode("utf-16-le", errors="replace")
            )
    return strings


def _read_xlwide_from(buf: io.BytesIO) -> str:
    cch_raw = buf.read(4)
    if len(cch_raw) < 4:
        return ""
    cch = struct.unpack("<I", cch_raw)[0]
    raw = buf.read(cch * 2)
    if len(raw) < cch * 2:
        return ""
    return raw.decode("utf-16-le", errors="replace")


# ---------------------------------------------------------------------------
# Workbook part  (xl/workbook.bin)
# ---------------------------------------------------------------------------

def _read_workbook(data: bytes) -> List[Tuple[str, str]]:
    """Return [(sheet_name, rel_id), …] in tab order."""
    sheets: List[Tuple[str, str]] = []
    for rec_type, payload in RecordReader(data):
        if rec_type == BRT_BUNDLE_SH and len(payload) >= 8:
            buf = io.BytesIO(payload)
            buf.read(8)                              # hsState + iTabID
            cch_rel = struct.unpack("<I", buf.read(4))[0]
            rel_id  = buf.read(cch_rel * 2).decode("utf-16-le", errors="replace")
            cch_nam = struct.unpack("<I", buf.read(4))[0]
            name    = buf.read(cch_nam * 2).decode("utf-16-le", errors="replace")
            sheets.append((name, rel_id))
    return sheets


def _read_defined_names(data: bytes) -> Dict[int, str]:
    """
    Return {1-based nameindex -> defined-name text} from BrtName records.
    """
    out: Dict[int, str] = {}
    idx = 0
    for rec_type, payload in RecordReader(data):
        if rec_type != BRT_NAME or len(payload) < 13:
            continue
        idx += 1
        buf = io.BytesIO(payload)
        buf.read(4)  # flags
        buf.read(1)  # chKey
        buf.read(4)  # itab
        cch_raw = buf.read(4)
        if len(cch_raw) < 4:
            continue
        cch = struct.unpack("<I", cch_raw)[0]
        rgch = buf.read(cch * 2)
        if len(rgch) < cch * 2:
            continue
        out[idx] = rgch.decode("utf-16-le", errors="replace")
    return out


# ---------------------------------------------------------------------------
# Relationships
# ---------------------------------------------------------------------------

def _read_rels(xml_data: bytes) -> Dict[str, str]:
    text = xml_data.decode("utf-8", errors="replace")
    return {
        m.group(1): m.group(2)
        for m in re.finditer(
            r'<Relationship[^>]+Id="([^"]+)"[^>]+Target="([^"]+)"', text
        )
    }


def _resolve_rel_target(base_part: str, target: str) -> str:
    """
    Resolve an OPC relationship target relative to the source part path.
    """
    base_dir = os.path.dirname(base_part.lstrip("/"))
    joined = os.path.normpath(os.path.join(base_dir, target))
    return joined.replace("\\", "/")


# ---------------------------------------------------------------------------
# Worksheet parser
# ---------------------------------------------------------------------------

_FMLA_RECS = {BRT_FMLA_STRING, BRT_FMLA_NUM, BRT_FMLA_BOOL, BRT_FMLA_ERROR}


def _read_cell_parsed_formula(
    buf: io.BytesIO,
) -> Optional[Tuple[bytes, bytes]]:
    """
    Read CellParsedFormula/SharedParsedFormula payload from current position.
    Returns (rgce, rgcb) or None if malformed.
    """
    cce_raw = buf.read(4)
    if len(cce_raw) < 4:
        return None
    cce = struct.unpack("<I", cce_raw)[0]
    if cce == 0 or cce > 16384:
        return None
    rgce = buf.read(cce)
    if len(rgce) < cce:
        return None

    cb_raw = buf.read(4)
    if len(cb_raw) < 4:
        return None
    cb = struct.unpack("<I", cb_raw)[0]
    rgcb = buf.read(cb) if cb else b""
    if len(rgcb) < cb:
        return None
    return rgce, rgcb


def _parse_worksheet(
    data: bytes,
    sheet_names: List[str],
    sst: List[str],
    defined_names: Optional[Dict[int, str]] = None,
) -> Dict[Tuple[int, int], str]:
    """
    Parse a worksheet .bin part.
    Returns {(row, col): "=formula_string"}.  Row and col are 0-based.
    """
    formulas: Dict[Tuple[int, int], str] = {}
    current_row = 0
    dn = defined_names or {}
    shared_formulas: List[Tuple[int, int, int, int, bytes, bytes]] = []
    pending_exp: Dict[Tuple[int, int], Tuple[bytes, bytes]] = {}
    last_formula_cell: Optional[Tuple[int, int]] = None

    for rec_type, payload in RecordReader(data):
        if rec_type == BRT_ROW_HDR:
            if len(payload) >= 4:
                current_row = struct.unpack_from("<I", payload, 0)[0]
            last_formula_cell = None

        elif rec_type in _FMLA_RECS:
            if len(payload) < 10:
                continue
            buf = io.BytesIO(payload)

            # col (4 bytes) + style ref (4 bytes)
            col = struct.unpack("<I", buf.read(4))[0]
            buf.read(4)

            # Skip cached result
            if rec_type == BRT_FMLA_STRING:
                cch = struct.unpack("<I", buf.read(4))[0]
                buf.read(cch * 2)
                
            elif rec_type == BRT_FMLA_NUM:
                buf.read(8)
            elif rec_type in (BRT_FMLA_BOOL, BRT_FMLA_ERROR):
                buf.read(1)
            
            # 2-byte formula flags
            buf.read(2)

            parsed = _read_cell_parsed_formula(buf)
            if parsed is None:
                continue
            rgce, rgcb = parsed

            try:
                if rgce[:1] == bytes([PTG_EXP]):
                    # Shared/array formula reference; resolve after BrtShrFmla/BrtArrFmla.
                    pending_exp[(current_row, col)] = (rgce, rgcb)
                    formula = "={shared_formula}"
                else:
                    formula = "=" + _Decompiler(
                        rgce,
                        sheet_names=sheet_names,
                        defined_names=dn,
                        base_row=current_row,
                        base_col=col,
                        rgcb=rgcb,
                    ).decompile()
            except Exception as exc:
                formula = f"=<parse_error:{exc}>"

            formulas[(current_row, col)] = formula
            last_formula_cell = (current_row, col)

        elif rec_type in (BRT_SHR_FMLA, BRT_ARR_FMLA):
            # Shared/array formula body for the immediately preceding formula cell.
            if len(payload) < 20:
                continue
            buf = io.BytesIO(payload)
            try:
                rw_first, rw_last, col_first, col_last = struct.unpack("<IIII", buf.read(16))
            except struct.error:
                continue
            parsed = _read_cell_parsed_formula(buf)
            if parsed is None:
                continue
            rgce, rgcb = parsed
            shared_formulas.append((rw_first, rw_last, col_first, col_last, rgce, rgcb))

            # Resolve anchor cell immediately when available.
            if last_formula_cell is not None:
                arow, acol = last_formula_cell
                if rw_first <= arow <= rw_last and col_first <= acol <= col_last:
                    try:
                        formulas[(arow, acol)] = "=" + _Decompiler(
                            rgce,
                            sheet_names=sheet_names,
                            defined_names=dn,
                            base_row=arow,
                            base_col=acol,
                            rgcb=rgcb,
                        ).decompile()
                    except Exception as exc:
                        formulas[(arow, acol)] = f"=<parse_error:{exc}>"

    # Resolve shared-formula references (PtgExp) after full pass.
    for (row, col), (_rgce, _rgcb) in pending_exp.items():
        matched = None
        for rw_first, rw_last, col_first, col_last, s_rgce, s_rgcb in shared_formulas:
            if rw_first <= row <= rw_last and col_first <= col <= col_last:
                matched = (s_rgce, s_rgcb)
                break
        if matched is None:
            continue
        s_rgce, s_rgcb = matched
        try:
            formulas[(row, col)] = "=" + _Decompiler(
                s_rgce,
                sheet_names=sheet_names,
                defined_names=dn,
                base_row=row,
                base_col=col,
                rgcb=s_rgcb,
            ).decompile()
        except Exception as exc:
            formulas[(row, col)] = f"=<parse_error:{exc}>"

    return formulas


def _parse_worksheet_values(
    data: bytes,
    sst: List[str],
) -> Dict[Tuple[int, int], object]:
    """
    Parse worksheet cached/constant cell values.
    Returns {(row, col): value}. Row/col are 0-based.
    """
    values: Dict[Tuple[int, int], object] = {}
    current_row = 0

    for rec_type, payload in RecordReader(data):
        if rec_type == BRT_ROW_HDR:
            if len(payload) >= 4:
                current_row = struct.unpack_from("<I", payload, 0)[0]
            continue

        # Cell records
        if rec_type == BRT_CELL_ISST and len(payload) >= 12:
            col = struct.unpack_from("<I", payload, 0)[0]
            isst = struct.unpack_from("<I", payload, 8)[0]
            values[(current_row, col)] = sst[isst] if 0 <= isst < len(sst) else ""

        elif rec_type == BRT_CELL_ST and len(payload) >= 8:
            col = struct.unpack_from("<I", payload, 0)[0]
            buf = io.BytesIO(payload)
            buf.read(8)  # Cell
            values[(current_row, col)] = _read_xlwide_from(buf)

        elif rec_type == BRT_CELL_REAL and len(payload) >= 16:
            col = struct.unpack_from("<I", payload, 0)[0]
            values[(current_row, col)] = struct.unpack_from("<d", payload, 8)[0]

        elif rec_type == BRT_CELL_RK and len(payload) >= 12:
            col = struct.unpack_from("<I", payload, 0)[0]
            rk = struct.unpack_from("<I", payload, 8)[0]
            values[(current_row, col)] = _rk_to_number(rk)

        elif rec_type == BRT_CELL_BOOL and len(payload) >= 9:
            col = struct.unpack_from("<I", payload, 0)[0]
            values[(current_row, col)] = bool(payload[8])

        elif rec_type == BRT_CELL_ERROR and len(payload) >= 9:
            col = struct.unpack_from("<I", payload, 0)[0]
            values[(current_row, col)] = ERR_CODES.get(payload[8], "#ERR")

        # Formula records (use cached evaluation result)
        elif rec_type in _FMLA_RECS and len(payload) >= 10:
            col = struct.unpack_from("<I", payload, 0)[0]
            buf = io.BytesIO(payload)
            buf.read(8)  # Cell
            if rec_type == BRT_FMLA_STRING:
                values[(current_row, col)] = _read_xlwide_from(buf)
            elif rec_type == BRT_FMLA_NUM:
                raw = buf.read(8)
                if len(raw) == 8:
                    values[(current_row, col)] = struct.unpack("<d", raw)[0]
            elif rec_type == BRT_FMLA_BOOL:
                b = buf.read(1)
                if b:
                    values[(current_row, col)] = bool(b[0])
            elif rec_type == BRT_FMLA_ERROR:
                b = buf.read(1)
                if b:
                    values[(current_row, col)] = ERR_CODES.get(b[0], "#ERR")

    return values


def _parse_pivot_table_part(data: bytes) -> Dict[str, object]:
    """
    Parse key metadata from a PivotTable part.
    """
    meta: Dict[str, object] = {
        "name": None,
        "cache_id": None,
        "data_caption": None,
        "location": None,
        "pivot_fields": 0,
        "pivot_items": 0,
    }
    for rec_type, payload in RecordReader(data):
        if rec_type == BRT_BEGIN_SX_VIEW and len(payload) >= 32:
            # Fixed part from BrtBeginSXView example (3.8.26)
            id_cache = struct.unpack_from("<I", payload, 28)[0]
            meta["cache_id"] = id_cache
            buf = io.BytesIO(payload)
            buf.seek(32)
            meta["name"] = _read_xlwide_from(buf)
            # fDisplayData is required to be 1 in spec for this record.
            data_caption = _read_xlwide_from(buf)
            if data_caption:
                meta["data_caption"] = data_caption

        elif rec_type == BRT_BEGIN_SX_LOCATION and len(payload) >= 36:
            rw_first, rw_last, col_first, col_last = struct.unpack_from("<IIII", payload, 0)
            rw_first_head, rw_first_data, col_first_data, crw_page, ccol_page = (
                struct.unpack_from("<IIIII", payload, 16)
            )
            meta["location"] = {
                "rfx_geom": {
                    "top_left": f"{col_to_letter(col_first)}{rw_first + 1}",
                    "bottom_right": f"{col_to_letter(col_last)}{rw_last + 1}",
                },
                "rw_first_head": rw_first_head + 1,
                "rw_first_data": rw_first_data + 1,
                "col_first_data": col_to_letter(col_first_data),
                "page_rows": crw_page,
                "page_cols": ccol_page,
            }

        elif rec_type == BRT_BEGIN_SXVD:
            meta["pivot_fields"] = int(meta["pivot_fields"]) + 1
        elif rec_type == BRT_BEGIN_SXVI:
            meta["pivot_items"] = int(meta["pivot_items"]) + 1

    return meta


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

class XlsbWorkbook:
    """
    Open an .xlsb workbook and iterate worksheet formulas/values and PivotTable metadata.

    Parameters
    ----------
    path : str or path-like
        Path to the .xlsb file.

    Example
    -------
    >>> with XlsbWorkbook("data.xlsb") as wb:
    ...     for sheet, formulas in wb.iter_formulas():
    ...         for (row, col), f in sorted(formulas.items()):
    ...             print(f"{sheet}!{col_to_letter(col)}{row+1}: {f}")
    """

    def __init__(self, path: "os.PathLike"):
        self._zf     = zipfile.ZipFile(str(path), "r")
        self._sst:   List[str]               = []
        self._sheets: List[Tuple[str, str]]  = []
        self._paths:  Dict[str, str]         = {}
        self._dnames: Dict[int, str]         = {}
        self._init_workbook()
        self._init_sst()

    def _read_part(self, path: str) -> bytes:
        path = path.replace("\\", "/").lstrip("/")
        try:
            return self._zf.read(path)
        except KeyError:
            lo = path.lower()
            for n in self._zf.namelist():
                if n.lower() == lo:
                    return self._zf.read(n)
            raise FileNotFoundError(f"Not found in XLSB zip: {path}")

    def _init_workbook(self):
        wb_data = self._read_part("xl/workbook.bin")
        self._sheets = _read_workbook(wb_data)
        self._dnames = _read_defined_names(wb_data)
        try:
            rels = _read_rels(self._read_part("xl/_rels/workbook.bin.rels"))
        except FileNotFoundError:
            rels = {}
        for _name, rel_id in self._sheets:
            target = rels.get(rel_id, "")
            if target:
                full = (f"xl/{target}" if not target.startswith("xl/")
                        else target)
                self._paths[rel_id] = full.replace("//", "/")

    def _init_sst(self):
        try:
            self._sst = _read_sst(self._read_part("xl/sharedStrings.bin"))
        except FileNotFoundError:
            self._sst = []

    @property
    def sheet_names(self) -> List[str]:
        """Ordered list of worksheet names."""
        return [n for n, _ in self._sheets]

    def iter_formulas(
        self,
    ) -> Iterator[Tuple[str, Dict[Tuple[int, int], str]]]:
        """
        Yield ``(sheet_name, formulas)`` for every sheet.

        ``formulas`` maps ``(row, col)`` → formula string (starts with ``=``).
        Row and col are **0-based**.
        """
        all_names = self.sheet_names
        for sheet_name, rel_id in self._sheets:
            zpath = self._paths.get(rel_id)
            if not zpath:
                yield sheet_name, {}; continue
            try:
                ws_data = self._read_part(zpath)
            except FileNotFoundError:
                yield sheet_name, {}; continue
            yield sheet_name, _parse_worksheet(
                ws_data, all_names, self._sst, self._dnames
            )

    def iter_values(
        self,
    ) -> Iterator[Tuple[str, Dict[Tuple[int, int], object]]]:
        """
        Yield ``(sheet_name, values)`` for every sheet.
        ``values`` maps ``(row, col)`` -> cached/constant value.
        """
        for sheet_name, rel_id in self._sheets:
            zpath = self._paths.get(rel_id)
            if not zpath:
                yield sheet_name, {}
                continue
            try:
                ws_data = self._read_part(zpath)
            except FileNotFoundError:
                yield sheet_name, {}
                continue
            yield sheet_name, _parse_worksheet_values(ws_data, self._sst)

    def iter_pivot_tables(self) -> Iterator[Dict[str, object]]:
        """
        Yield parsed PivotTable metadata, linked to sheets where possible.
        """
        # Map sheet part path -> sheet name
        sheet_by_path: Dict[str, str] = {}
        for sheet_name, rel_id in self._sheets:
            zpath = self._paths.get(rel_id)
            if zpath:
                sheet_by_path[zpath] = sheet_name

        # Discover pivot table parts through worksheet relationships.
        seen_parts: set[str] = set()
        for sheet_path, sheet_name in sheet_by_path.items():
            rel_path = (
                f"{os.path.dirname(sheet_path)}/_rels/{os.path.basename(sheet_path)}.rels"
            ).replace("\\", "/")
            try:
                rels = _read_rels(self._read_part(rel_path))
            except FileNotFoundError:
                continue
            for _, target in rels.items():
                resolved = _resolve_rel_target(sheet_path, target)
                if "/pivotTables/" not in resolved:
                    continue
                if resolved in seen_parts:
                    continue
                seen_parts.add(resolved)
                try:
                    pdata = self._read_part(resolved)
                except FileNotFoundError:
                    continue
                meta = _parse_pivot_table_part(pdata)
                meta["sheet"] = sheet_name
                meta["part"] = resolved

                # Resolve linked pivot cache definition if present.
                p_rel_path = (
                    f"{os.path.dirname(resolved)}/_rels/{os.path.basename(resolved)}.rels"
                ).replace("\\", "/")
                try:
                    p_rels = _read_rels(self._read_part(p_rel_path))
                except FileNotFoundError:
                    p_rels = {}
                cache_def = None
                for _, t in p_rels.items():
                    rr = _resolve_rel_target(resolved, t)
                    if "/pivotCache/pivotCacheDefinition" in rr:
                        cache_def = rr
                        break
                if cache_def:
                    meta["pivot_cache_definition"] = cache_def
                yield meta

    def close(self):
        self._zf.close()

    def __enter__(self):  return self
    def __exit__(self, *_): self.close()


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def _cellmap(formulas: Dict[Tuple[int, int], str]) -> Dict[str, str]:
    """Convert {(row,col): formula} to {'A1': formula} with stable ordering."""
    out: Dict[str, str] = {}
    for (row, col), formula in sorted(formulas.items()):
        out[f"{col_to_letter(col)}{row + 1}"] = formula
    return out


def _cellmap_any(cells: Dict[Tuple[int, int], object]) -> Dict[str, object]:
    out: Dict[str, object] = {}
    for (row, col), value in sorted(cells.items()):
        out[f"{col_to_letter(col)}{row + 1}"] = value
    return out


def _collect_formulas(
    wb: "XlsbWorkbook",
    filter_sheet: Optional[str] = None,
) -> Dict[str, Dict[str, str]]:
    out: Dict[str, Dict[str, str]] = {}
    for sheet_name, formulas in wb.iter_formulas():
        if filter_sheet and sheet_name != filter_sheet:
            continue
        if formulas:
            out[sheet_name] = _cellmap(formulas)
    return out


def _collect_values(
    wb: "XlsbWorkbook",
    filter_sheet: Optional[str] = None,
) -> Dict[str, Dict[str, object]]:
    out: Dict[str, Dict[str, object]] = {}
    for sheet_name, values in wb.iter_values():
        if filter_sheet and sheet_name != filter_sheet:
            continue
        if values:
            out[sheet_name] = _cellmap_any(values)
    return out


def _collect_pivots(
    wb: "XlsbWorkbook",
    filter_sheet: Optional[str] = None,
) -> List[Dict[str, object]]:
    out: List[Dict[str, object]] = []
    for pt in wb.iter_pivot_tables():
        if filter_sheet and pt.get("sheet") != filter_sheet:
            continue
        out.append(pt)
    return out


def _as_markdown(
    sheets: List[str],
    formulas: Optional[Dict[str, Dict[str, str]]] = None,
    values: Optional[Dict[str, Dict[str, object]]] = None,
    pivots: Optional[List[Dict[str, object]]] = None,
) -> str:
    lines: List[str] = [f"Sheets: {', '.join(sheets)}", ""]
    emitted = False
    if formulas:
        emitted = True
        lines.append("## Formulas")
        lines.append("")
        for sheet, cell_formulas in formulas.items():
            lines.append(f"### {sheet}")
            for cell, formula in cell_formulas.items():
                lines.append(f"- `{cell}`: `{formula}`")
            lines.append("")
    if values:
        emitted = True
        lines.append("## Values")
        lines.append("")
        for sheet, cell_values in values.items():
            lines.append(f"### {sheet}")
            for cell, value in cell_values.items():
                lines.append(f"- `{cell}`: `{value}`")
            lines.append("")
    if pivots:
        emitted = True
        lines.append("## Pivot Tables")
        lines.append("")
        for pt in pivots:
            lines.append(
                f"- `{pt.get('name') or '<unnamed>'}` "
                f"(sheet: `{pt.get('sheet')}`, cache_id: `{pt.get('cache_id')}`)"
            )
            if pt.get("location"):
                loc = pt["location"]["rfx_geom"]
                lines.append(
                    f"  body: `{loc['top_left']}:{loc['bottom_right']}`; "
                    f"fields: `{pt.get('pivot_fields')}`; "
                    f"items: `{pt.get('pivot_items')}`"
                )
    if not emitted:
        lines.append("(no formulas found)")
        return "\n".join(lines)
    return "\n".join(lines).rstrip() + "\n"


def main():
    import argparse
    parser = argparse.ArgumentParser(
        description="Extract formulas, values, and pivot metadata from an .xlsb workbook."
    )
    parser.add_argument("path", help="Path to .xlsb file")
    parser.add_argument("sheet_name", nargs="?", default=None,
                        help="Optional sheet name filter")
    parser.add_argument(
        "--format",
        dest="output_format",
        choices=("dict", "json", "markdown"),
        default="dict",
        help="Output format (default: dict)",
    )
    parser.add_argument(
        "--include",
        default="formulas",
        help="Comma-separated sections: formulas,values,pivots (default: formulas)",
    )
    args = parser.parse_args()

    with XlsbWorkbook(args.path) as wb:
        includes = {s.strip().lower() for s in args.include.split(",") if s.strip()}
        formulas = _collect_formulas(wb, filter_sheet=args.sheet_name) if "formulas" in includes else {}
        values = _collect_values(wb, filter_sheet=args.sheet_name) if "values" in includes else {}
        pivots = _collect_pivots(wb, filter_sheet=args.sheet_name) if "pivots" in includes else []
        data: Dict[str, object] = {}
        if "formulas" in includes:
            data["formulas"] = formulas
        if "values" in includes:
            data["values"] = values
        if "pivots" in includes:
            data["pivot_tables"] = pivots
        if args.output_format == "json":
            print(json.dumps(data, ensure_ascii=False, indent=2, sort_keys=True))
        elif args.output_format == "markdown":
            print(
                _as_markdown(
                    wb.sheet_names,
                    formulas=formulas if "formulas" in includes else None,
                    values=values if "values" in includes else None,
                    pivots=pivots if "pivots" in includes else None,
                ),
                end="",
            )
        else:
            print(pprint.pformat(data, sort_dicts=True))


if __name__ == "__main__":
    main()
