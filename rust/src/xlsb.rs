//! xlsb.rs — `.xlsb` (Excel Binary Workbook) support.
//!
//! Full port of `xlsb_reader/_reader.py` (the pure-Python BIFF12 parser) to
//! Rust. Produces byte-identical output to the Python reader for
//! `sheet_names`, `iter_formulas()`, `iter_values()`, `iter_filters()`, and
//! `iter_pivot_tables()`.
//!
//! The `.xlsb` container is a plain ZIP — parts are read by the same names
//! Python uses (`xl/workbook.bin`, `xl/worksheets/sheetN.bin`,
//! `xl/sharedStrings.bin`, `xl/_rels/workbook.bin.rels`, worksheet
//! `_rels`, pivot table & pivot cache parts), with the same
//! case-insensitive fallback as `XlsbWorkbook._read_part`.
//!
//! # Known deviations from the Python source (documented, not fixed)
//!
//! The pure-Python reader is written assuming well-formed input and in a
//! few places relies on Python-specific behaviour that has no exact Rust
//! analogue for malformed/truncated bytes:
//!   - `RecordReader._varint` raises an uncaught `EOFError` if a record's
//!     size varint runs off the end of the buffer; here we simply stop
//!     iterating (silently truncating). Never observed on well-formed
//!     `.xlsb` files.
//!   - `_Decompiler._wstr16` reads `self._buf.read(cch * 2)`, which in
//!     Python silently returns a short read past EOF (never raises); the
//!     Rust `common::Reader::wstr16` used here requires the full byte count
//!     to be present. Only diverges on truncated/malformed `PtgStr`
//!     operands.
//!   - `_fmt_num`'s `d == int(d)` raises `ValueError`/`OverflowError` for
//!     NaN/Infinity in Python (caught by the broad `except Exception` in
//!     `_parse_worksheet`, producing an `=<parse_error:...>` formula); the
//!     Rust `fmt_num` returns a parse error too, but the embedded message
//!     text will not match Python's exact exception string. NaN/Infinity
//!     formula literals do not occur in practice (Excel has no NaN/Inf
//!     literal token).
//!
//! None of these affect well-formed `.xlsb` files produced by Excel.

use std::collections::{BTreeMap, HashMap, HashSet};
use std::io::Read;
use std::path::{Path, PathBuf};
use std::sync::{Mutex, OnceLock};

use pyo3::prelude::*;
use pyo3::types::{PyDict, PyList, PyTuple};
use serde_json::{Map, Value};

use crate::common::{
    self, CellValue, FilterInfo, PivotTable, Reader, Result, WorkbookData, XlsbError,
};

// ---------------------------------------------------------------------------
// Record type IDs (subset) — mirrors the constants block in `_reader.py`.
// ---------------------------------------------------------------------------

const BRT_ROW_HDR: u32 = 0x0000;
const BRT_CELL_RK: u32 = 0x0002;
const BRT_CELL_ERROR: u32 = 0x0003;
const BRT_CELL_BOOL: u32 = 0x0004;
const BRT_CELL_REAL: u32 = 0x0005;
const BRT_CELL_ST: u32 = 0x0006;
const BRT_CELL_ISST: u32 = 0x0007;
const BRT_FMLA_STRING: u32 = 0x0008;
const BRT_FMLA_NUM: u32 = 0x0009;
const BRT_FMLA_BOOL: u32 = 0x000A;
const BRT_FMLA_ERROR: u32 = 0x000B;
const BRT_ARR_FMLA: u32 = 0x01AA;
const BRT_SHR_FMLA: u32 = 0x01AB;
const BRT_NAME: u32 = 0x0027;
const BRT_SST_ITEM: u32 = 0x0013;
const BRT_BUNDLE_SH: u32 = 0x009C;
const BRT_BEGIN_SX_VIEW: u32 = 0x0118;
const BRT_BEGIN_SX_LOCATION: u32 = 0x013A;
const BRT_BEGIN_SXVD: u32 = 0x011D;
const BRT_BEGIN_SXVI: u32 = 0x011A;
// AutoFilter records
const BRT_BEGIN_AFILTER: u32 = 0x00A1;
const BRT_END_AFILTER: u32 = 0x00A2;
const BRT_BEGIN_FILTER_COLUMN: u32 = 0x00A3;
const BRT_END_FILTER_COLUMN: u32 = 0x00A4;
const BRT_FILTER: u32 = 0x00A7;
const BRT_BEGIN_CUSTOM_FILTERS: u32 = 0x00AC;
const BRT_CUSTOM_FILTER: u32 = 0x00AE;
// PivotTable filter records
const BRT_BEGIN_SX_FILTER: u32 = 0x0259;
const BRT_END_SX_FILTER: u32 = 0x025A;
// PivotCache definition records
const BRT_BEGIN_PCDFIELD: u32 = 0x00B7;
const BRT_BEGIN_PCDIRUN: u32 = 0x00BF;
const BRT_PCDI_STRING: u32 = 0x0018;

// ---------------------------------------------------------------------------
// Ptg exact-byte token constants (ptg < 0x20)
// ---------------------------------------------------------------------------

const PTG_EXP: u8 = 0x01;
const PTG_TBL: u8 = 0x02;
const PTG_ADD: u8 = 0x03;
const PTG_SUB: u8 = 0x04;
const PTG_MUL: u8 = 0x05;
const PTG_DIV: u8 = 0x06;
const PTG_POWER: u8 = 0x07;
const PTG_CONCAT: u8 = 0x08;
const PTG_LT: u8 = 0x09;
const PTG_LE: u8 = 0x0A;
const PTG_EQ: u8 = 0x0B;
const PTG_GE: u8 = 0x0C;
const PTG_GT: u8 = 0x0D;
const PTG_NE: u8 = 0x0E;
const PTG_ISECT: u8 = 0x0F;
const PTG_UNION: u8 = 0x10;
const PTG_RANGE: u8 = 0x11;
const PTG_UPLUS: u8 = 0x12;
const PTG_UMINUS: u8 = 0x13;
const PTG_PERCENT: u8 = 0x14;
const PTG_PAREN: u8 = 0x15;
const PTG_MISSARG: u8 = 0x16;
const PTG_STR: u8 = 0x17;
const PTG_LIST: u8 = 0x18;
const PTG_ATTR: u8 = 0x19;
const PTG_ERR: u8 = 0x1C;
const PTG_BOOL: u8 = 0x1D;
const PTG_INT: u8 = 0x1E;
const PTG_NUM: u8 = 0x1F;

// Class-variant base values (ptg & 0x1F when ptg >= 0x20)
const B_ARRAY: u8 = 0x00;
const B_FUNC: u8 = 0x01;
const B_FUNC_VAR: u8 = 0x02;
const B_NAME: u8 = 0x03;
const B_REF: u8 = 0x04;
const B_AREA: u8 = 0x05;
const B_MEM_AREA: u8 = 0x06;
const B_MEM_ERR: u8 = 0x07;
const B_MEM_NOMEM: u8 = 0x08;
const B_MEM_FUNC: u8 = 0x09;
const B_REF_ERR: u8 = 0x0A;
const B_AREA_ERR: u8 = 0x0B;
const B_REF_N: u8 = 0x0C;
const B_AREA_N: u8 = 0x0D;
const B_NAME_X: u8 = 0x19;
const B_REF_3D: u8 = 0x1A;
const B_AREA_3D: u8 = 0x1B;
const B_REF_ERR3D: u8 = 0x1C;
const B_AREA_ERR3D: u8 = 0x1D;

// PtgAttr sub-type flags
const ATTR_CHOOSE: u8 = 0x04;
const ATTR_SUM: u8 = 0x10;
const ATTR_SPACE: u8 = 0x40;

/// Error code -> string.
fn err_code_str(code: u8) -> &'static str {
    match code {
        0x00 => "#NULL!",
        0x07 => "#DIV/0!",
        0x0F => "#VALUE!",
        0x17 => "#REF!",
        0x1D => "#NAME?",
        0x24 => "#NUM!",
        0x2A => "#N/A",
        0x2B => "#GETTING_DATA",
        _ => "#ERR",
    }
}

/// BrtCustomFilter grbitSgn -> operator string.
fn filter_op_str(grbit_sgn: u8) -> String {
    match grbit_sgn {
        0x01 => "<".to_string(),
        0x02 => "=".to_string(),
        0x03 => "<=".to_string(),
        0x04 => ">".to_string(),
        0x05 => "<>".to_string(),
        0x06 => ">=".to_string(),
        other => other.to_string(),
    }
}

/// Decode RkNumber (2.5.123). Returns `CellValue::Int` when the RK encodes
/// an exact signed 30-bit integer (fInt set, fX100 clear), `CellValue::Float`
/// otherwise — matches `_rk_to_number` exactly (Python: dividing by 100
/// always upgrades an `int` result to `float`, same as `f_int` clear always
/// yielding a `float` from the packed double bits).
fn rk_to_number(raw: u32) -> CellValue {
    let fx100 = raw & 0x1 != 0;
    let f_int = raw & 0x2 != 0;
    let num30: i64 = (raw >> 2) as i64;
    if f_int {
        let value: i64 = if num30 & (1 << 29) != 0 {
            num30 - (1 << 30)
        } else {
            num30
        };
        if fx100 {
            CellValue::Float(value as f64 / 100.0)
        } else {
            CellValue::Int(value)
        }
    } else {
        let bits64: u64 = (num30 as u64) << 34;
        let value = f64::from_bits(bits64);
        if fx100 {
            CellValue::Float(value / 100.0)
        } else {
            CellValue::Float(value)
        }
    }
}

// ---------------------------------------------------------------------------
// Built-in function table (iftab -> name) — verbatim port of `_FUNC`.
// The local `xslb-file-format.md` override file that `_load_ftab_from_local_spec`
// would read is absent from the repo, so that dynamic override is a no-op in
// practice; only the hardcoded table below is ported.
// ---------------------------------------------------------------------------

#[rustfmt::skip]
const FUNC_TABLE: &[(u16, &str)] = &[
    (0x0000, "COUNT"), (0x0001, "IF"), (0x0002, "ISNA"), (0x0003, "ISERROR"),
    (0x0004, "SUM"), (0x0005, "AVERAGE"), (0x0006, "MIN"), (0x0007, "MAX"),
    (0x0008, "ROW"), (0x0009, "COLUMN"), (0x000A, "NA"), (0x000B, "NPV"),
    (0x000C, "STDEV"), (0x000D, "DOLLAR"), (0x000E, "FIXED"), (0x000F, "SIN"),
    (0x0010, "COS"), (0x0011, "TAN"), (0x0012, "ATAN"), (0x0013, "PI"),
    (0x0014, "SQRT"), (0x0015, "EXP"), (0x0016, "LN"), (0x0017, "LOG10"),
    (0x0018, "ABS"), (0x0019, "INT"), (0x001A, "SIGN"), (0x001B, "ROUND"),
    (0x001C, "LOOKUP"), (0x001D, "INDEX"), (0x001E, "REPT"), (0x001F, "MID"),
    (0x0020, "LEN"), (0x0021, "VALUE"), (0x0022, "TRUE"), (0x0023, "FALSE"),
    (0x0024, "AND"), (0x0025, "OR"), (0x0026, "NOT"), (0x0027, "MOD"),
    (0x0028, "DCOUNT"), (0x0029, "DSUM"), (0x002A, "DAVERAGE"), (0x002B, "DMIN"),
    (0x002C, "DMAX"), (0x002D, "DSTDEV"), (0x002E, "VAR"), (0x002F, "DVAR"),
    (0x0030, "TEXT"), (0x0031, "LINEST"), (0x0032, "TREND"), (0x0033, "LOGEST"),
    (0x0034, "GROWTH"), (0x0038, "PV"), (0x0039, "FV"), (0x003A, "NPER"),
    (0x003B, "PMT"), (0x003C, "RATE"), (0x003D, "MIRR"), (0x003E, "IRR"),
    (0x003F, "RAND"), (0x0040, "MATCH"), (0x0041, "DATE"), (0x0042, "TIME"),
    (0x0043, "DAY"), (0x0044, "MONTH"), (0x0045, "YEAR"), (0x0046, "WEEKDAY"),
    (0x0047, "HOUR"), (0x0048, "MINUTE"), (0x0049, "SECOND"), (0x004A, "NOW"),
    (0x004B, "AREAS"), (0x004C, "ROWS"), (0x004D, "COLUMNS"), (0x004E, "OFFSET"),
    (0x0052, "SEARCH"), (0x0053, "TRANSPOSE"), (0x0056, "TYPE"), (0x0061, "ATAN2"),
    (0x0062, "ASIN"), (0x0063, "ACOS"), (0x0064, "CHOOSE"), (0x0065, "HLOOKUP"),
    (0x0066, "VLOOKUP"), (0x0069, "ISREF"), (0x006D, "LOG"), (0x006F, "CHAR"),
    (0x0070, "LOWER"), (0x0071, "UPPER"), (0x0072, "PROPER"), (0x0073, "LEFT"),
    (0x0074, "RIGHT"), (0x0075, "EXACT"), (0x0076, "TRIM"), (0x0077, "REPLACE"),
    (0x0078, "SUBSTITUTE"), (0x0079, "CODE"), (0x007C, "FIND"), (0x007D, "CELL"),
    (0x007E, "ISERR"), (0x007F, "ISTEXT"), (0x0080, "ISNUMBER"), (0x0081, "ISBLANK"),
    (0x0082, "T"), (0x0083, "N"), (0x008C, "DATEVALUE"), (0x008D, "TIMEVALUE"),
    (0x008E, "SLN"), (0x008F, "SYD"), (0x0090, "DDB"), (0x0094, "INDIRECT"),
    (0x00BD, "DPRODUCT"), (0x00BF, "STDEVP"), (0x00C0, "VARP"), (0x00C1, "DSTDEVP"),
    (0x00C2, "DVARP"), (0x00C3, "TRUNC"), (0x00A2, "CLEAN"), (0x00A9, "COUNTA"),
    (0x00C4, "ISLOGICAL"), (0x00C5, "DCOUNTA"), (0x00C6, "CLEAN"), (0x00C7, "MDETERM"),
    (0x00C8, "MINVERSE"), (0x00C9, "MMULT"), (0x00CB, "IPMT"), (0x00CC, "PPMT"),
    (0x00CD, "COUNTA"), (0x00D7, "PRODUCT"), (0x00D8, "FACT"), (0x00DB, "ADDRESS"),
    (0x00DC, "ISNONTEXT"), (0x00DD, "GETPIVOTDATA"), (0x00DE, "MEDIAN"), (0x00DF, "SUMPRODUCT"),
    (0x00E0, "SINH"), (0x00E1, "COSH"), (0x00E2, "TANH"), (0x00E3, "ASINH"),
    (0x00E4, "ACOSH"), (0x00E5, "ATANH"), (0x00E6, "DGET"), (0x00E8, "ASINH"),
    (0x00E9, "ACOSH"), (0x00EA, "ATANH"), (0x00EF, "INFO"), (0x00F2, "DB"),
    (0x00F9, "FREQUENCY"), (0x0100, "DAYS360"), (0x0101, "TODAY"), (0x0102, "VDB"),
    (0x0109, "ERRORTYPE"), (0x010B, "WORKDAY"), (0x010C, "NETWORKDAYS"), (0x010D, "WEEKNUM"),
    (0x010E, "FLOOR"), (0x010F, "CEILING"), (0x0110, "ISEVEN"), (0x0111, "ISODD"),
    (0x0112, "MROUND"), (0x0113, "QUOTIENT"), (0x0114, "GCD"), (0x0115, "LCM"),
    (0x0116, "MULTINOMIAL"), (0x0117, "COMBIN"), (0x0118, "PERMUT"), (0x011A, "CONFIDENCE"),
    (0x011C, "EVEN"), (0x011D, "EXPONDIST"), (0x011E, "FDIST"), (0x011F, "FINV"),
    (0x0120, "FISHER"), (0x0121, "FISHERINV"), (0x0123, "GAMMADIST"), (0x0124, "GAMMAINV"),
    (0x0126, "HYPGEOMDIST"), (0x0127, "LOGINV"), (0x0128, "LOGNORMDIST"), (0x0129, "NEGBINOMDIST"),
    (0x012A, "NORMDIST"), (0x012B, "NORMSDIST"), (0x012C, "NORMINV"), (0x012D, "NORMSINV"),
    (0x012E, "STANDARDIZE"), (0x012F, "ODD"), (0x0131, "POISSON"), (0x0132, "TDIST"),
    (0x0133, "WEIBULL"), (0x0134, "SUMXMY2"), (0x0135, "SUMX2MY2"), (0x0136, "SUMX2PY2"),
    (0x0137, "CHITEST"), (0x0138, "CORREL"), (0x0139, "COVAR"), (0x013A, "FORECAST"),
    (0x013B, "FTEST"), (0x013C, "INTERCEPT"), (0x013D, "PEARSON"), (0x013E, "RSQ"),
    (0x013F, "STEYX"), (0x0140, "SLOPE"), (0x0141, "TTEST"), (0x0142, "PROB"),
    (0x0143, "DEVSQ"), (0x0144, "GEOMEAN"), (0x0145, "HARMEAN"), (0x0146, "SUMSQ"),
    (0x0147, "KURT"), (0x0148, "SKEW"), (0x0149, "ZTEST"), (0x014A, "LARGE"),
    (0x014B, "SMALL"), (0x014C, "QUARTILE"), (0x014D, "PERCENTILE"), (0x014E, "PERCENTRANK"),
    (0x014F, "MODE"), (0x0150, "TRIMMEAN"), (0x0151, "TINV"), (0x0157, "CONCATENATE"),
    (0x0158, "POWER"), (0x015A, "RADIANS"), (0x015B, "DEGREES"), (0x015C, "SUBTOTAL"),
    (0x015D, "SUMIF"), (0x015E, "COUNTIF"), (0x015F, "COUNTBLANK"), (0x0162, "ROMAN"),
    (0x0163, "HYPERLINK"), (0x0164, "AVERAGEA"), (0x0165, "MAXA"), (0x0166, "MINA"),
    (0x0167, "STDEVPA"), (0x0168, "VARPA"), (0x0169, "STDEVA"), (0x016A, "VARA"),
    (0x0176, "RTD"), (0x017B, "HEX2BIN"), (0x017C, "HEX2DEC"), (0x017D, "HEX2OCT"),
    (0x017E, "DEC2BIN"), (0x017F, "DEC2HEX"), (0x0180, "DEC2OCT"), (0x0181, "OCT2BIN"),
    (0x0182, "OCT2DEC"), (0x0183, "OCT2HEX"), (0x0184, "BIN2DEC"), (0x0185, "BIN2OCT"),
    (0x0186, "BIN2HEX"), (0x0187, "IMSUB"), (0x0188, "IMDIV"), (0x0189, "IMPOWER"),
    (0x018A, "IMABS"), (0x018B, "IMSQRT"), (0x018C, "IMLN"), (0x018D, "IMLOG2"),
    (0x018E, "IMLOG10"), (0x018F, "IMSIN"), (0x0190, "IMCOS"), (0x0191, "IMEXP"),
    (0x0192, "IMARGUMENT"), (0x0193, "IMCONJUGATE"), (0x0194, "IMAGINARY"), (0x0195, "IMREAL"),
    (0x0196, "COMPLEX"), (0x0197, "IMSUM"), (0x0198, "IMPRODUCT"), (0x0199, "SERIESSUM"),
    (0x019A, "FACTDOUBLE"), (0x019B, "SQRTPI"), (0x019D, "DELTA"), (0x019E, "GESTEP"),
    (0x01A2, "ERF"), (0x01A3, "ERFC"), (0x01A4, "BESSELJ"), (0x01A5, "BESSELK"),
    (0x01A6, "BESSELY"), (0x01A7, "BESSELI"), (0x01A8, "XIRR"), (0x01A9, "XNPV"),
    (0x01AA, "PRICEMAT"), (0x01AB, "YIELDMAT"), (0x01AC, "INTRATE"), (0x01AD, "RECEIVED"),
    (0x01AE, "DISC"), (0x01AF, "PRICEDISC"), (0x01B0, "YIELDDISC"), (0x01B1, "TBILLEQ"),
    (0x01B2, "TBILLPRICE"), (0x01B3, "TBILLYIELD"), (0x01B4, "PRICE"), (0x01B5, "YIELD"),
    (0x01B6, "DOLLARDE"), (0x01B7, "DOLLARFR"), (0x01B8, "NOMINAL"), (0x01B9, "EFFECT"),
    (0x01BA, "CUMPRINC"), (0x01BB, "CUMIPMT"), (0x01BC, "EDATE"), (0x01BD, "EOMONTH"),
    (0x01BE, "YEARFRAC"), (0x01BF, "COUPDAYBS"), (0x01C0, "COUPDAYS"), (0x01C1, "COUPDAYSNC"),
    (0x01C2, "COUPNCD"), (0x01C3, "COUPNUM"), (0x01C4, "COUPPCD"), (0x01C5, "DURATION"),
    (0x01C6, "MDURATION"), (0x01C7, "ODDLPRICE"), (0x01C8, "ODDLYIELD"), (0x01C9, "ODDFPRICE"),
    (0x01CA, "ODDFYIELD"), (0x01CB, "RANDBETWEEN"), (0x01CC, "WEEKNUM"), (0x01CD, "AMORDEGRC"),
    (0x01CE, "AMORLINC"), (0x01CF, "CONVERT"), (0x01D0, "ACCRINT"), (0x01D1, "ACCRINTM"),
    (0x01D2, "WORKDAY"), (0x01D3, "NETWORKDAYS"), (0x01D4, "GCD"), (0x01D5, "MULTINOMIAL"),
    (0x01D6, "LCM"), (0x01D7, "FVSCHEDULE"), (0x01D8, "CUBESETCOUNT"), (0x01D9, "CUBESET"),
    (0x01DA, "IFERROR"), (0x01DB, "COUNTIFS"), (0x01DC, "SUMIFS"), (0x01DD, "AVERAGEIF"),
    (0x01DE, "AVERAGEIFS"), (0x01DF, "AGGREGATE"), (0x01E3, "AVERAGEIF"), (0x01E4, "AVERAGEIFS"),
    (0x0213, "IFNA"), (0x0214, "IFS"), (0x0215, "SWITCH"), (0x0216, "MAXIFS"),
    (0x0217, "MINIFS"), (0x0241, "XLOOKUP"), (0x0244, "XMATCH"), (0x0246, "FILTER"),
    (0x0247, "SORT"), (0x0248, "SORTBY"), (0x0249, "UNIQUE"), (0x024A, "SEQUENCE"),
    (0x024B, "RANDARRAY"), (0x024C, "LET"),
];

fn func_table() -> &'static HashMap<u16, &'static str> {
    static TABLE: OnceLock<HashMap<u16, &'static str>> = OnceLock::new();
    TABLE.get_or_init(|| {
        let mut m = HashMap::with_capacity(FUNC_TABLE.len());
        for &(k, v) in FUNC_TABLE {
            m.insert(k, v);
        }
        m
    })
}

fn func_name(iftab: u16) -> String {
    match func_table().get(&iftab) {
        Some(name) => name.to_string(),
        None => format!("_func{iftab:#06x}"),
    }
}

/// Fixed arity for PtgFunc tokens — verbatim port of `_FIXED_ARITY`.
#[rustfmt::skip]
const FIXED_ARITY: &[(u16, usize)] = &[
    (0x0003, 1), (0x000A, 0), (0x000F, 1), (0x0010, 1), (0x0011, 1), (0x0012, 1),
    (0x0013, 0), (0x0014, 1), (0x0015, 1), (0x0016, 1), (0x0017, 1), (0x0018, 1),
    (0x0019, 1), (0x001A, 1), (0x001B, 2), (0x0022, 0), (0x0023, 0), (0x0026, 1),
    (0x0027, 2), (0x003F, 0), (0x0041, 3), (0x0042, 3), (0x0043, 1), (0x0044, 1),
    (0x0045, 1), (0x0046, 1), (0x0047, 1), (0x0048, 1), (0x0049, 1), (0x004A, 0),
    (0x004B, 1), (0x004C, 1), (0x004D, 1), (0x0061, 2), (0x0062, 1), (0x0063, 1),
    (0x006D, 2), (0x006F, 1), (0x0070, 1), (0x0071, 1), (0x0072, 1), (0x0075, 2),
    (0x0076, 1), (0x007E, 1), (0x007F, 1), (0x0080, 1), (0x0081, 1), (0x0082, 1),
    (0x0083, 1), (0x00C3, 1), (0x00C4, 1), (0x00C6, 1), (0x00D8, 1), (0x0100, 2),
    (0x0101, 0), (0x010E, 2), (0x010F, 2), (0x011C, 1), (0x012F, 1), (0x0157, 2),
    (0x0158, 2), (0x015A, 1), (0x015B, 1),
];

fn fixed_arity() -> &'static HashMap<u16, usize> {
    static TABLE: OnceLock<HashMap<u16, usize>> = OnceLock::new();
    TABLE.get_or_init(|| {
        let mut m = HashMap::with_capacity(FIXED_ARITY.len());
        for &(k, v) in FIXED_ARITY {
            m.insert(k, v);
        }
        m
    })
}

// ---------------------------------------------------------------------------
// Small helpers
// ---------------------------------------------------------------------------

fn cell_ref(row: i64, col: u32) -> String {
    format!("{}{}", common::col_to_letter(col), row + 1)
}

fn sheet_prefix(sheet_names: &[String], ixti: u16) -> String {
    let idx = ixti as usize;
    if idx < sheet_names.len() {
        let name = &sheet_names[idx];
        if name.chars().any(|c| " ![]'".contains(c)) {
            format!("'{name}'!")
        } else {
            format!("{name}!")
        }
    } else {
        format!("[sheet{ixti}]!")
    }
}

/// Format a double like Excel (no unnecessary decimals). Port of `_fmt_num`.
/// Returns `Err` for non-finite values, mirroring the uncaught
/// `ValueError`/`OverflowError` Python's `int(d)` raises for NaN/Infinity
/// (caught by the broad `except Exception` around formula decompilation).
fn fmt_num(d: f64) -> Result<String> {
    if d.is_finite() && d == (d as i64) as f64 && d.abs() < 1e15 {
        return Ok((d as i64).to_string());
    }
    if !d.is_finite() {
        return Err(XlsbError::Parse(
            "cannot convert float NaN/Infinity to integer".to_string(),
        ));
    }
    Ok(python_float_repr(d))
}

/// Mimic CPython's `repr(float)` formatting: shortest round-trip digits,
/// switching to scientific notation when the decimal point position
/// (`decpt`) is `<= -4` or `> 16` (same rule as `format_float_short` in
/// CPython's `pystrtod.c` for the `'r'` format code).
pub(crate) fn python_float_repr(d: f64) -> String {
    let neg = d.is_sign_negative();
    let ad = d.abs();
    let sci = format!("{ad:e}");
    let (mantissa, exp_str) = sci.split_once('e').expect("LowerExp always emits 'e'");
    let exp: i32 = exp_str.parse().expect("valid exponent digits");
    let digits: String = mantissa.chars().filter(|c| *c != '.').collect();
    let decpt = exp + 1;
    let use_exp = decpt <= -4 || decpt > 16;
    let body = if use_exp {
        let mant = if digits.len() > 1 {
            format!("{}.{}", &digits[0..1], &digits[1..])
        } else {
            digits.clone()
        };
        let e = decpt - 1;
        let sign = if e < 0 { "-" } else { "+" };
        format!("{mant}e{sign}{:02}", e.abs())
    } else if decpt <= 0 {
        format!("0.{}{}", "0".repeat((-decpt) as usize), digits)
    } else if (decpt as usize) >= digits.len() {
        format!("{}{}.0", digits, "0".repeat(decpt as usize - digits.len()))
    } else {
        format!(
            "{}.{}",
            &digits[0..decpt as usize],
            &digits[decpt as usize..]
        )
    };
    if neg {
        format!("-{body}")
    } else {
        body
    }
}

/// Read a `u32`-length-prefixed UTF-16LE string leniently: if fewer than 4
/// bytes remain for the length prefix, or fewer than `cch*2` bytes remain
/// for the body, returns `""` WITHOUT erroring — matches `_read_xlwide_from`
/// (which uses plain `BytesIO.read()`, tolerating short reads, rather than
/// `struct.unpack` which would raise).
fn read_xlwide_from(r: &mut Reader<'_>) -> String {
    let n = 4.min(r.remaining());
    let cch_bytes = match r.take_bytes(n) {
        Ok(b) => b,
        Err(_) => return String::new(),
    };
    if cch_bytes.len() < 4 {
        return String::new();
    }
    let cch = u32::from_le_bytes([cch_bytes[0], cch_bytes[1], cch_bytes[2], cch_bytes[3]]) as usize;
    let want = cch.saturating_mul(2);
    let n2 = want.min(r.remaining());
    let raw = match r.take_bytes(n2) {
        Ok(b) => b,
        Err(_) => return String::new(),
    };
    if raw.len() < want {
        return String::new();
    }
    utf16le_lossy(raw)
}

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
// BIFF12 record stream reader — port of `RecordReader`.
// ---------------------------------------------------------------------------

struct RecordReader<'a> {
    buf: &'a [u8],
    pos: usize,
}

impl<'a> RecordReader<'a> {
    fn new(buf: &'a [u8]) -> Self {
        RecordReader { buf, pos: 0 }
    }

    /// Port of `RecordReader._varint`. Returns `None` on truncated input
    /// (Python raises an uncaught `EOFError` here; we stop iterating
    /// instead — see module-level doc comment).
    fn varint(&mut self) -> Option<u32> {
        let mut result: u32 = 0;
        let mut shift: u32 = 0;
        for _ in 0..4 {
            if self.pos >= self.buf.len() {
                return None;
            }
            let b = self.buf[self.pos];
            self.pos += 1;
            result |= ((b & 0x7F) as u32) << shift;
            shift += 7;
            if b & 0x80 == 0 {
                break;
            }
        }
        Some(result)
    }
}

impl<'a> Iterator for RecordReader<'a> {
    type Item = (u32, &'a [u8]);

    fn next(&mut self) -> Option<Self::Item> {
        if self.pos >= self.buf.len() {
            return None;
        }
        let b0 = self.buf[self.pos];
        self.pos += 1;
        let rec_type: u32 = if b0 & 0x80 != 0 {
            if self.pos >= self.buf.len() {
                return None;
            }
            let b1 = self.buf[self.pos];
            self.pos += 1;
            (b0 & 0x7F) as u32 | (((b1 & 0x7F) as u32) << 7)
        } else {
            b0 as u32
        };
        let size = self.varint()? as usize;
        let end = self.pos.saturating_add(size).min(self.buf.len());
        let data = &self.buf[self.pos..end];
        self.pos += size;
        Some((rec_type, data))
    }
}

// ---------------------------------------------------------------------------
// RgceLoc/RgceArea raw field reads + relative-offset resolution
// ---------------------------------------------------------------------------

fn read_rgce_loc_raw(r: &mut Reader<'_>) -> Result<(u32, u16)> {
    let row_raw = r.u32()?;
    let col_raw = r.u16()?;
    Ok((row_raw, col_raw))
}

fn read_rgce_area_raw(r: &mut Reader<'_>) -> Result<(u32, u32, u16, u16)> {
    let row_first = r.u32()?;
    let row_last = r.u32()?;
    let col_first = r.u16()?;
    let col_last = r.u16()?;
    Ok((row_first, row_last, col_first, col_last))
}

/// Resolve RgceLoc/RgceLocRel into a concrete row/col. Port of
/// `_resolve_loc`. Returns `(row, col, row_rel, col_rel)`.
#[allow(clippy::too_many_arguments)]
fn resolve_loc(
    row_raw: u32,
    col_raw: u16,
    base_row: Option<u32>,
    base_col: Option<u32>,
    apply_relative_offsets: bool,
) -> (u32, u32, bool, bool) {
    let col_masked = (col_raw & 0x3FFF) as u32;
    let col_rel = col_raw & 0x4000 != 0;
    let row_rel = col_raw & 0x8000 != 0;
    let mut row = row_raw;
    let mut col = col_masked;

    if apply_relative_offsets && row_rel {
        if let Some(base_row) = base_row {
            let raw_off = (row_raw & 0x000F_FFFF) as i64;
            let off = if raw_off >= 0x0008_0000 {
                raw_off - 0x0010_0000
            } else {
                raw_off
            };
            row = ((base_row as i64 + off).rem_euclid(0x0010_0000)) as u32;
        }
    }

    if apply_relative_offsets && col_rel {
        if let Some(base_col) = base_col {
            let col_i = col_masked as i64;
            let off = if col_i < 0x2000 {
                col_i
            } else {
                col_i - 0x4000
            };
            col = ((base_col as i64 + off).rem_euclid(0x4000)) as u32;
        }
    }

    (row, col, row_rel, col_rel)
}

fn loc_str(row: u32, col: u32, row_rel: bool, col_rel: bool) -> String {
    let c = if col_rel {
        common::col_to_letter(col)
    } else {
        format!("${}", common::col_to_letter(col))
    };
    let r = if row_rel {
        (row + 1).to_string()
    } else {
        format!("${}", row + 1)
    };
    format!("{c}{r}")
}

// ---------------------------------------------------------------------------
// Ptg decompiler — port of `_Decompiler`.
// ---------------------------------------------------------------------------

struct Decompiler<'a> {
    r: Reader<'a>,
    sheets: &'a [String],
    dnames: &'a BTreeMap<u32, String>,
    base_row: Option<u32>,
    base_col: Option<u32>,
    rgcb: &'a [u8],
}

fn pop(stack: &mut Vec<String>) -> String {
    stack.pop().unwrap_or_else(|| "?".to_string())
}

fn pop_n(stack: &mut Vec<String>, n: usize) -> Vec<String> {
    let mut args = Vec::with_capacity(n);
    for _ in 0..n {
        args.push(pop(stack));
    }
    args.reverse();
    args
}

impl<'a> Decompiler<'a> {
    /// Skip up to `n` bytes, clamped to what remains.
    ///
    /// Python builds `self._buf` as an `io.BytesIO` and discards trailing
    /// placeholder/reserved fields with plain `.read(n)`, which silently
    /// short-reads past EOF instead of raising. `Reader::skip` is strict, so
    /// call sites that mirror one of these discard reads (PtgArray's 7
    /// placeholder bytes, the `#REF!` error tokens' reserved fields, etc.)
    /// use this instead to match Python's tolerant behavior exactly.
    fn skip_lenient(&mut self, n: usize) -> Result<()> {
        let n = self.r.remaining().min(n);
        self.r.skip(n)
    }

    fn new(
        rgce: &'a [u8],
        sheet_names: &'a [String],
        defined_names: &'a BTreeMap<u32, String>,
        base_row: Option<u32>,
        base_col: Option<u32>,
        rgcb: &'a [u8],
    ) -> Self {
        Decompiler {
            r: Reader::new(rgce),
            sheets: sheet_names,
            dnames: defined_names,
            base_row,
            base_col,
            rgcb,
        }
    }

    fn decompile(&mut self) -> Result<String> {
        let mut stack: Vec<String> = Vec::new();

        while !self.r.is_empty() {
            let ptg = self.r.u8()?;

            if ptg < 0x20 {
                self.step_exact_byte(ptg, &mut stack)?;
            } else {
                let base = ptg & 0x1F;
                if !self.step_class_variant(base, &mut stack)? {
                    stack.push(format!("?ptg{ptg:#04x}"));
                    break;
                }
            }
        }

        Ok(stack.last().cloned().unwrap_or_default())
    }

    fn step_exact_byte(&mut self, ptg: u8, stack: &mut Vec<String>) -> Result<()> {
        let binop = match ptg {
            PTG_ADD => Some("+"),
            PTG_SUB => Some("-"),
            PTG_MUL => Some("*"),
            PTG_DIV => Some("/"),
            PTG_POWER => Some("^"),
            PTG_CONCAT => Some("&"),
            PTG_LT => Some("<"),
            PTG_LE => Some("<="),
            PTG_EQ => Some("="),
            PTG_GE => Some(">="),
            PTG_GT => Some(">"),
            PTG_NE => Some("<>"),
            PTG_RANGE => Some(":"),
            _ => None,
        };
        if let Some(op) = binop {
            let right = pop(stack);
            let left = pop(stack);
            stack.push(format!("{left}{op}{right}"));
            return Ok(());
        }

        match ptg {
            PTG_ISECT => {
                let right = pop(stack);
                let left = pop(stack);
                stack.push(format!("{left} {right}"));
            }
            PTG_UNION => {
                let right = pop(stack);
                let left = pop(stack);
                stack.push(format!("({left},{right})"));
            }
            PTG_UPLUS => {
                let v = pop(stack);
                stack.push(format!("+{v}"));
            }
            PTG_UMINUS => {
                let v = pop(stack);
                stack.push(format!("-{v}"));
            }
            PTG_PERCENT => {
                let v = pop(stack);
                stack.push(format!("{v}%"));
            }
            PTG_PAREN => {
                let v = pop(stack);
                stack.push(format!("({v})"));
            }
            PTG_MISSARG => {
                stack.push(String::new());
            }
            PTG_STR => {
                let s = self.r.wstr16()?;
                stack.push(format!("\"{s}\""));
            }
            PTG_LIST => {
                let eptg = self.r.u8()?;
                if eptg == 0x19 {
                    self.r.u16()?; // ixti
                    let f1 = self.r.u8()?;
                    let f2 = self.r.u8()?;
                    let columns = f1 & 0x03;
                    let invalid = f2 & 0x10 != 0;
                    let nonresident = f2 & 0x20 != 0;
                    let list_index = self.r.u32()?;
                    let col_first = self.r.u16()?;
                    let col_last = self.r.u16()?;
                    if invalid {
                        stack.push("#REF!".to_string());
                    } else if nonresident {
                        stack.push(format!("Table{list_index}[?]"));
                    } else if columns == 0 {
                        stack.push(format!("Table{list_index}"));
                    } else if columns == 1 {
                        stack.push(format!("Table{list_index}[col{}]", col_first + 1));
                    } else {
                        stack.push(format!(
                            "Table{list_index}[col{}:col{}]",
                            col_first + 1,
                            col_last + 1
                        ));
                    }
                } else if eptg == 0x1D {
                    let sx_index = self.r.u32()?;
                    stack.push(format!("PivotName{sx_index}"));
                } else {
                    stack.push(format!("?ptglist{eptg:#04x}"));
                }
            }
            PTG_ATTR => {
                let atype = self.r.u8()?;
                if atype & ATTR_SUM != 0 {
                    self.r.u16()?; // skip size word
                    let v = pop(stack);
                    stack.push(format!("SUM({v})"));
                } else if atype & ATTR_CHOOSE != 0 {
                    let ncases = self.r.u16()?;
                    for _ in 0..=ncases {
                        self.r.u16()?; // offset array
                    }
                } else if atype & ATTR_SPACE != 0 {
                    self.r.u8()?;
                    self.r.u8()?; // space type + count
                } else {
                    self.r.u16()?; // GOTO/IF/SEMI/BAXCEL offset
                }
            }
            PTG_ERR => {
                let code = self.r.u8()?;
                stack.push(err_code_str(code).to_string());
            }
            PTG_BOOL => {
                let v = self.r.u8()?;
                stack.push(if v != 0 { "TRUE" } else { "FALSE" }.to_string());
            }
            PTG_INT => {
                let v = self.r.u16()?;
                stack.push(v.to_string());
            }
            PTG_NUM => {
                let v = self.r.f64()?;
                stack.push(fmt_num(v)?);
            }
            PTG_EXP => {
                // XLSB PtgExp stores row only; column lives in rgcb as PtgExtraCol.
                let row = self.r.u32()?;
                let col = if self.rgcb.len() >= 4 {
                    u32::from_le_bytes([self.rgcb[0], self.rgcb[1], self.rgcb[2], self.rgcb[3]])
                } else {
                    0
                };
                stack.push(format!("{{array@{}}}", cell_ref(row as i64, col)));
            }
            PTG_TBL => {
                let row = self.r.i32()?;
                let col = self.r.u16()?;
                stack.push(format!("TABLE({})", cell_ref(row as i64, col as u32)));
            }
            _ => {
                stack.push(format!("?ptg{ptg:#04x}"));
            }
        }
        Ok(())
    }

    /// Returns `Ok(true)` if `base` was a recognized class-variant token,
    /// `Ok(false)` for the fallback `?ptg...` case (caller pushes the
    /// fallback text itself and stops decompiling, matching Python's
    /// trailing `break`).
    fn step_class_variant(&mut self, base: u8, stack: &mut Vec<String>) -> Result<bool> {
        match base {
            B_ARRAY => {
                // unused1(4) + unused2(2) + unused3(4) + unused4(4) = 14 bytes.
                self.skip_lenient(14)?;
                stack.push("{array}".to_string());
            }
            B_FUNC => {
                let iftab = self.r.u16()?;
                let name = func_name(iftab);
                let args = match fixed_arity().get(&iftab) {
                    Some(&arity) => pop_n(stack, arity),
                    None => vec![pop(stack)],
                };
                stack.push(format!("{name}({})", args.join(",")));
            }
            B_FUNC_VAR => {
                let cparams_raw = self.r.u8()?;
                let is_ce = cparams_raw & 0x80 != 0;
                let cparams = (cparams_raw & 0x7F) as usize;
                let iftab = self.r.u16()?;
                let args = pop_n(stack, cparams);
                if iftab == 0x00FF && !args.is_empty() {
                    let raw_name = args[0].trim();
                    let call_name = raw_name.replace("_xlfn.", "");
                    let call_name = if !call_name.is_empty() && !call_name.starts_with('_') {
                        format!("@{call_name}")
                    } else {
                        call_name
                    };
                    stack.push(format!("{call_name}({})", args[1..].join(",")));
                } else {
                    let name = if is_ce {
                        format!("_xlfn.{}", func_name(iftab))
                    } else {
                        func_name(iftab)
                    };
                    stack.push(format!("{name}({})", args.join(",")));
                }
            }
            B_NAME => {
                let nameindex = self.r.u32()?;
                stack.push(
                    self.dnames
                        .get(&nameindex)
                        .cloned()
                        .unwrap_or_else(|| format!("name{nameindex}")),
                );
            }
            B_REF => {
                let (row_raw, col_raw) = read_rgce_loc_raw(&mut self.r)?;
                let (row, col, rr, cr) =
                    resolve_loc(row_raw, col_raw, self.base_row, self.base_col, false);
                stack.push(loc_str(row, col, rr, cr));
            }
            B_AREA => {
                let (r1_raw, r2_raw, c1_raw, c2_raw) = read_rgce_area_raw(&mut self.r)?;
                let (r1, c1, r1r, c1r) =
                    resolve_loc(r1_raw, c1_raw, self.base_row, self.base_col, false);
                let (r2, c2, r2r, c2r) =
                    resolve_loc(r2_raw, c2_raw, self.base_row, self.base_col, false);
                stack.push(format!(
                    "{}:{}",
                    loc_str(r1, c1, r1r, c1r),
                    loc_str(r2, c2, r2r, c2r)
                ));
            }
            B_MEM_AREA | B_MEM_ERR | B_MEM_NOMEM | B_MEM_FUNC => {
                self.skip_lenient(4)?;
                self.r.u16()?; // 4 reserved + cce
            }
            B_REF_ERR => {
                self.skip_lenient(6)?;
                stack.push("#REF!".to_string());
            }
            B_AREA_ERR => {
                self.skip_lenient(12)?;
                stack.push("#REF!:#REF!".to_string());
            }
            B_REF_N => {
                let (row_raw, col_raw) = read_rgce_loc_raw(&mut self.r)?;
                let (row, col, rr, cr) =
                    resolve_loc(row_raw, col_raw, self.base_row, self.base_col, true);
                stack.push(loc_str(row, col, rr, cr));
            }
            B_AREA_N => {
                let (r1_raw, r2_raw, c1_raw, c2_raw) = read_rgce_area_raw(&mut self.r)?;
                let (r1, c1, r1r, c1r) =
                    resolve_loc(r1_raw, c1_raw, self.base_row, self.base_col, true);
                let (r2, c2, r2r, c2r) =
                    resolve_loc(r2_raw, c2_raw, self.base_row, self.base_col, true);
                stack.push(format!(
                    "{}:{}",
                    loc_str(r1, c1, r1r, c1r),
                    loc_str(r2, c2, r2r, c2r)
                ));
            }
            B_NAME_X => {
                self.r.u16()?; // ixti
                let nameindex = self.r.u32()?;
                stack.push(
                    self.dnames
                        .get(&nameindex)
                        .cloned()
                        .unwrap_or_else(|| format!("name{nameindex}")),
                );
            }
            B_REF_3D => {
                let ixti = self.r.u16()?;
                let (row_raw, col_raw) = read_rgce_loc_raw(&mut self.r)?;
                let (row, col, rr, cr) =
                    resolve_loc(row_raw, col_raw, self.base_row, self.base_col, false);
                let pfx = sheet_prefix(self.sheets, ixti);
                stack.push(format!("{pfx}{}", loc_str(row, col, rr, cr)));
            }
            B_AREA_3D => {
                let ixti = self.r.u16()?;
                let (r1_raw, r2_raw, c1_raw, c2_raw) = read_rgce_area_raw(&mut self.r)?;
                let (r1, c1, r1r, c1r) =
                    resolve_loc(r1_raw, c1_raw, self.base_row, self.base_col, false);
                let (r2, c2, r2r, c2r) =
                    resolve_loc(r2_raw, c2_raw, self.base_row, self.base_col, false);
                let pfx = sheet_prefix(self.sheets, ixti);
                stack.push(format!(
                    "{pfx}{}:{}",
                    loc_str(r1, c1, r1r, c1r),
                    loc_str(r2, c2, r2r, c2r)
                ));
            }
            B_REF_ERR3D => {
                self.skip_lenient(2 + 6)?;
                stack.push("#REF!".to_string());
            }
            B_AREA_ERR3D => {
                self.skip_lenient(2 + 12)?;
                stack.push("#REF!".to_string());
            }
            _ => return Ok(false),
        }
        Ok(true)
    }
}

// ---------------------------------------------------------------------------
// Shared-string table
// ---------------------------------------------------------------------------

fn read_sst(data: &[u8]) -> Vec<String> {
    let mut strings = Vec::new();
    for (rec_type, payload) in RecordReader::new(data) {
        if rec_type == BRT_SST_ITEM {
            if payload.len() < 5 {
                strings.push(String::new());
                continue;
            }
            let mut r = Reader::new(&payload[1..]); // skip flags byte
            let cch = match r.u32() {
                Ok(v) => v as usize,
                Err(_) => {
                    strings.push(String::new());
                    continue;
                }
            };
            let want = cch.saturating_mul(2);
            let n = want.min(r.remaining());
            let raw = r.take_bytes(n).unwrap_or(&[]);
            strings.push(utf16le_lossy(raw));
        }
    }
    strings
}

// ---------------------------------------------------------------------------
// Workbook part (xl/workbook.bin)
// ---------------------------------------------------------------------------

/// Return `[(sheet_name, rel_id), ...]` in tab order. Port of
/// `_read_workbook`.
fn read_workbook(data: &[u8]) -> Vec<(String, String)> {
    let mut sheets = Vec::new();
    for (rec_type, payload) in RecordReader::new(data) {
        if rec_type == BRT_BUNDLE_SH && payload.len() >= 8 {
            let mut r = Reader::new(payload);
            if r.skip(8).is_err() {
                continue;
            }
            let rel_id = read_xlwide_from(&mut r);
            let name = read_xlwide_from(&mut r);
            sheets.push((name, rel_id));
        }
    }
    sheets
}

/// Return `{1-based nameindex -> defined-name text}` from BrtName records.
/// Port of `_read_defined_names`.
fn read_defined_names(data: &[u8]) -> BTreeMap<u32, String> {
    let mut out = BTreeMap::new();
    let mut idx: u32 = 0;
    for (rec_type, payload) in RecordReader::new(data) {
        if rec_type != BRT_NAME || payload.len() < 13 {
            continue;
        }
        idx += 1;
        let mut r = Reader::new(payload);
        if r.skip(4).is_err() || r.skip(1).is_err() || r.skip(4).is_err() {
            continue;
        }
        let n = 4.min(r.remaining());
        let cch_bytes = match r.take_bytes(n) {
            Ok(b) if b.len() == 4 => b,
            _ => continue,
        };
        let cch =
            u32::from_le_bytes([cch_bytes[0], cch_bytes[1], cch_bytes[2], cch_bytes[3]]) as usize;
        let want = cch.saturating_mul(2);
        if r.remaining() < want {
            continue;
        }
        let rgch = r.take_bytes(want).unwrap_or(&[]);
        out.insert(idx, utf16le_lossy(rgch));
    }
    out
}

// ---------------------------------------------------------------------------
// Relationships
// ---------------------------------------------------------------------------

fn rels_regex() -> &'static regex::Regex {
    static RE: OnceLock<regex::Regex> = OnceLock::new();
    RE.get_or_init(|| {
        regex::Regex::new(r#"<Relationship[^>]+Id="([^"]+)"[^>]+Target="([^"]+)""#).unwrap()
    })
}

/// Returns `[(id, target), ...]` in document order (mirrors the fact that
/// Python's dict-comprehension-built `_read_rels` result preserves
/// insertion order, which `iter_pivot_tables` relies on when walking
/// `rels.items()`).
fn read_rels(xml_data: &[u8]) -> Vec<(String, String)> {
    let text = String::from_utf8_lossy(xml_data);
    rels_regex()
        .captures_iter(&text)
        .map(|c| (c[1].to_string(), c[2].to_string()))
        .collect()
}

fn rels_get<'a>(rels: &'a [(String, String)], id: &str) -> Option<&'a str> {
    rels.iter().find(|(k, _)| k == id).map(|(_, v)| v.as_str())
}

fn dirname(p: &str) -> String {
    match p.rfind('/') {
        Some(i) => p[..i].to_string(),
        None => String::new(),
    }
}

fn basename(p: &str) -> String {
    match p.rfind('/') {
        Some(i) => p[i + 1..].to_string(),
        None => p.to_string(),
    }
}

fn normalize_path(p: &str) -> String {
    let mut parts: Vec<&str> = Vec::new();
    for comp in p.split('/') {
        match comp {
            "" | "." => {}
            ".." => match parts.last() {
                Some(&last) if last != ".." => {
                    parts.pop();
                }
                _ => parts.push(".."),
            },
            other => parts.push(other),
        }
    }
    parts.join("/")
}

/// Resolve an OPC relationship target relative to the source part path.
/// Port of `_resolve_rel_target`.
fn resolve_rel_target(base_part: &str, target: &str) -> String {
    if target.starts_with('/') {
        return normalize_path(target);
    }
    let base_dir = dirname(base_part.trim_start_matches('/'));
    let joined = if base_dir.is_empty() {
        target.to_string()
    } else {
        format!("{base_dir}/{target}")
    };
    normalize_path(&joined)
}

// ---------------------------------------------------------------------------
// Worksheet parser
// ---------------------------------------------------------------------------

fn is_fmla_rec(rec_type: u32) -> bool {
    matches!(
        rec_type,
        BRT_FMLA_STRING | BRT_FMLA_NUM | BRT_FMLA_BOOL | BRT_FMLA_ERROR
    )
}

/// Read CellParsedFormula/SharedParsedFormula payload from the current
/// position. Returns `(rgce, rgcb)` or `None` if malformed. Port of
/// `_read_cell_parsed_formula`.
fn read_cell_parsed_formula<'a>(r: &mut Reader<'a>) -> Option<(&'a [u8], &'a [u8])> {
    let cce = r.u32().ok()? as usize;
    if cce == 0 || cce > 16384 {
        return None;
    }
    let rgce = r.take_bytes(cce).ok()?;
    let cb = r.u32().ok()? as usize;
    let rgcb = if cb != 0 { r.take_bytes(cb).ok()? } else { &[] };
    Some((rgce, rgcb))
}

fn decompile_formula(
    rgce: &[u8],
    sheet_names: &[String],
    dn: &BTreeMap<u32, String>,
    base_row: u32,
    base_col: u32,
    rgcb: &[u8],
) -> String {
    let mut dc = Decompiler::new(rgce, sheet_names, dn, Some(base_row), Some(base_col), rgcb);
    match dc.decompile() {
        Ok(s) => format!("={s}"),
        Err(e) => format!("=<parse_error:{e}>"),
    }
}

/// A pending shared/array formula: `(row_first, row_last, col_first, col_last, rgce, rgcb)`.
type SharedFormula = (u32, u32, u32, u32, Vec<u8>, Vec<u8>);

/// Parse a worksheet `.bin` part. Returns `{(row, col): "=formula_string"}`.
/// Row and col are 0-based. Port of `_parse_worksheet`.
fn parse_worksheet(
    data: &[u8],
    sheet_names: &[String],
    dn: &BTreeMap<u32, String>,
) -> BTreeMap<(u32, u32), String> {
    let mut formulas: BTreeMap<(u32, u32), String> = BTreeMap::new();
    let mut current_row: u32 = 0;
    let mut shared_formulas: Vec<SharedFormula> = Vec::new();
    let mut pending_exp: Vec<(u32, u32)> = Vec::new();
    let mut last_formula_cell: Option<(u32, u32)> = None;

    for (rec_type, payload) in RecordReader::new(data) {
        if rec_type == BRT_ROW_HDR {
            if payload.len() >= 4 {
                current_row = u32::from_le_bytes([payload[0], payload[1], payload[2], payload[3]]);
            }
            last_formula_cell = None;
        } else if is_fmla_rec(rec_type) {
            if payload.len() < 10 {
                continue;
            }
            let mut r = Reader::new(payload);
            let col = match r.u32() {
                Ok(v) => v,
                Err(_) => continue,
            };
            if r.skip(4).is_err() {
                continue;
            }
            let cached_ok = match rec_type {
                BRT_FMLA_STRING => (|| -> Result<()> {
                    let cch = r.u32()? as usize;
                    r.skip(cch.saturating_mul(2))
                })()
                .is_ok(),
                BRT_FMLA_NUM => r.skip(8).is_ok(),
                BRT_FMLA_BOOL | BRT_FMLA_ERROR => r.skip(1).is_ok(),
                _ => true,
            };
            if !cached_ok || r.skip(2).is_err() {
                continue;
            }
            let (rgce, rgcb) = match read_cell_parsed_formula(&mut r) {
                Some(v) => v,
                None => continue,
            };

            if rgce.first() == Some(&PTG_EXP) {
                pending_exp.push((current_row, col));
                formulas.insert((current_row, col), "={shared_formula}".to_string());
            } else {
                let formula = decompile_formula(rgce, sheet_names, dn, current_row, col, rgcb);
                formulas.insert((current_row, col), formula);
            }
            last_formula_cell = Some((current_row, col));
        } else if rec_type == BRT_SHR_FMLA || rec_type == BRT_ARR_FMLA {
            if payload.len() < 20 {
                continue;
            }
            let mut r = Reader::new(payload);
            let (rw_first, rw_last, col_first, col_last) =
                match (r.u32(), r.u32(), r.u32(), r.u32()) {
                    (Ok(a), Ok(b), Ok(c), Ok(d)) => (a, b, c, d),
                    _ => continue,
                };
            if rec_type == BRT_ARR_FMLA {
                // BrtArrFmla has a 1-byte fAlwaysCalc/unused flags field between
                // rfx and formula that BrtShrFmla does not have.
                if r.skip(1).is_err() {
                    continue;
                }
            }
            let (rgce, rgcb) = match read_cell_parsed_formula(&mut r) {
                Some(v) => v,
                None => continue,
            };
            shared_formulas.push((
                rw_first,
                rw_last,
                col_first,
                col_last,
                rgce.to_vec(),
                rgcb.to_vec(),
            ));

            if let Some((arow, acol)) = last_formula_cell {
                if rw_first <= arow && arow <= rw_last && col_first <= acol && acol <= col_last {
                    let formula = decompile_formula(rgce, sheet_names, dn, arow, acol, rgcb);
                    formulas.insert((arow, acol), formula);
                }
            }
        }
    }

    // Resolve shared-formula references (PtgExp) after full pass.
    for (row, col) in pending_exp {
        let matched =
            shared_formulas
                .iter()
                .find(|(rw_first, rw_last, col_first, col_last, _, _)| {
                    *rw_first <= row && row <= *rw_last && *col_first <= col && col <= *col_last
                });
        if let Some((_, _, _, _, s_rgce, s_rgcb)) = matched {
            let formula = decompile_formula(s_rgce, sheet_names, dn, row, col, s_rgcb);
            formulas.insert((row, col), formula);
        }
    }

    formulas
}

/// Parse worksheet cached/constant cell values. Returns
/// `{(row, col): value}`. Row/col are 0-based. Port of
/// `_parse_worksheet_values`.
fn parse_worksheet_values(data: &[u8], sst: &[String]) -> BTreeMap<(u32, u32), CellValue> {
    let mut values: BTreeMap<(u32, u32), CellValue> = BTreeMap::new();
    let mut current_row: u32 = 0;

    for (rec_type, payload) in RecordReader::new(data) {
        if rec_type == BRT_ROW_HDR {
            if payload.len() >= 4 {
                current_row = u32::from_le_bytes([payload[0], payload[1], payload[2], payload[3]]);
            }
            continue;
        }

        if rec_type == BRT_CELL_ISST && payload.len() >= 12 {
            let col = u32::from_le_bytes([payload[0], payload[1], payload[2], payload[3]]);
            let isst =
                u32::from_le_bytes([payload[8], payload[9], payload[10], payload[11]]) as usize;
            let s = sst.get(isst).cloned().unwrap_or_default();
            values.insert((current_row, col), CellValue::Str(s));
        } else if rec_type == BRT_CELL_ST && payload.len() >= 8 {
            let col = u32::from_le_bytes([payload[0], payload[1], payload[2], payload[3]]);
            let mut r = Reader::new(payload);
            if r.skip(8).is_ok() {
                values.insert((current_row, col), CellValue::Str(read_xlwide_from(&mut r)));
            }
        } else if rec_type == BRT_CELL_REAL && payload.len() >= 16 {
            let col = u32::from_le_bytes([payload[0], payload[1], payload[2], payload[3]]);
            let v = f64::from_le_bytes(payload[8..16].try_into().unwrap());
            values.insert((current_row, col), CellValue::Float(v));
        } else if rec_type == BRT_CELL_RK && payload.len() >= 12 {
            let col = u32::from_le_bytes([payload[0], payload[1], payload[2], payload[3]]);
            let rk = u32::from_le_bytes([payload[8], payload[9], payload[10], payload[11]]);
            values.insert((current_row, col), rk_to_number(rk));
        } else if rec_type == BRT_CELL_BOOL && payload.len() >= 9 {
            let col = u32::from_le_bytes([payload[0], payload[1], payload[2], payload[3]]);
            values.insert((current_row, col), CellValue::Bool(payload[8] != 0));
        } else if rec_type == BRT_CELL_ERROR && payload.len() >= 9 {
            let col = u32::from_le_bytes([payload[0], payload[1], payload[2], payload[3]]);
            values.insert(
                (current_row, col),
                CellValue::Str(err_code_str(payload[8]).to_string()),
            );
        } else if is_fmla_rec(rec_type) && payload.len() >= 10 {
            let col = u32::from_le_bytes([payload[0], payload[1], payload[2], payload[3]]);
            let mut r = Reader::new(payload);
            if r.skip(8).is_err() {
                continue;
            }
            match rec_type {
                BRT_FMLA_STRING => {
                    values.insert((current_row, col), CellValue::Str(read_xlwide_from(&mut r)));
                }
                BRT_FMLA_NUM => {
                    let n = 8.min(r.remaining());
                    if let Ok(raw) = r.take_bytes(n) {
                        if raw.len() == 8 {
                            let v = f64::from_le_bytes(raw.try_into().unwrap());
                            values.insert((current_row, col), CellValue::Float(v));
                        }
                    }
                }
                BRT_FMLA_BOOL => {
                    let n = 1.min(r.remaining());
                    if let Ok(b) = r.take_bytes(n) {
                        if let Some(&byte) = b.first() {
                            values.insert((current_row, col), CellValue::Bool(byte != 0));
                        }
                    }
                }
                BRT_FMLA_ERROR => {
                    let n = 1.min(r.remaining());
                    if let Ok(b) = r.take_bytes(n) {
                        if let Some(&byte) = b.first() {
                            values.insert(
                                (current_row, col),
                                CellValue::Str(err_code_str(byte).to_string()),
                            );
                        }
                    }
                }
                _ => {}
            }
        }
    }

    values
}

// ---------------------------------------------------------------------------
// AutoFilter parsing — port of `_parse_custom_filter` / `_parse_worksheet_filters`.
// ---------------------------------------------------------------------------

fn parse_custom_filter(payload: &[u8]) -> Option<Value> {
    if payload.len() < 10 {
        return None;
    }
    let vts = payload[0];
    let grbit_sgn = payload[1];
    let operator = filter_op_str(grbit_sgn);
    let union = &payload[2..10];
    let value: Value = match vts {
        0x04 => {
            let v = f64::from_le_bytes(union.try_into().unwrap());
            Value::from(v)
        }
        0x08 => Value::Bool(union[0] != 0),
        0x06 => {
            let mut r = Reader::new(&payload[10..]);
            Value::String(read_xlwide_from(&mut r))
        }
        0x0C | 0x0E => Value::Null,
        _ => return None,
    };
    let mut obj = Map::new();
    obj.insert("operator".to_string(), Value::String(operator));
    obj.insert("value".to_string(), value);
    Some(Value::Object(obj))
}

/// Parse AutoFilter records from a worksheet bin part. Returns a filter
/// info object, or `None` if no AutoFilter is present. Port of
/// `_parse_worksheet_filters`.
fn parse_worksheet_filters(data: &[u8]) -> Option<FilterInfo> {
    let mut result: Option<Map<String, Value>> = None;
    let mut current_col: Option<Map<String, Value>> = None;
    let mut in_pivot_afilter = false;
    let mut sx_depth: u32 = 0;

    for (rec_type, payload) in RecordReader::new(data) {
        if rec_type == BRT_BEGIN_SX_FILTER {
            sx_depth += 1;
        } else if rec_type == BRT_END_SX_FILTER {
            sx_depth = sx_depth.saturating_sub(1);
        } else if rec_type == BRT_BEGIN_AFILTER {
            if sx_depth > 0 {
                in_pivot_afilter = true;
                continue;
            }
            if payload.len() < 16 {
                continue;
            }
            let rw_first = u32::from_le_bytes([payload[0], payload[1], payload[2], payload[3]]);
            let rw_last = u32::from_le_bytes([payload[4], payload[5], payload[6], payload[7]]);
            let col_first = u32::from_le_bytes([payload[8], payload[9], payload[10], payload[11]]);
            let col_last = u32::from_le_bytes([payload[12], payload[13], payload[14], payload[15]]);
            let mut range = Map::new();
            range.insert(
                "top_left".to_string(),
                Value::String(format!(
                    "{}{}",
                    common::col_to_letter(col_first),
                    rw_first + 1
                )),
            );
            range.insert(
                "bottom_right".to_string(),
                Value::String(format!(
                    "{}{}",
                    common::col_to_letter(col_last),
                    rw_last + 1
                )),
            );
            let mut obj = Map::new();
            obj.insert("range".to_string(), Value::Object(range));
            obj.insert("columns".to_string(), Value::Array(Vec::new()));
            result = Some(obj);
        } else if rec_type == BRT_END_AFILTER {
            in_pivot_afilter = false;
        } else if rec_type == BRT_BEGIN_FILTER_COLUMN && !in_pivot_afilter {
            if result.is_none() || payload.len() < 4 {
                continue;
            }
            let dw_col = u32::from_le_bytes([payload[0], payload[1], payload[2], payload[3]]);
            let mut col = Map::new();
            col.insert("column_index".to_string(), Value::from(dw_col));
            col.insert("filters".to_string(), Value::Array(Vec::new()));
            current_col = Some(col);
        } else if rec_type == BRT_END_FILTER_COLUMN && !in_pivot_afilter {
            if let (Some(res), Some(col)) = (result.as_mut(), current_col.take()) {
                if let Some(Value::Array(cols)) = res.get_mut("columns") {
                    cols.push(Value::Object(col));
                }
            }
        } else if rec_type == BRT_BEGIN_CUSTOM_FILTERS && !in_pivot_afilter {
            if current_col.is_none() || payload.len() < 4 {
                continue;
            }
            let f_and = u32::from_le_bytes([payload[0], payload[1], payload[2], payload[3]]);
            let logic = if f_and != 0 { "or" } else { "and" };
            let mut cf = Map::new();
            cf.insert("logic".to_string(), Value::String(logic.to_string()));
            cf.insert("criteria".to_string(), Value::Array(Vec::new()));
            if let Some(col) = current_col.as_mut() {
                col.insert("custom_filters".to_string(), Value::Object(cf));
            }
        } else if rec_type == BRT_CUSTOM_FILTER && !in_pivot_afilter {
            let col = match current_col.as_mut() {
                Some(c) => c,
                None => continue,
            };
            if let Some(crit) = parse_custom_filter(payload) {
                match col.get_mut("custom_filters") {
                    Some(Value::Object(cf)) => {
                        if let Some(Value::Array(criteria)) = cf.get_mut("criteria") {
                            criteria.push(crit);
                        }
                    }
                    _ => {
                        let mut cf = Map::new();
                        cf.insert("logic".to_string(), Value::String("and".to_string()));
                        cf.insert("criteria".to_string(), Value::Array(vec![crit]));
                        col.insert("custom_filters".to_string(), Value::Object(cf));
                    }
                }
            }
        } else if rec_type == BRT_FILTER && !in_pivot_afilter {
            let col = match current_col.as_mut() {
                Some(c) => c,
                None => continue,
            };
            let mut r = Reader::new(payload);
            let val = read_xlwide_from(&mut r);
            if let Some(Value::Array(filters)) = col.get_mut("filters") {
                filters.push(Value::String(val));
            }
        }
    }

    result.map(Value::Object)
}

// ---------------------------------------------------------------------------
// PivotTable / PivotCache parsing
// ---------------------------------------------------------------------------

/// Parse key metadata from a PivotTable part. Port of
/// `_parse_pivot_table_part`.
fn parse_pivot_table_part(data: &[u8]) -> Map<String, Value> {
    let mut meta = Map::new();
    meta.insert("name".to_string(), Value::Null);
    meta.insert("cache_id".to_string(), Value::Null);
    meta.insert("data_caption".to_string(), Value::Null);
    meta.insert("location".to_string(), Value::Null);
    meta.insert("pivot_fields".to_string(), Value::from(0));
    meta.insert("pivot_items".to_string(), Value::from(0));
    meta.insert("sx_filters".to_string(), Value::Array(Vec::new()));
    meta.insert("field_filters".to_string(), Value::Array(Vec::new()));

    let mut current_sx_filter: Option<Map<String, Value>> = None;
    let mut current_field_index: i64 = -1;
    let mut current_hidden: Vec<i64> = Vec::new();

    macro_rules! flush_field {
        () => {
            if !current_hidden.is_empty() {
                let mut ff = Map::new();
                ff.insert("field_index".to_string(), Value::from(current_field_index));
                ff.insert(
                    "hidden_indices".to_string(),
                    Value::Array(current_hidden.iter().map(|&v| Value::from(v)).collect()),
                );
                if let Some(Value::Array(arr)) = meta.get_mut("field_filters") {
                    arr.push(Value::Object(ff));
                }
            }
        };
    }

    for (rec_type, payload) in RecordReader::new(data) {
        if rec_type == BRT_BEGIN_SX_VIEW && payload.len() >= 32 {
            let id_cache = u32::from_le_bytes([payload[28], payload[29], payload[30], payload[31]]);
            meta.insert("cache_id".to_string(), Value::from(id_cache));
            let mut r = Reader::new(payload);
            if r.skip(32).is_ok() {
                let name = read_xlwide_from(&mut r);
                meta.insert("name".to_string(), Value::String(name));
                let data_caption = read_xlwide_from(&mut r);
                if !data_caption.is_empty() {
                    meta.insert("data_caption".to_string(), Value::String(data_caption));
                }
            }
        } else if rec_type == BRT_BEGIN_SX_LOCATION && payload.len() >= 36 {
            let rw_first = u32::from_le_bytes(payload[0..4].try_into().unwrap());
            let rw_last = u32::from_le_bytes(payload[4..8].try_into().unwrap());
            let col_first = u32::from_le_bytes(payload[8..12].try_into().unwrap());
            let col_last = u32::from_le_bytes(payload[12..16].try_into().unwrap());
            let rw_first_head = u32::from_le_bytes(payload[16..20].try_into().unwrap());
            let rw_first_data = u32::from_le_bytes(payload[20..24].try_into().unwrap());
            let col_first_data = u32::from_le_bytes(payload[24..28].try_into().unwrap());
            let crw_page = u32::from_le_bytes(payload[28..32].try_into().unwrap());
            let ccol_page = u32::from_le_bytes(payload[32..36].try_into().unwrap());

            let mut rfx_geom = Map::new();
            rfx_geom.insert(
                "top_left".to_string(),
                Value::String(format!(
                    "{}{}",
                    common::col_to_letter(col_first),
                    rw_first + 1
                )),
            );
            rfx_geom.insert(
                "bottom_right".to_string(),
                Value::String(format!(
                    "{}{}",
                    common::col_to_letter(col_last),
                    rw_last + 1
                )),
            );
            let mut location = Map::new();
            location.insert("rfx_geom".to_string(), Value::Object(rfx_geom));
            location.insert("rw_first_head".to_string(), Value::from(rw_first_head + 1));
            location.insert("rw_first_data".to_string(), Value::from(rw_first_data + 1));
            location.insert(
                "col_first_data".to_string(),
                Value::String(common::col_to_letter(col_first_data)),
            );
            location.insert("page_rows".to_string(), Value::from(crw_page));
            location.insert("page_cols".to_string(), Value::from(ccol_page));
            meta.insert("location".to_string(), Value::Object(location));
        } else if rec_type == BRT_BEGIN_SXVD {
            flush_field!();
            current_field_index += 1;
            current_hidden.clear();
            if let Some(Value::Number(n)) = meta.get("pivot_fields") {
                let v = n.as_i64().unwrap_or(0) + 1;
                meta.insert("pivot_fields".to_string(), Value::from(v));
            }
        } else if rec_type == BRT_BEGIN_SXVI {
            if let Some(Value::Number(n)) = meta.get("pivot_items") {
                let v = n.as_i64().unwrap_or(0) + 1;
                meta.insert("pivot_items".to_string(), Value::from(v));
            }
            if payload.len() >= 7 {
                let itmtype = payload[0];
                let grbit = u16::from_le_bytes([payload[1], payload[2]]);
                let icache = i32::from_le_bytes(payload[3..7].try_into().unwrap());
                if itmtype == 0 && (grbit & 0x0001 != 0) && icache >= 0 {
                    current_hidden.push(icache as i64);
                }
            }
        } else if rec_type == BRT_BEGIN_SX_FILTER && payload.len() >= 12 {
            let i_field = i32::from_le_bytes(payload[0..4].try_into().unwrap());
            let sxft = i32::from_le_bytes(payload[8..12].try_into().unwrap());
            let mut sf = Map::new();
            sf.insert("field_index".to_string(), Value::from(i_field as i64));
            sf.insert("filter_type".to_string(), Value::from(sxft as i64));
            sf.insert("criteria".to_string(), Value::Array(Vec::new()));
            current_sx_filter = Some(sf);
        } else if rec_type == BRT_CUSTOM_FILTER {
            if let Some(sf) = current_sx_filter.as_mut() {
                if let Some(crit) = parse_custom_filter(payload) {
                    if let Some(Value::Array(criteria)) = sf.get_mut("criteria") {
                        criteria.push(crit);
                    }
                }
            }
        } else if rec_type == BRT_END_SX_FILTER {
            if let Some(sf) = current_sx_filter.take() {
                if let Some(Value::Array(arr)) = meta.get_mut("sx_filters") {
                    arr.push(Value::Object(sf));
                }
            }
        }
    }

    flush_field!();
    meta
}

/// Parse a `pivotCacheDefinitionN.bin` part. Returns, in cache-field order,
/// `[{"name": str, "shared_items": [str | float, ...]}, ...]`. Port of
/// `_parse_pivot_cache_fields`.
fn parse_pivot_cache_fields(data: &[u8]) -> Vec<Value> {
    let mut fields: Vec<Map<String, Value>> = Vec::new();

    for (rec_type, payload) in RecordReader::new(data) {
        if rec_type == BRT_BEGIN_PCDFIELD {
            let mut name = String::new();
            if payload.len() >= 20 {
                let mut r = Reader::new(payload);
                if r.skip(20).is_ok() {
                    name = read_xlwide_from(&mut r);
                }
            }
            let mut f = Map::new();
            f.insert("name".to_string(), Value::String(name));
            f.insert("shared_items".to_string(), Value::Array(Vec::new()));
            fields.push(f);
        } else if rec_type == BRT_BEGIN_PCDIRUN && !fields.is_empty() && payload.len() >= 6 {
            let md_sxoper = u16::from_le_bytes([payload[0], payload[1]]);
            let citems = u32::from_le_bytes(payload[2..6].try_into().unwrap()) as usize;
            let mut r = Reader::new(payload);
            if r.skip(6).is_err() {
                continue;
            }
            let mut items: Vec<Value> = Vec::new();
            if md_sxoper == 0x0002 {
                for _ in 0..citems {
                    items.push(Value::String(read_xlwide_from(&mut r)));
                }
            } else if md_sxoper == 0x0001 {
                for _ in 0..citems {
                    let n = 8.min(r.remaining());
                    let chunk = match r.take_bytes(n) {
                        Ok(c) if c.len() == 8 => c,
                        _ => break,
                    };
                    items.push(Value::from(f64::from_le_bytes(chunk.try_into().unwrap())));
                }
            }
            if let Some(last) = fields.last_mut() {
                if let Some(Value::Array(shared)) = last.get_mut("shared_items") {
                    shared.extend(items);
                }
            }
        } else if rec_type == BRT_PCDI_STRING && !fields.is_empty() {
            let mut r = Reader::new(payload);
            let s = read_xlwide_from(&mut r);
            if let Some(last) = fields.last_mut() {
                if let Some(Value::Array(shared)) = last.get_mut("shared_items") {
                    shared.push(Value::String(s));
                }
            }
        }
    }

    fields.into_iter().map(Value::Object).collect()
}

// ---------------------------------------------------------------------------
// Top-level parse: open the ZIP, orchestrate the reads above.
// ---------------------------------------------------------------------------

type ZipArchive = zip::ZipArchive<std::io::BufReader<std::fs::File>>;

/// Read a part by name, with the same case-insensitive fallback as
/// `XlsbWorkbook._read_part`.
fn read_part(zf: &mut ZipArchive, path: &str) -> Result<Vec<u8>> {
    let norm = path.replace('\\', "/");
    let norm = norm.trim_start_matches('/');
    if let Ok(mut f) = zf.by_name(norm) {
        let mut buf = Vec::with_capacity(f.size() as usize);
        f.read_to_end(&mut buf)?;
        return Ok(buf);
    }
    let lo = norm.to_lowercase();
    let found = zf
        .file_names()
        .find(|n| n.to_lowercase() == lo)
        .map(|s| s.to_string());
    match found {
        Some(n) => {
            let mut f = zf.by_name(&n)?;
            let mut buf = Vec::with_capacity(f.size() as usize);
            f.read_to_end(&mut buf)?;
            Ok(buf)
        }
        None => Err(XlsbError::Parse(format!("Not found in XLSB zip: {path}"))),
    }
}

/// Discover PivotTable parts through worksheet relationships and parse
/// them, linked to sheets where possible. Port of
/// `XlsbWorkbook.iter_pivot_tables`.
fn parse_pivot_tables(
    zf: &mut ZipArchive,
    sheets: &[(String, String)],
    paths: &HashMap<String, String>,
) -> Vec<PivotTable> {
    // Map sheet part path -> sheet name, preserving first-insertion order
    // (mirrors Python dict semantics: `sheet_by_path[zpath] = sheet_name`).
    let mut order: Vec<String> = Vec::new();
    let mut sheet_by_path: HashMap<String, String> = HashMap::new();
    for (name, rel_id) in sheets {
        if let Some(zp) = paths.get(rel_id) {
            if !sheet_by_path.contains_key(zp) {
                order.push(zp.clone());
            }
            sheet_by_path.insert(zp.clone(), name.clone());
        }
    }

    let mut seen_parts: HashSet<String> = HashSet::new();
    let mut cache_fields_by_part: HashMap<String, Vec<Value>> = HashMap::new();
    let mut pivots: Vec<PivotTable> = Vec::new();

    for sheet_path in &order {
        let sheet_name = match sheet_by_path.get(sheet_path) {
            Some(n) => n.clone(),
            None => continue,
        };
        let rel_path = format!(
            "{}/_rels/{}.rels",
            dirname(sheet_path),
            basename(sheet_path)
        );
        let rels_bytes = match read_part(zf, &rel_path) {
            Ok(b) => b,
            Err(_) => continue,
        };
        let rels = read_rels(&rels_bytes);

        for (_id, target) in &rels {
            let resolved = resolve_rel_target(sheet_path, target);
            if !resolved.contains("/pivotTables/") {
                continue;
            }
            if seen_parts.contains(&resolved) {
                continue;
            }
            seen_parts.insert(resolved.clone());
            let pdata = match read_part(zf, &resolved) {
                Ok(b) => b,
                Err(_) => continue,
            };
            let mut meta = parse_pivot_table_part(&pdata);
            meta.insert("sheet".to_string(), Value::String(sheet_name.clone()));
            meta.insert("part".to_string(), Value::String(resolved.clone()));

            let p_rel_path = format!("{}/_rels/{}.rels", dirname(&resolved), basename(&resolved));
            let p_rels = read_part(zf, &p_rel_path)
                .map(|b| read_rels(&b))
                .unwrap_or_default();
            let mut cache_def: Option<String> = None;
            for (_id, t) in &p_rels {
                let rr = resolve_rel_target(&resolved, t);
                if rr.contains("/pivotCache/pivotCacheDefinition") {
                    cache_def = Some(rr);
                    break;
                }
            }
            if let Some(cd) = cache_def {
                meta.insert(
                    "pivot_cache_definition".to_string(),
                    Value::String(cd.clone()),
                );
                if !cache_fields_by_part.contains_key(&cd) {
                    let fields = match read_part(zf, &cd) {
                        Ok(cdata) => parse_pivot_cache_fields(&cdata),
                        Err(_) => Vec::new(),
                    };
                    cache_fields_by_part.insert(cd.clone(), fields);
                }
                meta.insert(
                    "cache_fields".to_string(),
                    Value::Array(cache_fields_by_part[&cd].clone()),
                );
            }
            pivots.push(Value::Object(meta));
        }
    }

    pivots
}

/// Parse a `.xlsb` file at `path` into a [`WorkbookData`].
pub fn parse_xlsb(path: &Path) -> Result<WorkbookData> {
    let file = std::fs::File::open(path)?;
    let mut zf = zip::ZipArchive::new(std::io::BufReader::new(file))?;

    let wb_data = read_part(&mut zf, "xl/workbook.bin")?;
    let sheets = read_workbook(&wb_data);
    let dnames = read_defined_names(&wb_data);

    let wb_rels = read_part(&mut zf, "xl/_rels/workbook.bin.rels")
        .map(|b| read_rels(&b))
        .unwrap_or_default();

    let mut paths: HashMap<String, String> = HashMap::new();
    for (_name, rel_id) in &sheets {
        if let Some(target) = rels_get(&wb_rels, rel_id) {
            if !target.is_empty() {
                let full = if target.starts_with("xl/") {
                    target.to_string()
                } else {
                    format!("xl/{target}")
                };
                paths.insert(rel_id.clone(), full.replace("//", "/"));
            }
        }
    }

    let sst = read_part(&mut zf, "xl/sharedStrings.bin")
        .map(|b| read_sst(&b))
        .unwrap_or_default();

    let sheet_names: Vec<String> = sheets.iter().map(|(n, _)| n.clone()).collect();

    let mut formulas = Vec::with_capacity(sheets.len());
    let mut values = Vec::with_capacity(sheets.len());
    let mut filters = Vec::with_capacity(sheets.len());

    for (name, rel_id) in &sheets {
        let ws_bytes = match paths.get(rel_id) {
            Some(zp) => read_part(&mut zf, zp).ok(),
            None => None,
        };
        match &ws_bytes {
            Some(data) => {
                formulas.push((name.clone(), parse_worksheet(data, &sheet_names, &dnames)));
                values.push((name.clone(), parse_worksheet_values(data, &sst)));
                filters.push((name.clone(), parse_worksheet_filters(data)));
            }
            None => {
                formulas.push((name.clone(), BTreeMap::new()));
                values.push((name.clone(), BTreeMap::new()));
                filters.push((name.clone(), None));
            }
        }
    }

    let pivots = parse_pivot_tables(&mut zf, &sheets, &paths);

    Ok(WorkbookData {
        sheet_names,
        formulas,
        values,
        filters,
        pivots,
        vba: Vec::new(),
    })
}

// ---------------------------------------------------------------------------
// PyO3 pyclass wrapper
// ---------------------------------------------------------------------------

/// Mirrors `xlsb_reader._reader.XlsbWorkbook`.
///
/// Holds the source file path and lazily parses it into a [`WorkbookData`]
/// on first use (cached for the lifetime of the object — note this is a
/// deliberate behavioural improvement over the pure-Python reader, which
/// re-parses from scratch on every `iter_*()` call).
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

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

#[cfg(test)]
mod tests {
    use super::*;

    fn dc(rgce: &[u8]) -> String {
        let sheets: Vec<String> = vec![];
        let dn: BTreeMap<u32, String> = BTreeMap::new();
        let mut d = Decompiler::new(rgce, &sheets, &dn, Some(0), Some(0), &[]);
        d.decompile().unwrap()
    }

    #[test]
    fn decompile_simple_add() {
        // PtgInt(1) PtgInt(2) PtgAdd
        let rgce = [
            PTG_INT, 0x01, 0x00, // 1
            PTG_INT, 0x02, 0x00, // 2
            PTG_ADD,
        ];
        assert_eq!(dc(&rgce), "1+2");
    }

    #[test]
    fn decompile_num_literal() {
        let mut rgce = vec![PTG_NUM];
        rgce.extend_from_slice(&3.5f64.to_le_bytes());
        assert_eq!(dc(&rgce), "3.5");
    }

    #[test]
    fn decompile_string_literal() {
        let mut rgce = vec![PTG_STR];
        rgce.extend_from_slice(&2u16.to_le_bytes()); // cch=2
        rgce.extend_from_slice(
            "hi".encode_utf16()
                .flat_map(|u| u.to_le_bytes())
                .collect::<Vec<u8>>()
                .as_slice(),
        );
        assert_eq!(dc(&rgce), "\"hi\"");
    }

    #[test]
    fn decompile_func_sum_fixed_arity() {
        // PtgInt(1) PtgInt(2) PtgFuncVar(cparams=2, iftab=SUM=0x0004)
        let rgce = [
            PTG_INT,
            0x01,
            0x00,
            PTG_INT,
            0x02,
            0x00,
            0x20 | B_FUNC_VAR,
            0x02,
            0x04,
            0x00,
        ];
        assert_eq!(dc(&rgce), "SUM(1,2)");
    }

    #[test]
    fn decompile_ref() {
        // PtgRefV: base=REF, class VAL (0x20 offset) = 0x44
        let mut rgce = vec![0x20 | B_REF];
        rgce.extend_from_slice(&0u32.to_le_bytes()); // row 0
        rgce.extend_from_slice(&0u16.to_le_bytes()); // col 0, absolute
        assert_eq!(dc(&rgce), "$A$1");
    }

    #[test]
    fn decompile_unknown_ptg_breaks() {
        let rgce = [0x1B]; // unused exact-byte value
        assert_eq!(dc(&rgce), "?ptg0x1b");
    }

    #[test]
    fn fmt_num_integral() {
        assert_eq!(fmt_num(5.0).unwrap(), "5");
        assert_eq!(fmt_num(-5.0).unwrap(), "-5");
        assert_eq!(fmt_num(0.0).unwrap(), "0");
    }

    #[test]
    fn fmt_num_fraction() {
        assert_eq!(fmt_num(3.5).unwrap(), "3.5");
        assert_eq!(fmt_num(0.1).unwrap(), "0.1");
    }

    #[test]
    fn fmt_num_large_scientific() {
        // decpt = 17 > 16 -> scientific
        assert_eq!(fmt_num(1e16).unwrap(), "1e+16");
    }

    #[test]
    fn fmt_num_small_scientific() {
        // decpt <= -4 -> scientific
        assert_eq!(fmt_num(0.00001).unwrap(), "1e-05");
    }

    #[test]
    fn rk_number_integer() {
        // f_int set, fx100 clear: raw = (5 << 2) | 0x2
        let raw = (5i64 << 2) as u32 | 0x2;
        assert_eq!(rk_to_number(raw), CellValue::Int(5));
    }

    #[test]
    fn rk_number_negative_integer() {
        let num30 = (-5i64) & 0x3FFF_FFFF; // 30-bit two's complement of -5
        let raw = (num30 as u32) << 2 | 0x2;
        assert_eq!(rk_to_number(raw), CellValue::Int(-5));
    }

    #[test]
    fn rk_number_fx100() {
        // f_int set, fx100 set: value 500 / 100 = 5.0
        let raw = (500i64 << 2) as u32 | 0x3;
        assert_eq!(rk_to_number(raw), CellValue::Float(5.0));
    }

    #[test]
    fn record_reader_iterates() {
        // type=0x02 (single byte, no high bit), size=0x03, data=[0xAA,0xBB,0xCC]
        let data = [0x02, 0x03, 0xAA, 0xBB, 0xCC];
        let recs: Vec<(u32, &[u8])> = RecordReader::new(&data).collect();
        assert_eq!(recs, vec![(0x02, &[0xAA, 0xBB, 0xCC][..])]);
    }

    #[test]
    fn record_reader_two_byte_type() {
        // type = 0x0118 (BrtBeginSXView): low byte 0x98 (0x18|0x80), high byte 0x02
        let data = [0x98, 0x02, 0x00];
        let recs: Vec<(u32, &[u8])> = RecordReader::new(&data).collect();
        assert_eq!(recs, vec![(0x0118, &[][..])]);
    }

    #[test]
    fn parse_worksheet_filters_none_without_afilter() {
        assert!(parse_worksheet_filters(&[]).is_none());
    }

    #[test]
    fn parse_worksheet_filters_basic_range() {
        // BrtBeginAFilter (0x00A1): rw_first=0, rw_last=9, col_first=0, col_last=2
        let mut payload = Vec::new();
        payload.extend_from_slice(&0u32.to_le_bytes());
        payload.extend_from_slice(&9u32.to_le_bytes());
        payload.extend_from_slice(&0u32.to_le_bytes());
        payload.extend_from_slice(&2u32.to_le_bytes());
        let mut data = Vec::new();
        // 0x00A1's low byte has the 0x80 bit set, so BIFF12's varint record-type
        // encoding needs two bytes here: [0xA1, 0x01] (see
        // `record_reader_two_byte_type` above for the same encoding rule).
        data.push(0xA1);
        data.push(0x01);
        data.push(payload.len() as u8); // varint size
        data.extend_from_slice(&payload);

        let info = parse_worksheet_filters(&data).expect("expected filter info");
        let range = &info["range"];
        assert_eq!(range["top_left"], "A1");
        assert_eq!(range["bottom_right"], "C10");
        assert_eq!(info["columns"].as_array().unwrap().len(), 0);
    }

    #[test]
    fn parse_custom_filter_real() {
        let mut payload = vec![0x04u8, 0x04]; // vts=Real, grbitSgn=0x04 ">"
        payload.extend_from_slice(&1.5f64.to_le_bytes());
        let crit = parse_custom_filter(&payload).unwrap();
        assert_eq!(crit["operator"], ">");
        assert_eq!(crit["value"], 1.5);
    }

    #[test]
    fn col_letter_and_cell_ref() {
        assert_eq!(cell_ref(0, 0), "A1");
        assert_eq!(cell_ref(9, 26), "AA10");
    }

    #[test]
    fn sheet_prefix_quotes_special_chars() {
        let sheets = vec!["Sheet 1".to_string(), "Plain".to_string()];
        assert_eq!(sheet_prefix(&sheets, 0), "'Sheet 1'!");
        assert_eq!(sheet_prefix(&sheets, 1), "Plain!");
        assert_eq!(sheet_prefix(&sheets, 5), "[sheet5]!");
    }
}
