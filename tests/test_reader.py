import pytest
from xlsb_reader import XlsbWorkbook, XlsxWorkbook
from xlsb_reader._vba_reader import _decompress
import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))
from create_test_xlsm import _compress
from pathlib import Path

TEST_WORKBOOK = (
    Path(__file__).resolve().parents[1]
    / "test-data"
    / "Finance_Ledger_100_Unique_Functions.xlsb"
)


@pytest.fixture(scope="module")
def workbook():
    with XlsbWorkbook(TEST_WORKBOOK) as wb:
        yield wb


def test_formula_extraction(workbook):
    formulas_by_sheet = {
        sheet: formulas for sheet, formulas in workbook.iter_formulas()
    }
    function_lab = formulas_by_sheet["Function_Lab"]
    ledger = formulas_by_sheet["Ledger"]

    assert function_lab[(17, 2)] == "=ACOSH(2)"  # C18
    assert function_lab[(18, 2)] == "=@ACOT(1)"  # C19
    assert function_lab[(19, 2)] == "=@ACOTH(2)"  # C20
    assert function_lab[(20, 2)] == "=ADDRESS(3,1)"  # C21
    assert (
        function_lab[(31, 2)] == '=AVERAGEIFS(Ledger!J186:J241,Ledger!J186:J241,">100")'
    )  # C32

    assert ledger[(1, 10)] == "=MAX(J2,0)"  # K2
    assert ledger[(1, 11)] == "=ABS(MIN(J2,0))"  # L2


def test_value_extraction(workbook):
    values_by_sheet = {sheet: values for sheet, values in workbook.iter_values()}
    function_inventory = values_by_sheet["Function_Inventory"]
    pivot_sheet = values_by_sheet["Pivot_Tables"]

    assert function_inventory[(5, 0)] == "ABS"  # A6
    assert function_inventory[(94, 0)] == "PROPER"  # A95

    assert pivot_sheet[(13, 0)] == "Row Labels"  # A14
    assert pivot_sheet[(20, 4)] == 27.0  # E21 — Grand Total (6 categories × 3 currencies)
    assert pivot_sheet[(27, 1)] == pytest.approx(59719.2104, rel=1e-9)  # B28 Finance NetGBP
    assert pivot_sheet[(44, 1)] == 169.0  # B45 — Grand Total count of TxnID


def test_pivot_extraction(workbook):
    pivots = list(workbook.iter_pivot_tables())
    assert len(pivots) == 3

    by_name = {p["name"]: p for p in pivots}
    assert set(by_name) == {"PivotTable1", "PivotTable2", "PivotTable3"}

    p1 = by_name["PivotTable1"]
    assert p1["sheet"] == "Pivot_Tables"
    assert p1["pivot_fields"] == 13
    assert p1["pivot_items"] == 13
    assert p1["pivot_cache_definition"] == "xl/pivotCache/pivotCacheDefinition1.bin"
    assert p1["location"]["rfx_geom"]["top_left"] == "A13"
    assert p1["location"]["rfx_geom"]["bottom_right"] == "E23"

    p2 = by_name["PivotTable2"]
    assert p2["location"]["rfx_geom"]["top_left"] == "A27"
    assert p2["location"]["rfx_geom"]["bottom_right"] == "B34"

    p3 = by_name["PivotTable3"]
    assert p3["location"]["rfx_geom"]["top_left"] == "A37"
    assert p3["location"]["rfx_geom"]["bottom_right"] == "B45"


# ---------------------------------------------------------------------------
# XLSX tests
# ---------------------------------------------------------------------------

XLSX_WORKBOOK = (
    Path(__file__).resolve().parents[1]
    / "test-data"
    / "Finance_Ledger_100_Unique_Functions.xlsx"
)


@pytest.fixture(scope="module")
def xlsx_workbook():
    with XlsxWorkbook(XLSX_WORKBOOK) as wb:
        yield wb


def test_xlsx_sheet_names(xlsx_workbook):
    assert xlsx_workbook.sheet_names == [
        "Ledger", "Function_Lab", "Function_Inventory", "Pivot_Tables"
    ]


def test_xlsx_formulas(xlsx_workbook):
    formulas_by_sheet = {s: f for s, f in xlsx_workbook.iter_formulas()}
    ledger = formulas_by_sheet["Ledger"]
    fl = formulas_by_sheet["Function_Lab"]

    # Anchor shared formulas in Ledger
    assert ledger[(1, 10)] == "=MAX(J2,0)"   # K2
    assert ledger[(1, 11)] == "=ABS(MIN(J2,0))"  # L2
    assert ledger[(1, 12)] == "=J2*I2"        # M2

    # Function_Lab has 100 unique formulas
    assert len(fl) == 100
    # UDF prefix stripped
    assert fl[(18, 2)] == "=ACOT(1)"          # C19 — was _xludf.ACOT
    # Cross-sheet reference preserved
    assert fl[(31, 2)] == '=AVERAGEIFS(Ledger!J186:J241,Ledger!J186:J241,">100")'


def test_xlsx_shared_formula_expansion(xlsx_workbook):
    formulas_by_sheet = {s: f for s, f in xlsx_workbook.iter_formulas()}
    ledger = formulas_by_sheet["Ledger"]

    # Row shifted by 1 relative to anchor K2/L2
    assert ledger[(2, 10)] == "=MAX(J3,0)"       # K3
    assert ledger[(2, 11)] == "=ABS(MIN(J3,0))"  # L3
    assert ledger[(2, 12)] == "=J3*I3"            # M3

    # Mid-range rows also shift correctly
    assert ledger[(5, 10)] == "=MAX(J6,0)"        # K6


def test_xlsx_values(xlsx_workbook):
    values_by_sheet = {s: v for s, v in xlsx_workbook.iter_values()}
    ledger = values_by_sheet["Ledger"]
    fi = values_by_sheet["Function_Inventory"]

    # Numeric value in Ledger
    assert ledger[(1, 9)] == pytest.approx(-455.88)  # J2

    # Function_Inventory text values
    assert fi[(5, 0)] == "ABS"    # A6


def test_xlsx_pivot_tables(xlsx_workbook):
    pivots = list(xlsx_workbook.iter_pivot_tables())
    assert len(pivots) == 3

    by_name = {p["name"]: p for p in pivots}
    assert set(by_name) == {"PivotTable1", "PivotTable2", "PivotTable3"}

    for p in pivots:
        assert p["sheet"] == "Pivot_Tables"
        assert p["pivot_fields"] == 13
        assert p["cache_id"] == 10
        assert p["pivot_cache_definition"] == "xl/pivotCache/pivotCacheDefinition1.xml"

    p1 = by_name["PivotTable1"]
    assert p1["location"]["rfx_geom"]["top_left"] == "A13"
    assert p1["location"]["rfx_geom"]["bottom_right"] == "E21"

    p2 = by_name["PivotTable2"]
    assert p2["location"]["rfx_geom"]["top_left"] == "A27"
    assert p2["location"]["rfx_geom"]["bottom_right"] == "B34"

    p3 = by_name["PivotTable3"]
    assert p3["location"]["rfx_geom"]["top_left"] == "A37"
    assert p3["location"]["rfx_geom"]["bottom_right"] == "B45"


def test_xlsx_filters(xlsx_workbook):
    filters = list(xlsx_workbook.iter_filters())
    assert len(filters) == 1

    f = filters[0]
    assert f["sheet"] == "Ledger"
    assert f["ref"] == "A1:M241"

    cols = f["columns"]
    assert len(cols) == 1
    col = cols[0]
    assert col["col_id"] == 12
    assert col["type"] == "custom"
    assert col["conditions"] == [{"operator": "greaterThan", "val": "1"}]


# ---------------------------------------------------------------------------
# VBA reader unit tests
# ---------------------------------------------------------------------------

def test_ovba_compress_decompress_roundtrip():
    """Compress then decompress must return original bytes."""
    for original in [
        b"Hello, World!",
        b"Sub Foo()\n    MsgBox 42\nEnd Sub\n",
        b"\x00" * 200,
        b"abc" * 500,   # repetitive → exercises back-refs on decompression side
        b"",
    ]:
        compressed = _compress(original)
        assert compressed[0] == 0x01, "Missing SignatureByte"
        recovered = _decompress(compressed)
        # For empty input we still get a chunk written; decompressed may have trailing zeros
        assert recovered[: len(original)] == original


# ---------------------------------------------------------------------------
# VBA .xlsm tests
# ---------------------------------------------------------------------------

XLSM_WORKBOOK = (
    Path(__file__).resolve().parents[1]
    / "test-data"
    / "Finance_Ledger_100_Unique_Functions.xlsm"
)


@pytest.fixture(scope="module")
def xlsm_workbook():
    with XlsxWorkbook(XLSM_WORKBOOK) as wb:
        yield wb


def test_xlsm_vba_modules_found(xlsm_workbook):
    mods = xlsm_workbook.iter_vba_modules()
    print(set(mods.keys()))
    assert set(mods.keys()) == {"ThisWorkbook", "Sheet1", "Sheet2", "Sheet3", "Sheet4", "Sheet5", "modComplex", "modLog", "modMain", "modMedium", "modSimple"}


def test_xlsm_vba_modLog_content(xlsm_workbook):
    mods = xlsm_workbook.iter_vba_modules()
    src1 = mods["modLog"]

    assert 'Public Sub LogEvent(ByVal source As String, ByVal action As String, ByVal details As String)' in src1
    assert 'Dim ws As Worksheet' in src1
    assert 'nr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1' in src1


def test_xlsm_vba_main_module_content(xlsm_workbook):
    mods = xlsm_workbook.iter_vba_modules()
    src2 = mods["modMain"]
    assert 'Application.ScreenUpdating = False' in src2
    assert 'Application.Calculation = xlCalculationManual' in src2
    assert 'modMedium.RunDataQuality' in src2


def test_xlsm_vba_no_modules_on_plain_xlsx(xlsx_workbook):
    """Plain .xlsx without macros should return an empty dict."""
    mods = xlsx_workbook.iter_vba_modules()
    assert mods == {}


# ---------------------------------------------------------------------------
# Finance_Ledger_100_Unique_Functions.xlsm — macro tests
# ---------------------------------------------------------------------------

FINANCE_XLSM = (
    Path(__file__).resolve().parents[1]
    / "test-data"
    / "Finance_Ledger_100_Unique_Functions.xlsm"
)


@pytest.fixture(scope="module")
def finance_xlsm():
    with XlsxWorkbook(FINANCE_XLSM) as wb:
        yield wb


def test_finance_xlsm_sheet_names(finance_xlsm):
    assert finance_xlsm.sheet_names == [
        "Ledger", "Function_Lab", "Function_Inventory", "Pivot_Tables", "Month_End_Pack"
    ]


def test_finance_xlsm_vba_module_found(finance_xlsm):
    mods = finance_xlsm.iter_vba_modules()
    assert "ThisWorkbook" in mods


def test_finance_xlsm_vba_workbook_event(finance_xlsm):
    src = finance_xlsm.iter_vba_modules()["ThisWorkbook"]
    assert "Private Sub Workbook_Open()" in src
    assert "Application.EnableEvents = False" in src
    assert "modMain.Main" in src

def test_finance_xlsm_formulas_preserved(finance_xlsm):
    """Formulas in the xlsm should match the xlsx/xlsb counterparts."""
    formulas_by_sheet = {s: f for s, f in finance_xlsm.iter_formulas()}
    fl = formulas_by_sheet["Function_Lab"]
    assert len(fl) == 100
    assert fl[(18, 2)] == "=ACOT(1)"   # C19 — UDF prefix stripped
    assert fl[(31, 2)] == '=AVERAGEIFS(Ledger!J186:J241,Ledger!J186:J241,">100")'
