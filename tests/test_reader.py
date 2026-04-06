import pytest
from xlsb_reader import XlsbWorkbook
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
    assert pivot_sheet[(22, 4)] == 240.0  # E23
    assert pivot_sheet[(27, 1)] == pytest.approx(59719.2104, rel=1e-9)  # B28
    assert pivot_sheet[(44, 1)] == 169.0  # B45 (PivotTable3 filtered to TxnID > 20)


def test_filter_extraction(workbook):
    # --- Worksheet AutoFilter (Ledger sheet) ---
    filters_by_sheet = {
        sheet: f for sheet, f in workbook.iter_filters() if f is not None
    }

    assert "Ledger" in filters_by_sheet
    ledger_filter = filters_by_sheet["Ledger"]

    assert ledger_filter["range"]["top_left"] == "A1"
    assert ledger_filter["range"]["bottom_right"] == "M241"

    cols = ledger_filter["columns"]
    assert len(cols) == 1
    col = cols[0]
    assert col["column_index"] == 12  # M column (0-based within range)

    cf = col["custom_filters"]
    assert cf["logic"] == "and"
    assert len(cf["criteria"]) == 1
    assert cf["criteria"][0] == {"operator": ">", "value": 1.0}

    # --- PivotTable SX filter (PivotTable3) ---
    pivots = list(workbook.iter_pivot_tables())
    by_name = {p["name"]: p for p in pivots}

    p3 = by_name["PivotTable3"]
    sx = p3["sx_filters"]
    assert len(sx) == 1

    pivot_filter = sx[0]
    assert pivot_filter["field_index"] == 2
    assert pivot_filter["filter_type"] == 20  # valueGreaterThan
    assert len(pivot_filter["criteria"]) == 1
    assert pivot_filter["criteria"][0] == {"operator": ">", "value": 20.0}

    # Sheets without filters return None
    assert filters_by_sheet.get("Function_Inventory") is None


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
    assert p3["location"]["rfx_geom"]["bottom_right"] == "B45"  # filtered: fewer rows
