import sys
from pathlib import Path

import pytest

REPO_ROOT = Path(__file__).resolve().parents[1]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from reader import XlsbWorkbook  # noqa: E402

TEST_WORKBOOK = (
    REPO_ROOT
    / "test-data"
    / "Finance_Ledger_100_Unique_Functions.xlsb"
)


@pytest.fixture(scope="module")
def workbook():
    with XlsbWorkbook(TEST_WORKBOOK) as wb:
        yield wb


def test_formula_extraction(workbook):
    formulas_by_sheet = {sheet: formulas for sheet, formulas in workbook.iter_formulas()}
    function_lab = formulas_by_sheet["Function_Lab"]
    ledger = formulas_by_sheet["Ledger"]

    assert function_lab[(17, 2)] == "=ACOSH(2)"  # C18
    assert function_lab[(18, 2)] == "=@ACOT(1)"  # C19
    assert function_lab[(19, 2)] == "=@ACOTH(2)"  # C20
    assert function_lab[(20, 2)] == "=ADDRESS(3,1)"  # C21
    assert function_lab[(31, 2)] == '=AVERAGEIFS(#REF!,#REF!,"GBP")'  # C32

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
    assert pivot_sheet[(48, 1)] == 240.0  # B49


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
    assert p3["location"]["rfx_geom"]["bottom_right"] == "B49"
