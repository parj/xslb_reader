from pathlib import Path

from xlsb_reader import XlsbWorkbook, XlsxWorkbook
from xlsb_reader._spec_extractor import (
    detect_headers,
    normalize_filters,
    normalize_pivots,
    sheet_dimensions,
)

XLSX_WORKBOOK = (
    Path(__file__).resolve().parents[1]
    / "test-data"
    / "Finance_Ledger_100_Unique_Functions.xlsx"
)
XLSB_WORKBOOK = (
    Path(__file__).resolve().parents[1]
    / "test-data"
    / "Finance_Ledger_100_Unique_Functions.xlsb"
)


def _header_maps(wb):
    header_maps = {}
    for sheet_name, values in wb.iter_values():
        ncols = sheet_dimensions(values)[1]
        header_names, _ = detect_headers(values, ncols)
        header_maps[sheet_name] = header_names
    return header_maps


def test_normalize_pivots_resolves_item_filter_and_value_filter_xlsx():
    with XlsxWorkbook(XLSX_WORKBOOK) as wb:
        pivots = normalize_pivots(wb, lambda _name: True)

    by_name = {p["name"]: p for p in pivots}

    p1_filters = by_name["PivotTable1"]["filters"]
    assert p1_filters == [
        {
            "kind": "item_filter",
            "field_index": 3,
            "field_name": "Category",
            "excluded_values": ["Payroll", "Professional Services"],
        }
    ]

    p3_filters = by_name["PivotTable3"]["filters"]
    assert p3_filters == [
        {
            "kind": "value_filter",
            "field_index": 2,
            "field_name": "AccountCode",
            "operator": "greaterThan",
            "value": "20",
        }
    ]

    assert "filters" not in by_name["PivotTable2"]


def test_normalize_pivots_resolves_value_filter_xlsb():
    with XlsbWorkbook(XLSB_WORKBOOK) as wb:
        pivots = normalize_pivots(wb, lambda _name: True)

    by_name = {p["name"]: p for p in pivots}
    p3_filters = by_name["PivotTable3"]["filters"]
    assert p3_filters == [
        {
            "kind": "value_filter",
            "field_index": 2,
            "field_name": "AccountCode",
            "operator": ">",
            "value": 20.0,
        }
    ]


def test_normalize_filters_resolves_column_name_xlsx():
    with XlsxWorkbook(XLSX_WORKBOOK) as wb:
        header_maps = _header_maps(wb)
        filters = normalize_filters(wb, lambda _name: True, header_maps)

    assert len(filters) == 1
    f = filters[0]
    assert f["sheet"] == "Ledger"
    assert f["columns"] == [
        {
            "index": 12,
            "name": "NetGBP",
            "type": "custom",
            "operator": "greaterThan",
            "value": "1",
        }
    ]


def test_normalize_filters_resolves_column_name_xlsb():
    with XlsbWorkbook(XLSB_WORKBOOK) as wb:
        header_maps = _header_maps(wb)
        filters = normalize_filters(wb, lambda _name: True, header_maps)

    assert len(filters) == 1
    f = filters[0]
    assert f["sheet"] == "Ledger"
    assert f["columns"][0]["name"] == "NetGBP"
    assert f["columns"][0]["index"] == 12
