"""Cross-backend parity tests.

Verify the optional Rust extension (``xlsb_reader_rs``), when installed,
produces byte-for-byte identical output to the pure-Python backend for
every public API surface: the ``XlsbWorkbook``/``XlsxWorkbook``
``iter_*`` methods, the ``_render`` module's ``to_dict``/``to_json``/
``to_markdown``, and ``_spec_extractor.build_spec``.

These tests compare the pure-Python implementations directly (imported
from ``xlsb_reader._reader``/``_xlsx_reader``/``_render``/
``_spec_extractor``) against ``xlsb_reader_rs`` directly — bypassing
``xlsb_reader``'s backend dispatch entirely, since parity must hold
regardless of which backend a caller happens to have active.

Skipped entirely when ``xlsb_reader_rs`` is not installed (e.g. a base
``pip install xlsb_reader`` without the ``[fast]`` extra).
"""

import re
from pathlib import Path

import pytest

xlsb_reader_rs = pytest.importorskip("xlsb_reader_rs")

from xlsb_reader import _render  # noqa: E402
from xlsb_reader._reader import XlsbWorkbook  # noqa: E402
from xlsb_reader._spec_extractor import build_spec  # noqa: E402
from xlsb_reader._xlsx_reader import XlsxWorkbook  # noqa: E402

TEST_DATA = Path(__file__).resolve().parents[1] / "test-data"
XLSB_PATH = TEST_DATA / "Finance_Ledger_100_Unique_Functions.xlsb"
XLSX_PATH = TEST_DATA / "Finance_Ledger_100_Unique_Functions.xlsx"
XLSM_PATH = TEST_DATA / "Finance_Ledger_100_Unique_Functions.xlsm"

_TIMESTAMP_RE = re.compile(r"^  extracted_at: .*$", re.MULTILINE)


def _normalize_spec(text):
    """Mask the non-deterministic ``extracted_at`` line before comparing."""
    return _TIMESTAMP_RE.sub("  extracted_at: <normalized>", text)


def _collect_xlsb(wb):
    return {
        "sheet_names": list(wb.sheet_names),
        "formulas": sorted(wb.iter_formulas(), key=lambda x: x[0]),
        "values": sorted(wb.iter_values(), key=lambda x: x[0]),
        "filters": sorted(wb.iter_filters(), key=lambda x: x[0]),
        "pivots": list(wb.iter_pivot_tables()),
    }


def _collect_xlsx(wb):
    return {
        "sheet_names": list(wb.sheet_names),
        "formulas": sorted(wb.iter_formulas(), key=lambda x: x[0]),
        "values": sorted(wb.iter_values(), key=lambda x: x[0]),
        "filters": sorted(wb.iter_filters(), key=lambda f: f.get("sheet", "")),
        "pivots": list(wb.iter_pivot_tables()),
        "vba": dict(wb.iter_vba_modules()),
    }


def test_xlsb_workbook_parity():
    with XlsbWorkbook(XLSB_PATH) as py_wb:
        py_data = _collect_xlsb(py_wb)
    with xlsb_reader_rs.XlsbWorkbook(XLSB_PATH) as rs_wb:
        rs_data = _collect_xlsb(rs_wb)
    assert py_data == rs_data


@pytest.mark.parametrize("path", [XLSX_PATH, XLSM_PATH], ids=["xlsx", "xlsm"])
def test_xlsx_workbook_parity(path):
    with XlsxWorkbook(path) as py_wb:
        py_data = _collect_xlsx(py_wb)
    with xlsb_reader_rs.XlsxWorkbook(path) as rs_wb:
        rs_data = _collect_xlsx(rs_wb)
    assert py_data == rs_data


_INCLUDE_SETS = [
    ["formulas", "values", "pivots"],
    ["formulas"],
    ["values"],
    ["pivots"],
    ["filters"],
    ["vba"],
    ["formulas", "values", "pivots", "filters", "vba"],
]
_WORKBOOKS = [
    (XLSB_PATH, XlsbWorkbook),
    (XLSX_PATH, XlsxWorkbook),
    (XLSM_PATH, XlsxWorkbook),
]


@pytest.mark.parametrize("path,cls", _WORKBOOKS, ids=[p.name for p, _ in _WORKBOOKS])
@pytest.mark.parametrize(
    "include", _INCLUDE_SETS, ids=["+".join(s) for s in _INCLUDE_SETS]
)
@pytest.mark.parametrize("sheet", [None, "Function_Lab"])
def test_render_parity(path, cls, include, sheet):
    with cls(path) as wb:
        py_dict = _render.to_dict(wb, include=include, sheet=sheet)
        py_json = _render.to_json(wb, include=include, sheet=sheet)
        py_md = _render.to_markdown(wb, include=include, sheet=sheet)

    rs_dict = xlsb_reader_rs.to_dict(str(path), include=include, sheet=sheet)
    rs_json = xlsb_reader_rs.to_json(str(path), include=include, sheet=sheet)
    rs_md = xlsb_reader_rs.to_markdown(str(path), include=include, sheet=sheet)

    assert py_dict == rs_dict
    assert py_json == rs_json
    assert py_md == rs_md


@pytest.mark.parametrize(
    "path", [XLSB_PATH, XLSX_PATH, XLSM_PATH], ids=lambda p: p.name
)
@pytest.mark.parametrize("sample_rows", [50, 0])
@pytest.mark.parametrize("sheets", ["all", "Function_Lab", "Ledger,Function_Lab"])
def test_spec_parity(path, sample_rows, sheets):
    py_spec = _normalize_spec(build_spec(path, sample_rows, sheets))
    rs_spec = _normalize_spec(
        xlsb_reader_rs.extract_spec(str(path), sample_rows, sheets)
    )
    assert py_spec == rs_spec
