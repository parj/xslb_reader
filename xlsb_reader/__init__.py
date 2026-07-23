"""
xlsb_reader
===========
Reader for Excel Binary Workbook (.xlsb), Excel Open XML (.xlsx/.xlsm)
workbooks.

By default this package uses a pure-Python implementation (stdlib only).
If the optional ``xlsb_reader_rs`` Rust extension (installable via
``pip install xlsb_reader[fast]``) is present, it is used instead for
better performance — the public API is identical either way.

Backend selection
------------------
Controlled by the ``XLSB_READER_BACKEND`` environment variable:

  - unset / anything other than ``"python"`` (default): try to import
    ``xlsb_reader_rs``; use it if available, otherwise fall back to the
    pure-Python implementation.
  - ``"python"``: always use the pure-Python implementation, even if
    ``xlsb_reader_rs`` is installed.
  - ``"rust"``: require the Rust extension; raise ``ImportError`` with a
    helpful message if it cannot be imported.

Call :func:`get_backend` to find out which backend is actually active.
"""

import os
from typing import Iterable, Optional

__version__ = "1.0.0rc1"

_ENV_VAR = "XLSB_READER_BACKEND"

_rust = None
_backend = "python"

_requested = os.environ.get(_ENV_VAR, "").strip().lower()

if _requested == "python":
    _backend = "python"
else:
    try:
        import xlsb_reader_rs as _rust_module

        _rust = _rust_module
        _backend = "rust"
    except ImportError:
        if _requested == "rust":
            raise ImportError(
                f"{_ENV_VAR}=rust was set, but the 'xlsb_reader_rs' extension "
                "module could not be imported. Install it with "
                "`pip install xlsb_reader[fast]` (or `pip install "
                "xlsb_reader_rs`), or unset XLSB_READER_BACKEND / set it to "
                "'python' to use the pure-Python implementation."
            ) from None
        _backend = "python"

if _backend == "rust":
    XlsbWorkbook = _rust.XlsbWorkbook
    XlsxWorkbook = _rust.XlsxWorkbook
    col_to_letter = _rust.col_to_letter
else:
    from xlsb_reader._reader import XlsbWorkbook, col_to_letter
    from xlsb_reader._xlsx_reader import XlsxWorkbook


def get_backend() -> str:
    """Return the active backend: ``"rust"`` or ``"python"``."""
    return _backend


def _open_python_workbook(path: str):
    """Open ``path`` with the pure-Python workbook class matching its extension."""
    from xlsb_reader._spec_extractor import open_workbook

    return open_workbook(path)


def to_dict(
    path: str,
    include: Iterable[str] = ("formulas", "values", "pivots"),
    sheet: Optional[str] = None,
) -> dict:
    """
    Open the workbook at ``path`` and return a dict of the requested
    sections (``formulas``, ``values``, ``pivots``, ``filters``, ``vba``).

    Uses the Rust backend's one-shot implementation when available (fully
    parsed and rendered in Rust); otherwise opens the workbook with the
    pure-Python reader and assembles the dict via ``xlsb_reader._render``.
    """
    if _backend == "rust" and hasattr(_rust, "to_dict"):
        return _rust.to_dict(path, include=include, sheet=sheet)
    from xlsb_reader._render import to_dict as _py_to_dict

    with _open_python_workbook(path) as wb:
        return _py_to_dict(wb, include=include, sheet=sheet)


def to_json(
    path: str,
    include: Iterable[str] = ("formulas", "values", "pivots"),
    sheet: Optional[str] = None,
) -> str:
    """Same as :func:`to_dict` but returns a JSON string."""
    if _backend == "rust" and hasattr(_rust, "to_json"):
        return _rust.to_json(path, include=include, sheet=sheet)
    from xlsb_reader._render import to_json as _py_to_json

    with _open_python_workbook(path) as wb:
        return _py_to_json(wb, include=include, sheet=sheet)


def to_markdown(
    path: str,
    include: Iterable[str] = ("formulas", "values", "pivots"),
    sheet: Optional[str] = None,
) -> str:
    """Same as :func:`to_dict` but returns a Markdown report."""
    if _backend == "rust" and hasattr(_rust, "to_markdown"):
        return _rust.to_markdown(path, include=include, sheet=sheet)
    from xlsb_reader._render import to_markdown as _py_to_markdown

    with _open_python_workbook(path) as wb:
        return _py_to_markdown(wb, include=include, sheet=sheet)


def extract_spec(path: str, **opts):
    """
    Extract a token-efficient structural spec (schema, formulas, pivots,
    filters, VBA) from the workbook at ``path``, suitable for feeding to an
    LLM writing equivalent pandas/polars/openpyxl code.

    Accepts the same options as ``xlsb_reader._spec_extractor.build_spec``
    (``sample_rows``, ``sheets``) and returns the spec as a string.
    """
    if _backend == "rust" and hasattr(_rust, "extract_spec"):
        return _rust.extract_spec(path, **opts)
    import pathlib

    from xlsb_reader._spec_extractor import build_spec

    sample_rows = opts.pop("sample_rows", 50)
    sheets = opts.pop("sheets", "all")
    if opts:
        raise TypeError(f"Unexpected extract_spec option(s): {sorted(opts)}")
    return build_spec(pathlib.Path(path), sample_rows, sheets)


__all__ = [
    "XlsbWorkbook",
    "XlsxWorkbook",
    "col_to_letter",
    "get_backend",
    "to_dict",
    "to_json",
    "to_markdown",
    "extract_spec",
    "__version__",
]
