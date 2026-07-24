![Claude](https://img.shields.io/badge/Claude-D97757?style=for-the-badge&logo=claude&logoColor=white) ![ChatGPT](https://img.shields.io/badge/chatGPT-74aa9c?style=for-the-badge&logo=openai&logoColor=white) ![GitHub Actions](https://img.shields.io/badge/github%20actions-%232671E5.svg?style=for-the-badge&logo=githubactions&logoColor=white)
# xlsb_reader

A pure-Python module for reading Excel workbooks with no third-party dependencies (stdlib only).

Supported formats:

| Format | Extension | Class |
|--------|-----------|-------|
| Excel Binary Workbook | `.xlsb` | `XlsbWorkbook` |
| Excel Open XML Workbook | `.xlsx` | `XlsxWorkbook` |
| Excel Macro-Enabled Workbook | `.xlsm` | `XlsxWorkbook` |

> [!WARNING]  
> This has been coded using a mixture of claude (sonnet 4.5) and codex (gpt-5.3 codex).

Supports reading:
- Formulas
- Values
- Pivot tables
- Filters (worksheet AutoFilter and PivotTable value filters)
- VBA module source code (`.xlsm` / `.xlsb` files with embedded VBA)
- A token-efficient structural spec for LLM code generation via `extract_spec.py` — see [below](#extract_specpy--llm-spec-extraction)

---

## Module structure

```
xlsb_reader/
├── __init__.py          # exports: XlsbWorkbook, XlsxWorkbook, col_to_letter
├── _reader.py           # XlsbWorkbook — .xlsb binary format parser
├── _xlsx_reader.py      # XlsxWorkbook — .xlsx / .xlsm Open XML parser
├── _vba_reader.py       # read_vba_modules(cfb_data) — OLE/OVBA extractor
├── _cli.py              # CLI entry point (xlsb_reader command)
├── _render.py           # to_dict/to_json/to_markdown — shared by the CLI and xlsb_reader.to_*()
└── _spec_extractor.py   # CLI entry point (xlsb-extract-spec command) — see below

rust/                    # optional xlsb_reader_rs extension (pip install xlsb_reader[fast])
└── src/                 # PyO3 port of the modules above, built with maturin
```

Both `XlsbWorkbook` and `XlsxWorkbook` expose the **same public API**:

| Method | Returns | Notes |
|--------|---------|-------|
| `.sheet_names` | `list[str]` | Ordered sheet names |
| `.iter_formulas()` | `Iterator[(sheet, {(row,col): str})]` | Formula strings start with `=` |
| `.iter_values()` | `Iterator[(sheet, {(row,col): value})]` | `int \| float \| str \| bool` |
| `.iter_pivot_tables()` | `Iterator[dict]` | One dict per pivot table |
| `.iter_filters()` | `Iterator[dict]` | One dict per sheet with AutoFilter |
| `.iter_vba_modules()` | `dict[str, str]` | VBA only; empty dict if none |

`row` and `col` are always **0-based** integers.

---

## Installation

```bash
pip install xlsb_reader
```

For very large workbooks, install the optional Rust-accelerated backend:

```bash
pip install "xlsb_reader[fast]"
```

This pulls in a companion compiled extension, `xlsb_reader_rs` (prebuilt
wheels for Linux/macOS/Windows on amd64 and arm64), which reimplements the
same parsing/rendering/spec-extraction logic in Rust for large-file
performance. `xlsb_reader` auto-detects it at import time and uses it
transparently — no code changes needed. Everything below (`XlsbWorkbook`,
`XlsxWorkbook`, the CLI, `extract_spec`) behaves identically either way;
the two backends are verified byte-for-byte identical against each other
in CI (`tests/test_backend_parity.py`).

```python
import xlsb_reader

xlsb_reader.get_backend()  # "rust" if xlsb_reader_rs is installed, else "python"
```

Set `XLSB_READER_BACKEND=python` or `=rust` to force a specific backend
(raises `ImportError` if `rust` is forced but `xlsb_reader_rs` isn't
installed). Without `[fast]`, `xlsb_reader` stays pure-Python with zero
third-party dependencies, exactly as before.

---

## CLI Usage

The `xlsb_reader` command works with `.xlsb`, `.xlsx`, and `.xlsm` files. It auto-detects the format from the file extension.

```
xlsb_reader <path> [sheet_name] [--format dict|json|markdown] [--include formulas,values,pivots,filters,vba]
```

### Output all data (default dict format)

```bash
xlsb_reader workbook.xlsb
xlsb_reader workbook.xlsx
xlsb_reader workbook.xlsm
```

### Filter to a single sheet

```bash
xlsb_reader workbook.xlsb "Sheet1"
xlsb_reader workbook.xlsx "Sheet1"
```

### JSON output

```bash
xlsb_reader workbook.xlsb --format json
xlsb_reader workbook.xlsx --format json
```

### Markdown output

```bash
xlsb_reader workbook.xlsb --format markdown
xlsb_reader workbook.xlsx --format markdown
```

### Only formulas, as JSON

```bash
xlsb_reader workbook.xlsb --include formulas --format json
xlsb_reader workbook.xlsx --include formulas --format json
```

### Only values from a specific sheet

```bash
xlsb_reader workbook.xlsb "Sheet1" --include values --format json
xlsb_reader workbook.xlsx "Sheet1" --include values --format json
```

### Only pivot table metadata

```bash
xlsb_reader workbook.xlsb --include pivots --format json
xlsb_reader workbook.xlsx --include pivots --format json
```

### Extract VBA source (xlsm / xlsb with macros)

```bash
xlsb_reader workbook.xlsm --include vba --format markdown
xlsb_reader workbook.xlsb --include vba --format json
```

---

## `extract_spec.py` — LLM Spec Extraction

`extract_spec.py` is a tool built on top of `xlsb_reader` (uses only `argparse` from the stdlib — no
`click`, no third-party CLI framework, consistent with this project having no dependencies beyond
`xlsb_reader` itself). It turns a workbook — even a 60MB+ one — into a single compact `.spec` text file
describing its **structure**, not its data: column schemas, deduplicated formula patterns, pivot/filter
definitions, VBA source, and a small row sample. The goal is to hand that `.spec` file to an LLM (Claude
Code, ChatGPT, etc.) instead of the raw workbook, so it can write equivalent pandas/polars/openpyxl code
without needing to ingest hundreds of thousands of data rows or burn its context window on near-duplicate
formulas.

### Option A: `pip install` (a command available anywhere)

```bash
pip install xlsb_reader
xlsb-extract-spec workbook.xlsx
xlsb-extract-spec workbook.xlsb --sample-rows 10
xlsb-extract-spec workbook.xlsm --sheets Ledger,Summary --output workbook_ledger.spec
```

`xlsb-extract-spec` is registered as a `[project.scripts]` entry point (see `pyproject.toml`), the same
way the `xlsb_reader` command itself is — so once installed it's on your `PATH`, independent of the
current directory.

### Option B: run the script directly from a repo clone

```bash
git clone <this-repo>
cd xlsb_reader
python extract_spec.py workbook.xlsx
```

`extract_spec.py` at the repo root is a thin wrapper around `xlsb_reader/_spec_extractor.py` (the actual
implementation) — it exists so the tool works without installing anything, straight out of a clone. Both
options run identical code and accept identical flags.

### CLI flags

| Flag | Default | Meaning |
|------|---------|---------|
| `input_file` | — | Path to `.xlsx`, `.xlsb`, or `.xlsm` |
| `--output` | `<input>.spec` | Output path for the spec file |
| `--sample-rows` | `50` | Rows of real data to include per sheet; `0` to omit sample data entirely |
| `--sheets` | `all` | `all` or a comma-separated list of sheet names to include |
| `--validate` | off | Re-reads `--output` (or the default `.spec` path) and checks it against `input_file` instead of generating a new spec — see below |

### What's in a `.spec` file

Sections, in order, separated by `---`:

1. `workbook:` header — file name, format, sheet list, extraction timestamp.
2. Per-sheet `columns[N]{...}` TOON table — inferred type, nullability, up to 5 sample values, and
   notes (e.g. `primary key candidate`, `3 distinct values seen`) for every column.
3. A fenced `json:formulas` block — formulas deduplicated by **pattern**, not by cell. `=MAX(J2,0)`
   repeated over 240 rows collapses to one entry with `row_count: 240`, instead of 240 near-identical
   lines.
4. A fenced `json:dependencies` block — which sheets' formulas reference which other sheets.
5. Fenced `json:pivots` and `json:filters` blocks — pivot table and AutoFilter definitions, normalized to
   one shape regardless of whether the source file was `.xlsb` or `.xlsx`/`.xlsm` (the underlying library
   returns different dict shapes for each).
6. Fenced `vba:ModuleName` blocks — full VBA source, if the workbook has any.
7. Per-sheet `rows[N]{...}` TOON table — the actual `--sample-rows` data sample.
8. `[hints]` — a short, LLM-facing summary: primary table, key transformations with a vectorisation
   suggestion (`pd.merge`/dict map for lookups, otherwise vectorisable), sheet dependencies, and how to
   reproduce each pivot/filter.

Example (trimmed, from a real ledger workbook):

````
sheet: Ledger
dimensions: 241 rows x 13 cols
columns[13]{name,inferred_type,nullable,sample_values,notes}:
  TxnID,string,false,"TX-400000,TX-400001,TX-400002,TX-400003,TX-400004",primary key candidate
  Amount,float,false,"-455.88,6087.62,-1697.91,544.35,-11652.88",
  NetGBP,float,false,"-455.88,4809.2198,-1697.91,544.35,-11652.88",
---
```json:formulas
{"Ledger":{"col_K":{"pattern":"=MAX(<Amount>,0)","example_cell":"K2","example_formula":"=MAX(J2,0)","row_count":240},
           "col_M":{"pattern":"=<Amount>*<FXRate>","example_cell":"M2","example_formula":"=J2*I2","row_count":240}}}
```
---
[hints]
- Primary table: Ledger (241 rows, grain = one row per record)
- Key transformation: Ledger.Debit = =MAX(<Amount>,0) (vectorisable)
- Key transformation: Ledger.NetGBP = =<Amount>*<FXRate> (vectorisable)
````

### Using it with Claude Code (or other coding agents)

1. Generate the spec once, before starting the coding session (`xlsb-extract-spec` if installed via pip,
   or `python extract_spec.py` from a clone):
   ```bash
   xlsb-extract-spec workbook.xlsx --sample-rows 20
   ```
2. Point the agent at the `.spec` file instead of the workbook, e.g.:
   ```
   Read workbook.spec and write a pandas script that reproduces the Ledger sheet's
   Debit/Credit/NetGBP columns and the pivot tables on Pivot_Tables.
   ```
   Since the spec has no bulk data (only the small explicit sample), it's both far smaller than the
   workbook and avoids putting row-level data in the model's context beyond what you asked for.
3. After the agent writes code against the spec, sanity-check the spec itself matches the workbook with:
   ```bash
   xlsb-extract-spec workbook.xlsx --validate
   ```
   This checks that every sheet name and every `<ColumnName>` token used in the `formulas` JSON block
   actually resolves against the live workbook — useful after hand-editing a spec, or after the source
   workbook changes.

### Known limitations

- Column headers are only detected from **row 0**. Sheets with multi-row headers (a title row above the
  real header row) will fall back to `col_A`, `col_B`, ... for columns with no row-0 value.
- Dates stored as Excel serial numbers (the common case) are typed `integer`/`float` per the stated type
  rules, with a `possible Excel serial date` note when the header name contains "date" — they are not
  auto-converted to a `date` type.

---

## Python API

### Imports

```python
# For .xlsb files
from xlsb_reader import XlsbWorkbook, col_to_letter

# For .xlsx / .xlsm files
from xlsb_reader import XlsxWorkbook, col_to_letter

# Or import both
from xlsb_reader import XlsbWorkbook, XlsxWorkbook, col_to_letter
```

### List sheet names

```python
# .xlsb
with XlsbWorkbook("workbook.xlsb") as wb:
    print(wb.sheet_names)
# ['Sheet1', 'Sheet2', 'Summary']

# .xlsx / .xlsm
with XlsxWorkbook("workbook.xlsx") as wb:
    print(wb.sheet_names)
# ['Sheet1', 'Sheet2', 'Summary']
```

### Read all formulas

`iter_formulas()` yields `(sheet_name: str, formulas: dict[tuple[int, int], str])`.

Formula strings always start with `=`. If a formula cannot be decoded, the value will be
`=<parse_error:...>` rather than raising an exception — filter these out if needed:

```python
# Works identically for XlsbWorkbook and XlsxWorkbook
with XlsbWorkbook("workbook.xlsb") as wb:
    for sheet_name, formulas in wb.iter_formulas():
        for (row, col), formula in sorted(formulas.items()):
            if formula.startswith("=<parse_error:"):
                continue  # skip cells that failed to decode
            cell = f"{col_to_letter(col)}{row + 1}"
            print(f"{sheet_name}!{cell}: {formula}")
# Sheet1!A1: =SUM(B1:B10)
# Sheet1!C3: =IF(A3>0,A3*1.2,0)
```

### Read all cell values

`iter_values()` yields `(sheet_name: str, values: dict[tuple[int, int], str | int | float | bool])`.

Possible value types per cell:

| Type | Example | Notes |
|------|---------|-------|
| `int` | `42` | Integer-valued numbers |
| `float` | `3.14` | Decimal numbers |
| `str` | `"Hello"` | Text cells |
| `bool` | `True` | Boolean cells |
| `str` (error) | `"#DIV/0!"` | Excel error; possible values: `#DIV/0!`, `#N/A`, `#NAME?`, `#NULL!`, `#NUM!`, `#REF!`, `#VALUE!` |

```python
# Works identically for XlsbWorkbook and XlsxWorkbook
with XlsxWorkbook("workbook.xlsx") as wb:
    for sheet_name, values in wb.iter_values():
        for (row, col), value in sorted(values.items()):
            cell = f"{col_to_letter(col)}{row + 1}"
            print(f"{sheet_name}!{cell}: {value!r}")
# Sheet1!A1: 42
# Sheet1!B2: 'Hello World'
# Sheet1!C5: True
# Sheet1!D9: '#DIV/0!'
```

### Read pivot table metadata

`iter_pivot_tables()` yields one `dict` per pivot table. Full schema:

```python
{
    "name": "PivotTable1",  # str | None
    "cache_id": 1,  # int | None — links to the pivot cache
    "data_caption": "Values",  # str | None
    "sheet": "Sheet1",  # str — sheet the pivot table lives on
    "pivot_fields": 5,  # int — number of fields (columns) in the cache
    "pivot_items": 42,  # int — total number of items across all fields
    "location": {
        "rfx_geom": {
            "top_left": "A3",  # str — first cell of the pivot table body
            "bottom_right": "D20",  # str — last cell of the pivot table body
        },
        "rw_first_head": 3,  # int — 1-based row of the header row
        "rw_first_data": 5,  # int — 1-based row where data rows start
        "col_first_data": "B",  # str — column letter where data columns start
        "page_rows": 1,  # int — number of page-filter rows
        "page_cols": 0,  # int — number of page-filter columns
    },
    "part": "xl/pivotTables/pivotTable1.bin",  # str — internal zip path
    "pivot_cache_definition": "xl/pivotCache/pivotCacheDefinition1.bin",  # str | None
    "sx_filters": [  # list — PivotTable value filters (empty if none)
        {
            "field_index": 2,  # int — 0-based index of the filtered pivot field
            "filter_type": 20,  # int — PivotFilterType (e.g. 20 = valueGreaterThan)
            "criteria": [
                {"operator": ">", "value": 20.0},
            ],
        },
    ],
}
```

```python
# Works identically for XlsbWorkbook and XlsxWorkbook
with XlsxWorkbook("workbook.xlsx") as wb:
    for pt in wb.iter_pivot_tables():
        print(pt["name"], "on sheet:", pt["sheet"])
        print("  cache_id:", pt.get("cache_id"))
        print("  fields:", pt.get("pivot_fields"))
        loc = pt.get("location") or {}
        geom = loc.get("rfx_geom") or {}
        print(f"  range: {geom.get('top_left')}:{geom.get('bottom_right')}")
```

### Read filters

`iter_filters()` yields one `dict` per sheet that has an AutoFilter (sheets without a filter are skipped).

The dict describes the AutoFilter range and the criteria applied to each filtered column:

```python
# XlsbWorkbook filter dict schema
{
    "range": {
        "top_left": "A1",  # str — first cell of the AutoFilter range
        "bottom_right": "M241",  # str — last cell of the AutoFilter range
    },
    "columns": [
        {
            "column_index": 12,  # int — 0-based column index within the range
            "filters": [],  # list[str] — simple string-match values (BrtFilter)
            "custom_filters": {  # present when comparison criteria are used
                "logic": "and",  # "and" | "or" — how multiple criteria combine
                "criteria": [
                    {
                        "operator": ">",  # "<" | "<=" | "=" | ">=" | ">" | "<>"
                        "value": 1.0,  # float | bool | str | None
                    },
                ],
            },
        },
    ],
}

# XlsxWorkbook filter dict schema
{
    "sheet": "Sheet1",  # str — sheet name
    "ref": "A1:M241",  # str — autoFilter range reference
    "columns": [
        {
            "col_id": 12,  # int — 0-based column index within the range
            "type": "custom",  # "custom" | "discrete" | "top10" | "dynamic"
            # For type="custom":
            "conditions": [
                {
                    "operator": "greaterThan",  # OOXML operator name
                    "val": "1.0",  # str — the comparison value
                },
            ],
            # For type="discrete":
            # "values": ["Apple", "Banana"],   # list[str]
            # For type="top10" or "dynamic":
            # "attrs": {...},                  # raw XML attributes dict
        },
    ],
}
```

PivotTable value filters are exposed via `iter_pivot_tables()` in the `"sx_filters"` key:

```python
{
    # ... other pivot fields ...
    "sx_filters": [
        {
            "field_index": 2,  # int — 0-based index of the filtered pivot field
            "filter_type": 20,  # int — PivotFilterType value (e.g. 20 = valueGreaterThan)
            "criteria": [
                {
                    "operator": ">",
                    "value": 20.0,
                },
            ],
        },
    ],
}
```

```python
# XlsbWorkbook — iter_filters() yields dicts with "range" and "columns" keys
with XlsbWorkbook("workbook.xlsb") as wb:
    for finfo in wb.iter_filters():
        r = finfo["range"]
        print(f"AutoFilter on {r['top_left']}:{r['bottom_right']}")
        for col in finfo["columns"]:
            cf = col.get("custom_filters")
            if cf:
                for c in cf["criteria"]:
                    print(
                        f"  col {col['column_index']}: {c['operator']} {c['value']!r}"
                    )
            for val in col.get("filters", []):
                print(f"  col {col['column_index']}: = {val!r}")

# XlsxWorkbook — iter_filters() yields dicts with "sheet", "ref", and "columns" keys
with XlsxWorkbook("workbook.xlsx") as wb:
    for finfo in wb.iter_filters():
        print(f"{finfo['sheet']}: AutoFilter on {finfo['ref']}")
        for col in finfo["columns"]:
            col_type = col.get("type", "")
            if col_type == "custom":
                for c in col.get("conditions", []):
                    print(f"  col {col['col_id']}: {c['operator']} {c['val']!r}")
            elif col_type == "discrete":
                for val in col.get("values", []):
                    print(f"  col {col['col_id']}: = {val!r}")

# PivotTable filters (same for both workbook types)
with XlsxWorkbook("workbook.xlsx") as wb:
    for pt in wb.iter_pivot_tables():
        for sf in pt.get("sx_filters", []):
            for c in sf["criteria"]:
                print(
                    f"{pt['name']}: field {sf['field_index']} "
                    f"(type {sf['filter_type']}) {c['operator']} {c['value']!r}"
                )
```

### Read VBA module source

`iter_vba_modules()` returns `dict[str, str]` mapping module name to plain-text VBA source.
Returns an empty dict if the workbook contains no VBA project.

Works for `.xlsm` files (macro-enabled Open XML) and `.xlsb` files with embedded VBA.

```python
# .xlsm — macro-enabled workbook
with XlsxWorkbook("workbook.xlsm") as wb:
    modules = wb.iter_vba_modules()
    for module_name, source in modules.items():
        print(f"--- {module_name} ---")
        print(source)

# .xlsb — binary workbook with macros
with XlsbWorkbook("workbook.xlsb") as wb:
    modules = wb.iter_vba_modules()
    for module_name, source in modules.items():
        print(f"--- {module_name} ---")
        print(source)
```

### Filter to a specific sheet

```python
# Works identically for XlsbWorkbook and XlsxWorkbook
with XlsbWorkbook("workbook.xlsb") as wb:
    for sheet_name, formulas in wb.iter_formulas():
        if sheet_name != "Sheet1":
            continue
        for (row, col), formula in sorted(formulas.items()):
            print(f"{col_to_letter(col)}{row + 1}: {formula}")
```

### Convert (row, col) to a cell address

`row` and `col` from `iter_formulas` / `iter_values` are **0-based**.

```python
from xlsb_reader import col_to_letter

col_to_letter(0)  # 'A'
col_to_letter(25)  # 'Z'
col_to_letter(26)  # 'AA'

row, col = 2, 3  # 0-based → D3
cell = f"{col_to_letter(col)}{row + 1}"
print(cell)  # 'D3'
```

### Select workbook class by file extension

```python
import pathlib
from xlsb_reader import XlsbWorkbook, XlsxWorkbook


def open_workbook(path: str):
    suffix = pathlib.Path(path).suffix.lower()
    if suffix in (".xlsx", ".xlsm"):
        return XlsxWorkbook(path)
    return XlsbWorkbook(path)


with open_workbook("data.xlsx") as wb:
    print(wb.sheet_names)
```
 