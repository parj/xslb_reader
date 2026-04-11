![Claude](https://img.shields.io/badge/Claude-D97757?style=for-the-badge&logo=claude&logoColor=white) ![ChatGPT](https://img.shields.io/badge/chatGPT-74aa9c?style=for-the-badge&logo=openai&logoColor=white) ![GitHub Actions](https://img.shields.io/badge/github%20actions-%232671E5.svg?style=for-the-badge&logo=githubactions&logoColor=white)
# xlsb_reader

A pure-Python module for reading Excel Binary Workbook (`.xlsb`) files.

The [XLSB format](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-xlsb/acc8aa92-1f02-4167-99f5-84f9f676b95a) published from Microsoft was used to code this up.

> [!WARNING]  
> This has been coded using a mixture of claude (sonnet 4.5) and codex (gpt-5.3 codex).

Supports reading:
- Formulas
- Values
- Pivot tables
- Filters (worksheet AutoFilter and PivotTable value filters)

---

## Installation

```bash
pip install xlsb_reader
```

---

## CLI Usage

The `xlsb_reader` command extracts formulas, values, and pivot table metadata from an `.xlsb` file.

```
xlsb_reader <path> [sheet_name] [--format dict|json|markdown] [--include formulas,values,pivots,filters]
```

### Output all data (default dict format)

```bash
xlsb_reader workbook.xlsb
```

### Filter to a single sheet

```bash
xlsb_reader workbook.xlsb "Sheet1"
```

### JSON output

```bash
xlsb_reader workbook.xlsb --format json
```

### Markdown output

```bash
xlsb_reader workbook.xlsb --format markdown
```

### Only formulas, as JSON

```bash
xlsb_reader workbook.xlsb --include formulas --format json
```

### Only values from a specific sheet

```bash
xlsb_reader workbook.xlsb "Sheet1" --include values --format json
```

### Only pivot table metadata

```bash
xlsb_reader workbook.xlsb --include pivots --format json
```

---

## Python Module Usage

```python
from xlsb_reader import XlsbWorkbook, col_to_letter
```

### List sheet names

```python
with XlsbWorkbook("workbook.xlsb") as wb:
    print(wb.sheet_names)
# ['Sheet1', 'Sheet2', 'Summary']
```

### Read all formulas

`iter_formulas()` yields `(sheet_name: str, formulas: dict[tuple[int, int], str])`.

Formula strings always start with `=`. If a formula cannot be decoded, the value will be
`=<parse_error:...>` rather than raising an exception — filter these out if needed:

```python
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

`iter_values()` yields `(sheet_name: str, values: dict[tuple[int, int], str | int | float | bool | str])`.

Possible value types per cell:

| Type | Example | Notes |
|------|---------|-------|
| `int` | `42` | Integer-valued numbers |
| `float` | `3.14` | Decimal numbers |
| `str` | `"Hello"` | Text cells |
| `bool` | `True` | Boolean cells |
| `str` (error) | `"#DIV/0!"` | Excel error; possible values: `#DIV/0!`, `#N/A`, `#NAME?`, `#NULL!`, `#NUM!`, `#REF!`, `#VALUE!` |

```python
with XlsbWorkbook("workbook.xlsb") as wb:
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
    "name": "PivotTable1",          # str | None
    "cache_id": 1,                  # int | None — links to the pivot cache
    "data_caption": "Values",       # str | None
    "sheet": "Sheet1",              # str — sheet the pivot table lives on
    "pivot_fields": 5,              # int — number of fields (columns) in the cache
    "pivot_items": 42,              # int — total number of items across all fields
    "location": {
        "rfx_geom": {
            "top_left": "A3",       # str — first cell of the pivot table body
            "bottom_right": "D20",  # str — last cell of the pivot table body
        },
        "rw_first_head":  3,        # int — 1-based row of the header row
        "rw_first_data":  5,        # int — 1-based row where data rows start
        "col_first_data": "B",      # str — column letter where data columns start
        "page_rows":      1,        # int — number of page-filter rows
        "page_cols":      0,        # int — number of page-filter columns
    },
    "part": "xl/pivotTables/pivotTable1.bin",                        # str — internal zip path
    "pivot_cache_definition": "xl/pivotCache/pivotCacheDefinition1.bin",  # str | None
    "sx_filters": [                     # list — PivotTable value filters (empty if none)
        {
            "field_index": 2,           # int — 0-based index of the filtered pivot field
            "filter_type": 20,          # int — PivotFilterType (e.g. 20 = valueGreaterThan)
            "criteria": [
                {"operator": ">", "value": 20.0},
            ],
        },
    ],
}
```

```python
with XlsbWorkbook("workbook.xlsb") as wb:
    for pt in wb.iter_pivot_tables():
        print(pt["name"], "on sheet:", pt["sheet"])
        print("  cache_id:", pt.get("cache_id"))
        print("  fields:", pt.get("pivot_fields"))
        loc = pt.get("location") or {}
        geom = loc.get("rfx_geom") or {}
        print(f"  range: {geom.get('top_left')}:{geom.get('bottom_right')}")
```

### Read filters

`iter_filters()` yields `(sheet_name: str, filter_info: dict | None)` for every sheet. `filter_info` is `None` when a sheet has no AutoFilter.

The dict describes the AutoFilter range and the criteria applied to each filtered column:

```python
{
    "range": {
        "top_left":     "A1",   # str — first cell of the AutoFilter range
        "bottom_right": "M241", # str — last cell of the AutoFilter range
    },
    "columns": [
        {
            "column_index": 12,         # int — 0-based column index within the range
            "filters": [],              # list[str] — simple string-match values (BrtFilter)
            "custom_filters": {         # present when comparison criteria are used
                "logic": "and",         # "and" | "or" — how multiple criteria combine
                "criteria": [
                    {
                        "operator": ">",  # "<" | "<=" | "=" | ">=" | ">" | "<>"
                        "value": 1.0,     # float | bool | str | None
                    },
                ],
            },
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
            "field_index": 2,     # int — 0-based index of the filtered pivot field
            "filter_type": 20,    # int — PivotFilterType value (e.g. 20 = valueGreaterThan)
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
with XlsbWorkbook("workbook.xlsb") as wb:
    for sheet_name, finfo in wb.iter_filters():
        if finfo is None:
            continue
        r = finfo["range"]
        print(f"{sheet_name}: AutoFilter on {r['top_left']}:{r['bottom_right']}")
        for col in finfo["columns"]:
            cf = col.get("custom_filters")
            if cf:
                for c in cf["criteria"]:
                    print(f"  col {col['column_index']}: {c['operator']} {c['value']!r}")
            for val in col.get("filters", []):
                print(f"  col {col['column_index']}: = {val!r}")

    for pt in wb.iter_pivot_tables():
        for sf in pt.get("sx_filters", []):
            for c in sf["criteria"]:
                print(
                    f"{pt['name']}: field {sf['field_index']} "
                    f"(type {sf['filter_type']}) {c['operator']} {c['value']!r}"
                )
# Sheet1: AutoFilter on A1:M241
#   col 12: > 1.0
# PivotTable3: field 2 (type 20) > 20.0
```

### Filter to a specific sheet

```python
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

col_to_letter(0)   # 'A'
col_to_letter(25)  # 'Z'
col_to_letter(26)  # 'AA'

row, col = 2, 3    # 0-based → D3
cell = f"{col_to_letter(col)}{row + 1}"
print(cell)        # 'D3'
```
