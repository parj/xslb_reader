"""Extract a token-efficient structural spec from an Excel workbook.

The output is a plan document (schema, formulas, pivots, filters, VBA) that
an LLM can use to write equivalent pandas/polars/openpyxl code -- not the
underlying data itself.
"""

import argparse
import datetime
import json
import pathlib
import re
import sys
from collections import Counter, defaultdict
from typing import Any, Dict, List, Optional, Set, Tuple

from xlsb_reader._reader import XlsbWorkbook, col_to_letter
from xlsb_reader._xlsx_reader import XlsxWorkbook

MIN_SCAN_ROWS = 500
MAX_SAMPLE_VALUES = 5

CELL_REF_RE = re.compile(
    r"(?<![A-Za-z0-9_])"
    r"(?:(?:'(?P<qsheet>[^']+)'|(?P<sheet>[A-Za-z_][A-Za-z0-9_.]*))!)?"
    r"\$?(?P<col>[A-Z]{1,3})\$?(?P<row>\d{1,7})"
    r"(?![A-Za-z0-9_(])"
)
ISO_DATE_RE = re.compile(r"^\d{4}-\d{2}-\d{2}")
FORMULA_TOKEN_RE = re.compile(r"<([^<>]+)>")


def open_workbook(path: str):
    suffix = pathlib.Path(path).suffix.lower()
    if suffix in (".xlsx", ".xlsm"):
        return XlsxWorkbook(path)
    return XlsbWorkbook(path)


def letters_to_col(letters: str) -> int:
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - 64)
    return n - 1


def cell_addr(row: int, col: int) -> str:
    return f"{col_to_letter(col)}{row + 1}"


def csv_field(value: Any) -> str:
    s = "" if value is None else str(value)
    if any(ch in s for ch in (",", '"', "\n")):
        s = '"' + s.replace('"', '""') + '"'
    return s


def toon_table(name: str, columns: List[str], rows: List[List[Any]]) -> str:
    lines = [f"{name}[{len(rows)}]{{{','.join(columns)}}}:"]
    for row in rows:
        lines.append("  " + ",".join(csv_field(v) for v in row))
    return "\n".join(lines)


def fenced_json(label: str, obj: Any) -> str:
    body = json.dumps(obj, ensure_ascii=False, separators=(",", ":"), sort_keys=True)
    return f"```json:{label}\n{body}\n```"


def fenced_vba(name: str, src: str) -> str:
    return f"```vba:{name}\n{src.rstrip()}\n```"


# ---------------------------------------------------------------------------
# Step 2: per-sheet column schema
# ---------------------------------------------------------------------------


def _is_num(v: Any) -> bool:
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def _is_iso_date(v: Any) -> bool:
    if not isinstance(v, str) or not ISO_DATE_RE.match(v):
        return False
    try:
        datetime.date.fromisoformat(v[:10])
        return True
    except ValueError:
        return False


def infer_column_type(values: List[Any]) -> str:
    non_null = [v for v in values if v is not None and v != ""]
    if not non_null:
        return "string"
    if all(_is_num(v) and float(v).is_integer() for v in non_null):
        return "integer"
    if all(_is_num(v) for v in non_null):
        return "float"
    if all(_is_iso_date(v) for v in non_null):
        return "date"
    if all(isinstance(v, bool) for v in non_null):
        return "boolean"
    return "string"


def sheet_dimensions(values: Dict[Tuple[int, int], Any]) -> Tuple[int, int]:
    if not values:
        return 0, 0
    max_row = max(r for r, _ in values)
    max_col = max(c for _, c in values)
    return max_row + 1, max_col + 1


def detect_headers(
    values: Dict[Tuple[int, int], Any], ncols: int
) -> Tuple[Dict[int, str], int]:
    has_header = any(values.get((0, c)) not in (None, "") for c in range(ncols))
    header_names: Dict[int, str] = {}
    for c in range(ncols):
        v = values.get((0, c)) if has_header else None
        header_names[c] = str(v) if v not in (None, "") else f"col_{col_to_letter(c)}"
    return header_names, (1 if has_header else 0)


def build_column_schema(
    values: Dict[Tuple[int, int], Any],
    ncols: int,
    header_names: Dict[int, str],
    data_start_row: int,
    nrows: int,
    scan_rows: int,
) -> List[Dict[str, Any]]:
    window_end = min(nrows, data_start_row + scan_rows)
    columns = []
    for c in range(ncols):
        col_values = [values.get((r, c)) for r in range(data_start_row, window_end)]
        non_null_vals = [v for v in col_values if v is not None and v != ""]
        null_count = len(col_values) - len(non_null_vals)
        col_type = infer_column_type(col_values)

        samples = []
        for v in col_values:
            if len(samples) >= MAX_SAMPLE_VALUES:
                break
            if v is None:
                samples.append("null")
            elif col_type == "integer" and isinstance(v, float):
                samples.append(str(int(v)))
            else:
                samples.append(str(v))

        notes = []
        distinct = set(non_null_vals)
        if (
            c == 0
            and null_count == 0
            and non_null_vals
            and len(distinct) == len(non_null_vals)
        ):
            notes.append("primary key candidate")
        elif (
            col_type == "string"
            and distinct
            and len(distinct) <= 10
            and len(distinct) < len(non_null_vals)
        ):
            notes.append(f"{len(distinct)} distinct values seen")
        if col_type in ("integer", "float") and "date" in header_names[c].lower():
            notes.append("possible Excel serial date")

        columns.append(
            {
                "name": header_names[c],
                "inferred_type": col_type,
                "nullable": null_count > 0,
                "sample_values": ",".join(samples),
                "notes": "; ".join(notes),
            }
        )
    return columns


def render_sheet_schema(
    sheet_name: str, nrows: int, ncols: int, columns: List[Dict[str, Any]]
) -> str:
    lines = [f"sheet: {sheet_name}", f"dimensions: {nrows} rows x {ncols} cols"]
    rows = [
        [
            c["name"],
            c["inferred_type"],
            "true" if c["nullable"] else "false",
            c["sample_values"],
            c["notes"],
        ]
        for c in columns
    ]
    lines.append(
        toon_table(
            "columns",
            ["name", "inferred_type", "nullable", "sample_values", "notes"],
            rows,
        )
    )
    return "\n".join(lines)


def render_sheet_sample(
    sheet_name: str,
    values: Dict[Tuple[int, int], Any],
    header_names: Dict[int, str],
    data_start_row: int,
    nrows: int,
    ncols: int,
    sample_rows: int,
) -> str:
    n = min(sample_rows, max(0, nrows - data_start_row))
    col_order = [header_names[c] for c in range(ncols)]
    rows = []
    for r in range(data_start_row, data_start_row + n):
        rows.append(
            ["" if (v := values.get((r, c))) is None else v for c in range(ncols)]
        )
    lines = [f"sample: {sheet_name}"]
    lines.append(toon_table("rows", col_order, rows))
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# Step 3/4: formula pattern extraction + inter-sheet dependencies
# ---------------------------------------------------------------------------


def normalize_formula(
    formula: str, cur_sheet: str, cur_row: int, header_maps: Dict[str, Dict[int, str]]
) -> Tuple[str, Set[str]]:
    referenced_sheets: Set[str] = set()

    def repl(m: "re.Match[str]") -> str:
        sheet_ref = m.group("qsheet") or m.group("sheet")
        col_idx = letters_to_col(m.group("col"))
        name = header_maps.get(sheet_ref or cur_sheet, {}).get(col_idx)
        token = name if name else col_to_letter(col_idx)
        if sheet_ref:
            referenced_sheets.add(sheet_ref)
            return f"<{sheet_ref}!{token}>"
        offset = (int(m.group("row")) - 1) - cur_row
        if offset == 0:
            return f"<{token}>"
        sign = "+" if offset > 0 else ""
        return f"<{token}{sign}{offset}>"

    pattern = CELL_REF_RE.sub(repl, formula)
    return pattern, referenced_sheets


def extract_formulas(
    wb,
    header_maps: Dict[str, Dict[int, str]],
    sheet_selected,
) -> Tuple[Dict[str, Dict[str, Any]], Dict[str, List[str]]]:
    formulas_by_sheet: Dict[str, Dict[str, Any]] = {}
    dependencies: Dict[str, List[str]] = {}

    for sheet_name, formulas in wb.iter_formulas():
        if not formulas or not sheet_selected(sheet_name):
            continue
        print(f"Processing {sheet_name} (formulas)...", file=sys.stderr)

        col_pattern_counts: Dict[int, Counter] = defaultdict(Counter)
        col_pattern_example: Dict[int, Dict[str, Tuple[str, str]]] = defaultdict(dict)
        col_parse_errors: Dict[int, int] = defaultdict(int)
        deps: Set[str] = set()

        for (row, col), formula in formulas.items():
            if formula.startswith("=<parse_error:"):
                col_parse_errors[col] += 1
                continue
            pattern, refd = normalize_formula(formula, sheet_name, row, header_maps)
            deps.update(refd)
            col_pattern_counts[col][pattern] += 1
            col_pattern_example[col].setdefault(pattern, (cell_addr(row, col), formula))

        deps.discard(sheet_name)
        if deps:
            dependencies[sheet_name] = sorted(deps)

        sheet_json: Dict[str, Any] = {}
        for col, counter in col_pattern_counts.items():
            dominant_pattern, dominant_count = counter.most_common(1)[0]
            entry: Dict[str, Any] = {
                "pattern": dominant_pattern,
                "example_cell": col_pattern_example[col][dominant_pattern][0],
                "example_formula": col_pattern_example[col][dominant_pattern][1],
                "row_count": dominant_count,
            }
            other_count = sum(c for p, c in counter.items() if p != dominant_pattern)
            if other_count:
                entry["other_pattern_count"] = other_count
            if col_parse_errors.get(col):
                entry["parse_error_count"] = col_parse_errors[col]
            sheet_json[f"col_{col_to_letter(col)}"] = entry
        for col, err_count in col_parse_errors.items():
            if col not in col_pattern_counts:
                sheet_json[f"col_{col_to_letter(col)}"] = {
                    "parse_error": True,
                    "count": err_count,
                }

        if sheet_json:
            formulas_by_sheet[sheet_name] = sheet_json

    return formulas_by_sheet, dependencies


# ---------------------------------------------------------------------------
# Step 5/6: pivot tables + autofilters
#
# XlsbWorkbook and XlsxWorkbook return differently-shaped dicts for pivots
# and filters (binary vs. XML parsing paths), so both shapes are normalized
# to one common representation here.
# ---------------------------------------------------------------------------


def normalize_pivots(wb, sheet_selected) -> List[Dict[str, Any]]:
    out = []
    for pt in wb.iter_pivot_tables():
        sheet = pt.get("sheet")
        if sheet is not None and not sheet_selected(sheet):
            continue

        location_str = None
        loc = pt.get("location")
        if isinstance(loc, dict):
            geom = loc.get("rfx_geom") or {}
            tl, br = geom.get("top_left"), geom.get("bottom_right")
            if tl and br:
                location_str = f"{tl}:{br}"

        filters = []
        for sf in pt.get("sx_filters") or []:
            for crit in sf.get("criteria") or []:
                filters.append(
                    {
                        "field_index": sf.get("field_index"),
                        "operator": crit.get("operator"),
                        "value": crit.get("value"),
                    }
                )
        for rf in pt.get("report_filters") or []:
            filters.append({"field_index": rf.get("fld"), "operator": rf.get("type")})

        entry: Dict[str, Any] = {
            "name": pt.get("name"),
            "sheet": sheet,
            "source_cache_id": pt.get("cache_id"),
            "location": location_str,
            "fields": pt.get("pivot_fields"),
        }
        if filters:
            entry["filters"] = filters
        if pt.get("data_fields"):
            entry["data_fields"] = pt["data_fields"]
        if pt.get("row_fields"):
            entry["row_fields"] = pt["row_fields"]
        if pt.get("col_fields"):
            entry["col_fields"] = pt["col_fields"]
        out.append(entry)
    return out


def normalize_filters(wb, sheet_selected) -> List[Dict[str, Any]]:
    out = []
    if not hasattr(wb, "iter_filters"):
        return out
    for item in wb.iter_filters():
        if isinstance(item, tuple):
            sheet, info = item
            if info is None:
                continue
        else:
            info = item
            sheet = info.get("sheet")
        if sheet is not None and not sheet_selected(sheet):
            continue

        rng = None
        if isinstance(info.get("range"), dict):
            rng = f"{info['range'].get('top_left')}:{info['range'].get('bottom_right')}"
        elif info.get("ref"):
            rng = info["ref"]

        columns_out = []
        for col in info.get("columns", []):
            idx = col.get("column_index", col.get("col_id"))
            ctype = col.get("type")
            custom = col.get("custom_filters")
            if custom:
                crits = custom.get("criteria", [])
                logic = custom.get("logic")
                for crit in crits:
                    entry = {
                        "index": idx,
                        "type": "custom",
                        "operator": crit.get("operator"),
                        "value": crit.get("value"),
                    }
                    if logic and len(crits) > 1:
                        entry["logic"] = logic
                    columns_out.append(entry)
            elif col.get("conditions"):
                for crit in col["conditions"]:
                    columns_out.append(
                        {
                            "index": idx,
                            "type": ctype or "custom",
                            "operator": crit.get("operator"),
                            "value": crit.get("val"),
                        }
                    )
            elif col.get("filters"):
                columns_out.append(
                    {"index": idx, "type": "discrete", "values": col["filters"]}
                )
            elif col.get("values"):
                columns_out.append(
                    {"index": idx, "type": ctype or "discrete", "values": col["values"]}
                )
            elif col.get("attrs"):
                columns_out.append(
                    {"index": idx, "type": ctype or "unknown", "attrs": col["attrs"]}
                )
            else:
                columns_out.append({"index": idx, "type": ctype or "unknown"})

        out.append({"sheet": sheet, "range": rng, "columns": columns_out})
    return out


# ---------------------------------------------------------------------------
# Step 9: hints
# ---------------------------------------------------------------------------


def render_hints(
    sheet_dims: Dict[str, Tuple[int, int]],
    formulas_by_sheet: Dict[str, Dict[str, Any]],
    header_maps: Dict[str, Dict[int, str]],
    dependencies: Dict[str, List[str]],
    pivots: List[Dict[str, Any]],
    filters: List[Dict[str, Any]],
    vba_modules: Dict[str, str],
) -> str:
    lines = ["[hints]"]

    if sheet_dims:
        primary_name, (primary_rows, _) = max(
            sheet_dims.items(), key=lambda kv: kv[1][0]
        )
        lines.append(
            f"- Primary table: {primary_name} ({primary_rows} rows, grain = one row per record)"
        )

    for sheet, cols in formulas_by_sheet.items():
        for colkey, info in cols.items():
            if info.get("parse_error"):
                continue
            pattern = info.get("pattern", "")
            col_idx = letters_to_col(colkey.split("_", 1)[1])
            colname = header_maps.get(sheet, {}).get(col_idx, colkey)
            approach = (
                "use pd.merge or dict map"
                if any(f in pattern for f in ("VLOOKUP", "INDEX", "MATCH"))
                else "vectorisable"
            )
            lines.append(
                f"- Key transformation: {sheet}.{colname} = {pattern} ({approach})"
            )

    for sheet, deps in dependencies.items():
        lines.append(f"- {sheet} depends on: {', '.join(deps)}")

    for pt in pivots:
        f_desc = ""
        if pt.get("filters"):
            parts = []
            for f in pt["filters"]:
                val = f.get("value")
                parts.append(
                    f"field {f.get('field_index')} {f.get('operator')}"
                    + (f" {val}" if val is not None else "")
                )
            f_desc = " filters " + "; ".join(parts)
        lines.append(
            f"- Pivot: {pt.get('name')} on {pt.get('sheet')}{f_desc} -> implement via pandas pivot_table()/groupby()"
        )

    for f in filters:
        for col in f.get("columns", []):
            val = col.get("value", col.get("values", ""))
            lines.append(
                f"- AutoFilter on {f['sheet']} col {col.get('index')} {col.get('operator', '')} {val} "
                "-> document as input filter parameter"
            )

    if vba_modules:
        lines.append(
            f"- VBA modules present: {', '.join(vba_modules)} (see vba:... blocks for macro logic)"
        )
    else:
        lines.append("- No VBA found")

    return "\n".join(lines)


# ---------------------------------------------------------------------------
# Orchestration
# ---------------------------------------------------------------------------


def render_header(input_path: pathlib.Path, fmt: str, sheet_names: List[str]) -> str:
    lines = [
        "workbook:",
        f"  file: {input_path.name}",
        f"  format: {fmt}",
        f"  sheets[{len(sheet_names)}]: {','.join(sheet_names)}",
        f"  extracted_at: {datetime.datetime.now().isoformat(timespec='seconds')}",
    ]
    return "\n".join(lines)


def build_spec(input_path: pathlib.Path, sample_rows: int, sheets_arg: str) -> str:
    selected: Optional[Set[str]] = (
        None
        if sheets_arg == "all"
        else {s.strip() for s in sheets_arg.split(",") if s.strip()}
    )

    def sheet_selected(name: str) -> bool:
        return selected is None or name in selected

    fmt = input_path.suffix.lower().lstrip(".")
    scan_rows = max(sample_rows, MIN_SCAN_ROWS)

    with open_workbook(str(input_path)) as wb:
        sheet_names = wb.sheet_names
        header_maps: Dict[str, Dict[int, str]] = {}
        sheet_dims: Dict[str, Tuple[int, int]] = {}
        schema_sections: List[str] = []
        sample_sections: List[str] = []

        for sheet_name, values in wb.iter_values():
            print(f"Processing {sheet_name} (schema)...", file=sys.stderr)
            nrows, ncols = sheet_dimensions(values)
            header_names, data_start_row = detect_headers(values, ncols)
            header_maps[sheet_name] = header_names
            sheet_dims[sheet_name] = (nrows, ncols)

            if not sheet_selected(sheet_name):
                continue

            columns = build_column_schema(
                values, ncols, header_names, data_start_row, nrows, scan_rows
            )
            schema_sections.append(
                render_sheet_schema(sheet_name, nrows, ncols, columns)
            )
            if sample_rows > 0:
                sample_sections.append(
                    render_sheet_sample(
                        sheet_name,
                        values,
                        header_names,
                        data_start_row,
                        nrows,
                        ncols,
                        sample_rows,
                    )
                )

        formulas_by_sheet, dependencies = extract_formulas(
            wb, header_maps, sheet_selected
        )
        pivots = normalize_pivots(wb, sheet_selected)
        filters = normalize_filters(wb, sheet_selected)

        vba_modules: Dict[str, str] = {}
        if hasattr(wb, "iter_vba_modules"):
            print("Processing VBA modules...", file=sys.stderr)
            vba_modules = wb.iter_vba_modules()

    sections = [render_header(input_path, fmt, sheet_names)]
    sections.extend(schema_sections)
    sections.append(fenced_json("formulas", formulas_by_sheet))
    sections.append(fenced_json("dependencies", dependencies))
    sections.append(fenced_json("pivots", pivots))
    sections.append(fenced_json("filters", filters))
    for name, src in vba_modules.items():
        sections.append(fenced_vba(name, src))
    sections.extend(sample_sections)
    sections.append(
        render_hints(
            sheet_dims,
            formulas_by_sheet,
            header_maps,
            dependencies,
            pivots,
            filters,
            vba_modules,
        )
    )

    return "\n---\n".join(sections) + "\n"


# ---------------------------------------------------------------------------
# Bonus: --validate
# ---------------------------------------------------------------------------


def validate_spec(input_path: pathlib.Path, spec_path: pathlib.Path) -> bool:
    if not spec_path.exists():
        print(f"spec file not found: {spec_path}", file=sys.stderr)
        return False
    text = spec_path.read_text(encoding="utf-8")
    problems: List[str] = []

    m = re.search(r"sheets\[\d+\]:\s*(.*)", text)
    spec_sheets = [s.strip() for s in m.group(1).split(",") if s.strip()] if m else []

    with open_workbook(str(input_path)) as wb:
        real_sheets = set(wb.sheet_names)
        header_maps: Dict[str, Dict[int, str]] = {}
        for sheet_name, values in wb.iter_values():
            ncols = sheet_dimensions(values)[1]
            header_names, _ = detect_headers(values, ncols)
            header_maps[sheet_name] = header_names

    for s in spec_sheets:
        if s not in real_sheets:
            problems.append(
                f"sheet '{s}' listed in spec header but not found in workbook"
            )

    fm = re.search(r"```json:formulas\n(.*?)\n```", text, re.DOTALL)
    if fm:
        try:
            formulas_json = json.loads(fm.group(1))
        except json.JSONDecodeError as e:
            problems.append(f"formulas JSON block failed to parse: {e}")
            formulas_json = {}

        for sheet, cols in formulas_json.items():
            if sheet not in real_sheets:
                problems.append(f"formulas reference unknown sheet '{sheet}'")
                continue
            if not isinstance(cols, dict):
                continue
            for colkey, info in cols.items():
                pattern = info.get("pattern", "") if isinstance(info, dict) else ""
                for tok in FORMULA_TOKEN_RE.findall(pattern):
                    tok_sheet, name = tok.split("!", 1) if "!" in tok else (sheet, tok)
                    name = re.sub(r"[+-]\d+$", "", name)
                    known_names = set(header_maps.get(tok_sheet, {}).values())
                    known_letters = {
                        col_to_letter(i) for i in header_maps.get(tok_sheet, {})
                    }
                    if name not in known_names and name not in known_letters:
                        problems.append(
                            f"{sheet}.{colkey}: token <{tok}> does not resolve to a known column in '{tok_sheet}'"
                        )

    if problems:
        for p in problems:
            print(f"INVALID: {p}", file=sys.stderr)
        print(f"{len(problems)} problem(s) found", file=sys.stderr)
        return False

    print("Spec file is valid.", file=sys.stderr)
    return True


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Extract a token-efficient structural spec from an Excel workbook for LLM code generation."
    )
    parser.add_argument("input_file")
    parser.add_argument("--output", default=None)
    parser.add_argument("--sample-rows", type=int, default=50)
    parser.add_argument("--sheets", default="all")
    parser.add_argument(
        "--validate",
        action="store_true",
        help="Validate an existing .spec file against the workbook instead of generating one.",
    )
    args = parser.parse_args()

    input_path = pathlib.Path(args.input_file)
    output_path = (
        pathlib.Path(args.output) if args.output else input_path.with_suffix(".spec")
    )

    if args.validate:
        ok = validate_spec(input_path, output_path)
        sys.exit(0 if ok else 1)

    content = build_spec(input_path, args.sample_rows, args.sheets)
    output_path.write_text(content, encoding="utf-8")
    print(f"Wrote {output_path}", file=sys.stderr)

    try:
        import tiktoken

        enc = tiktoken.get_encoding("cl100k_base")
        print(f"Estimated token count: {len(enc.encode(content))}", file=sys.stderr)
    except ImportError:
        pass


if __name__ == "__main__":
    main()
