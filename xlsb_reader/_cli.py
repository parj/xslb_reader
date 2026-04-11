# xlsb_reader/_cli.py
import json
import pprint
from typing import Dict, List, Optional, Tuple

from xlsb_reader._reader import col_to_letter


def _cellmap(formulas: Dict[Tuple[int, int], str]) -> Dict[str, str]:
    """Convert {(row,col): formula} to {'A1': formula} with stable ordering."""
    out: Dict[str, str] = {}
    for (row, col), formula in sorted(formulas.items()):
        out[f"{col_to_letter(col)}{row + 1}"] = formula
    return out


def _cellmap_any(cells: Dict[Tuple[int, int], object]) -> Dict[str, object]:
    out: Dict[str, object] = {}
    for (row, col), value in sorted(cells.items()):
        out[f"{col_to_letter(col)}{row + 1}"] = value
    return out


def _collect_formulas(
    wb,
    filter_sheet: Optional[str] = None,
) -> Dict[str, Dict[str, str]]:
    out: Dict[str, Dict[str, str]] = {}
    for sheet_name, formulas in wb.iter_formulas():
        if filter_sheet and sheet_name != filter_sheet:
            continue
        if formulas:
            out[sheet_name] = _cellmap(formulas)
    return out


def _collect_values(
    wb,
    filter_sheet: Optional[str] = None,
) -> Dict[str, Dict[str, object]]:
    out: Dict[str, Dict[str, object]] = {}
    for sheet_name, values in wb.iter_values():
        if filter_sheet and sheet_name != filter_sheet:
            continue
        if values:
            out[sheet_name] = _cellmap_any(values)
    return out


def _collect_pivots(
    wb,
    filter_sheet: Optional[str] = None,
) -> List[Dict[str, object]]:
    out: List[Dict[str, object]] = []
    for pt in wb.iter_pivot_tables():
        if filter_sheet and pt.get("sheet") != filter_sheet:
            continue
        out.append(pt)
    return out


def _collect_filters(wb, filter_sheet=None):
    out = []
    for f in wb.iter_filters():
        if filter_sheet and f.get("sheet") != filter_sheet:
            continue
        out.append(f)
    return out


def _as_markdown(
    sheets: List[str],
    formulas: Optional[Dict[str, Dict[str, str]]] = None,
    values: Optional[Dict[str, Dict[str, object]]] = None,
    pivots: Optional[List[Dict[str, object]]] = None,
    filters: Optional[List[Dict[str, object]]] = None,
) -> str:
    lines: List[str] = [f"Sheets: {', '.join(sheets)}", ""]
    emitted = False
    if formulas:
        emitted = True
        lines.append("## Formulas")
        lines.append("")
        for sheet, cell_formulas in formulas.items():
            lines.append(f"### {sheet}")
            for cell, formula in cell_formulas.items():
                lines.append(f"- `{cell}`: `{formula}`")
            lines.append("")
    if values:
        emitted = True
        lines.append("## Values")
        lines.append("")
        for sheet, cell_values in values.items():
            lines.append(f"### {sheet}")
            for cell, value in cell_values.items():
                lines.append(f"- `{cell}`: `{value}`")
            lines.append("")
    if pivots:
        emitted = True
        lines.append("## Pivot Tables")
        lines.append("")
        for pt in pivots:
            lines.append(
                f"- `{pt.get('name') or '<unnamed>'}` "
                f"(sheet: `{pt.get('sheet')}`, cache_id: `{pt.get('cache_id')}`)"
            )
            location = pt.get("location")
            if isinstance(location, dict):
                rfx = location.get("rfx_geom")
                if isinstance(rfx, dict):
                    top_left = rfx.get("top_left")
                    bottom_right = rfx.get("bottom_right")
                    if top_left and bottom_right:
                        lines.append(
                            f"  body: `{top_left}:{bottom_right}`; "
                            f"fields: `{pt.get('pivot_fields')}`; "
                            f"items: `{pt.get('pivot_items')}`"
                        )
    if filters:
        emitted = True
        lines.append("## Filters")
        lines.append("")
        by_sheet: Dict[str, List[Dict[str, object]]] = {}
        for f in filters:
            sheet = f.get("sheet") or "<unknown>"
            by_sheet.setdefault(sheet, []).append(f)
        for sheet, sheet_filters in by_sheet.items():
            lines.append(f"### {sheet}")
            for f in sheet_filters:
                ref = f.get("ref", "")
                columns = f.get("columns", [])
                if columns:
                    for col in columns:
                        col_id = col.get("col_id", "?")
                        col_type = col.get("type", "")
                        conditions = col.get("conditions", [])
                        cond_str = "; ".join(
                            f"{c.get('operator', '')} {c.get('val', '')}"
                            for c in conditions
                        ) if conditions else str(col.get("attrs", ""))
                        lines.append(f"- `{ref}` — column {col_id}: {col_type} {cond_str}")
                else:
                    lines.append(f"- `{ref}`")
            lines.append("")
    if not emitted:
        lines.append("(no formulas found)")
        return "\n".join(lines)
    return "\n".join(lines).rstrip() + "\n"


def main():
    import argparse

    parser = argparse.ArgumentParser(
        description="Extract formulas, values, pivot metadata, and filters from an .xlsb or .xlsx workbook."
    )
    parser.add_argument("path", help="Path to .xlsb or .xlsx file")
    parser.add_argument(
        "sheet_name", nargs="?", default=None, help="Optional sheet name filter"
    )
    parser.add_argument(
        "--format",
        dest="output_format",
        choices=("dict", "json", "markdown"),
        default="dict",
        help="Output format (default: dict)",
    )
    parser.add_argument(
        "--include",
        default="formulas,values,pivots",
        help="Comma-separated sections: formulas,values,pivots,filters,vba (default: formulas,values,pivots)",
    )
    args = parser.parse_args()

    if args.path.lower().endswith((".xlsx", ".xlsm")):
        from xlsb_reader._xlsx_reader import XlsxWorkbook as WorkbookClass
    else:
        from xlsb_reader._reader import XlsbWorkbook as WorkbookClass

    with WorkbookClass(args.path) as wb:
        includes = {s.strip().lower() for s in args.include.split(",") if s.strip()}
        formulas = (
            _collect_formulas(wb, filter_sheet=args.sheet_name)
            if "formulas" in includes
            else {}
        )
        values = (
            _collect_values(wb, filter_sheet=args.sheet_name)
            if "values" in includes
            else {}
        )
        pivots = (
            _collect_pivots(wb, filter_sheet=args.sheet_name)
            if "pivots" in includes
            else []
        )
        data: Dict[str, object] = {}
        if "formulas" in includes:
            data["formulas"] = formulas
        if "values" in includes:
            data["values"] = values
        if "pivots" in includes:
            data["pivot_tables"] = pivots
        filters: List[Dict[str, object]] = []
        if "filters" in includes and hasattr(wb, "iter_filters"):
            filters = _collect_filters(wb, filter_sheet=args.sheet_name)
            data["filters"] = filters
        vba: Dict[str, str] = {}
        if "vba" in includes and hasattr(wb, "iter_vba_modules"):
            vba = wb.iter_vba_modules()
            data["vba_modules"] = vba
        if args.output_format == "json":
            print(json.dumps(data, ensure_ascii=False, indent=2, sort_keys=True))
        elif args.output_format == "markdown":
            md = _as_markdown(
                wb.sheet_names,
                formulas=formulas if "formulas" in includes else None,
                values=values if "values" in includes else None,
                pivots=pivots if "pivots" in includes else None,
                filters=filters if "filters" in includes else None,
            )
            if vba:
                vba_lines = ["## VBA Modules", ""]
                for mod_name, src in vba.items():
                    vba_lines.append(f"### {mod_name}")
                    vba_lines.append("```vba")
                    vba_lines.append(src.rstrip())
                    vba_lines.append("```")
                    vba_lines.append("")
                md = md.rstrip() + "\n\n" + "\n".join(vba_lines)
            print(md, end="")
        else:
            print(pprint.pformat(data, sort_dicts=True))


if __name__ == "__main__":
    main()
