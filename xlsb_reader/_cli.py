# xlsb_reader/_cli.py
import json
import pprint
from typing import Dict, List, Optional, Tuple

from xlsb_reader._reader import XlsbWorkbook, col_to_letter


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
    wb: XlsbWorkbook,
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
    wb: XlsbWorkbook,
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
    wb: XlsbWorkbook,
    filter_sheet: Optional[str] = None,
) -> List[Dict[str, object]]:
    out: List[Dict[str, object]] = []
    for pt in wb.iter_pivot_tables():
        if filter_sheet and pt.get("sheet") != filter_sheet:
            continue
        out.append(pt)
    return out


def _as_markdown(
    sheets: List[str],
    formulas: Optional[Dict[str, Dict[str, str]]] = None,
    values: Optional[Dict[str, Dict[str, object]]] = None,
    pivots: Optional[List[Dict[str, object]]] = None,
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
    if not emitted:
        lines.append("(no formulas found)")
        return "\n".join(lines)
    return "\n".join(lines).rstrip() + "\n"


def main():
    import argparse

    parser = argparse.ArgumentParser(
        description="Extract formulas, values, and pivot metadata from an .xlsb workbook."
    )
    parser.add_argument("path", help="Path to .xlsb file")
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
        help="Comma-separated sections: formulas,values,pivots (default: formulas,values,pivots)",
    )
    args = parser.parse_args()

    with XlsbWorkbook(args.path) as wb:
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
        if args.output_format == "json":
            print(json.dumps(data, ensure_ascii=False, indent=2, sort_keys=True))
        elif args.output_format == "markdown":
            print(
                _as_markdown(
                    wb.sheet_names,
                    formulas=formulas if "formulas" in includes else None,
                    values=values if "values" in includes else None,
                    pivots=pivots if "pivots" in includes else None,
                ),
                end="",
            )
        else:
            print(pprint.pformat(data, sort_dicts=True))


if __name__ == "__main__":
    main()
