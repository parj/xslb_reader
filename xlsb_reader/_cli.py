# xlsb_reader/_cli.py
import pprint
from typing import Dict

from xlsb_reader._render import (
    _as_markdown,
    _cellmap,
    _cellmap_any,
    _collect_filters,
    _collect_formulas,
    _collect_pivots,
    _collect_values,
    to_dict,
    to_json,
    to_markdown,
)

__all__ = [
    "_cellmap",
    "_cellmap_any",
    "_collect_filters",
    "_collect_formulas",
    "_collect_pivots",
    "_collect_values",
    "_as_markdown",
    "main",
]


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
        includes = [s.strip() for s in args.include.split(",") if s.strip()]

        if args.output_format == "json":
            print(to_json(wb, include=includes, sheet=args.sheet_name))
        elif args.output_format == "markdown":
            print(to_markdown(wb, include=includes, sheet=args.sheet_name), end="")
        else:
            data: Dict[str, object] = to_dict(
                wb, include=includes, sheet=args.sheet_name
            )
            print(pprint.pformat(data, sort_dicts=True))


if __name__ == "__main__":
    main()
