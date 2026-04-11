"""
_xlsx_reader.py
===============
Pure-Python reader for Excel Open XML (.xlsx) files.
Mirrors the public API of XlsbWorkbook in _reader.py.
No third-party dependencies — stdlib only.

Supports:
  - Sheet names
  - Cell formulas (shared, array, normal, UDF)
  - Cell values (strings, numbers, booleans, errors, inline strings)
  - Pivot table metadata
  - Auto-filter metadata (XLSX-only)
"""

import os
import re
import xml.etree.ElementTree as ET
import zipfile
from typing import Dict, Iterator, List, Optional, Tuple


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def col_to_letter(col: int) -> str:
    """Convert 0-based column index to Excel letter(s), e.g. 0->A, 26->AA."""
    result = ""
    col += 1
    while col:
        col, rem = divmod(col - 1, 26)
        result = chr(65 + rem) + result
    return result


def _col_from_str(s: str) -> int:
    """Convert Excel column letters to 0-based index, e.g. 'A'->0, 'AA'->26."""
    result = 0
    for ch in s:
        result = result * 26 + (ord(ch) - ord("A") + 1)
    return result - 1


def _cell_ref_to_row_col(ref: str) -> Tuple[int, int]:
    """
    Parse a cell reference like 'K3' into (row, col), both 0-based.
    e.g. 'A1' -> (0, 0), 'K3' -> (2, 10)
    """
    # Split into letter part and digit part
    i = 0
    while i < len(ref) and ref[i].isalpha():
        i += 1
    col_str = ref[:i]
    row_str = ref[i:]
    col = _col_from_str(col_str)
    row = int(row_str) - 1
    return row, col


def _ns(tag: str) -> str:
    """Strip namespace from a tag, e.g. '{ns}foo' -> 'foo'."""
    return tag.split("}")[-1] if "}" in tag else tag


# SpreadsheetML namespace
_SML_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

# Namespace prefix map for iterating
_NS_MAP = {"s": _SML_NS, "r": _REL_NS}


def _q(local: str) -> str:
    """Build qualified name: {sml_ns}local."""
    return f"{{{_SML_NS}}}{local}"


def _rq(local: str) -> str:
    """Build qualified name: {rel_ns}local."""
    return f"{{{_REL_NS}}}{local}"


# ---------------------------------------------------------------------------
# Shared formula expansion helpers
# ---------------------------------------------------------------------------

_CELL_RE = re.compile(r"(\$?)([A-Z]+)(\$?)(\d+)")


def _shift_ref(match: re.Match, dr: int, dc: int) -> str:
    col_abs, col_str, row_abs, row_str = match.groups()
    col = _col_from_str(col_str)
    row = int(row_str) - 1  # 0-based
    if not col_abs:
        col += dc
    if not row_abs:
        row += dr
    return f"{col_abs}{col_to_letter(col)}{row_abs}{row + 1}"


def _expand_shared_formula(formula: str, dr: int, dc: int) -> str:
    if dr == 0 and dc == 0:
        return formula
    return _CELL_RE.sub(lambda m: _shift_ref(m, dr, dc), formula)


# ---------------------------------------------------------------------------
# Relationships helpers
# ---------------------------------------------------------------------------


def _parse_rels(data: bytes) -> Dict[str, str]:
    """Return {Id: Target} from a .rels XML file."""
    try:
        root = ET.fromstring(data)
    except ET.ParseError:
        return {}
    result: Dict[str, str] = {}
    for child in root:
        tag = _ns(child.tag)
        if tag == "Relationship":
            rid = child.get("Id", "")
            target = child.get("Target", "")
            if rid and target:
                result[rid] = target
    return result


def _resolve_rel_target(part_path: str, target: str) -> str:
    """Resolve a relationship target relative to part_path's directory."""
    base = part_path.rsplit("/", 1)[0] if "/" in part_path else ""
    if target.startswith("/"):
        return target.lstrip("/")
    parts = (base + "/" + target).split("/")
    resolved: List[str] = []
    for p in parts:
        if p == "..":
            if resolved:
                resolved.pop()
        elif p and p != ".":
            resolved.append(p)
    return "/".join(resolved)


# ---------------------------------------------------------------------------
# Shared string table parsing
# ---------------------------------------------------------------------------


def _parse_shared_strings(data: bytes) -> List[str]:
    """
    Parse xl/sharedStrings.xml and return ordered list of strings.
    Handles both simple <t> and rich text <r><t> runs.
    """
    try:
        root = ET.fromstring(data)
    except ET.ParseError:
        return []

    strings: List[str] = []
    for si in root:
        if _ns(si.tag) != "si":
            continue
        # Check for rich text runs first
        runs = [child for child in si if _ns(child.tag) == "r"]
        if runs:
            parts: List[str] = []
            for r in runs:
                for t_el in r:
                    if _ns(t_el.tag) == "t":
                        parts.append(t_el.text or "")
            strings.append("".join(parts))
        else:
            # Simple <t> element
            t_el = si.find(_q("t"))
            strings.append(t_el.text if t_el is not None and t_el.text else "")
    return strings


# ---------------------------------------------------------------------------
# Workbook XML parsing
# ---------------------------------------------------------------------------


def _parse_workbook(data: bytes) -> Tuple[List[Tuple[str, str]], Dict[str, str]]:
    """
    Parse xl/workbook.xml.
    Returns:
      - sheets: [(sheet_name, rId), ...]
      - defined_names: {name: formula_text}
    """
    try:
        root = ET.fromstring(data)
    except ET.ParseError:
        return [], {}

    sheets: List[Tuple[str, str]] = []
    defined_names: Dict[str, str] = {}

    for child in root:
        tag = _ns(child.tag)
        if tag == "sheets":
            for sheet_el in child:
                if _ns(sheet_el.tag) == "sheet":
                    name = sheet_el.get("name", "")
                    # rId is in the r: namespace
                    rid = sheet_el.get(
                        "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id",
                        "",
                    ) or sheet_el.get("r:id", "")
                    if name and rid:
                        sheets.append((name, rid))
        elif tag == "definedNames":
            for dn_el in child:
                if _ns(dn_el.tag) == "definedName":
                    dn_name = dn_el.get("name", "")
                    dn_text = dn_el.text or ""
                    if dn_name:
                        defined_names[dn_name] = dn_text

    return sheets, defined_names


# ---------------------------------------------------------------------------
# Worksheet XML parsing — formulas
# ---------------------------------------------------------------------------


def _parse_worksheet_formulas(
    data: bytes,
    sst: List[str],
) -> Dict[Tuple[int, int], str]:
    """
    Parse worksheet XML and return {(row, col): formula_string}.
    Formula strings start with '=' (or '={' for array formulas).
    Row/col are 0-based.
    """
    try:
        root = ET.fromstring(data)
    except ET.ParseError:
        return {}

    formulas: Dict[Tuple[int, int], str] = {}
    # si index -> (formula_text, anchor_row, anchor_col)
    shared_formulas: Dict[str, Tuple[str, int, int]] = {}

    sheet_data = root.find(_q("sheetData"))
    if sheet_data is None:
        return {}

    for row_el in sheet_data:
        if _ns(row_el.tag) != "row":
            continue
        for c_el in row_el:
            if _ns(c_el.tag) != "c":
                continue
            ref = c_el.get("r", "")
            if not ref:
                continue
            row, col = _cell_ref_to_row_col(ref)

            f_el = c_el.find(_q("f"))
            if f_el is None:
                continue

            f_text = f_el.text or ""
            f_type = f_el.get("t", "normal")
            f_si = f_el.get("si", "")
            f_ref = f_el.get("ref", "")  # e.g. "K2:K65" for shared anchor
            f_ca = f_el.get("ca", "0")

            if f_type == "array":
                # Array formula: wrap in braces
                formula_text = f_text
                # Strip _xludf. prefix if ca=1
                if f_ca == "1":
                    formula_text = re.sub(r"_xludf\.", "", formula_text)
                formulas[(row, col)] = "{=" + formula_text + "}"

            elif f_type == "shared":
                if f_ref:
                    # Anchor cell: has the formula text and 'ref' attr
                    formula_text = f_text
                    # Strip _xludf. prefix if ca=1
                    if f_ca == "1":
                        formula_text = re.sub(r"_xludf\.", "", formula_text)
                    if f_si:
                        shared_formulas[f_si] = (formula_text, row, col)
                    formulas[(row, col)] = "=" + formula_text
                else:
                    # Follower cell: look up shared formula and expand
                    if f_si and f_si in shared_formulas:
                        anchor_formula, anchor_row, anchor_col = shared_formulas[f_si]
                        dr = row - anchor_row
                        dc = col - anchor_col
                        expanded = _expand_shared_formula(anchor_formula, dr, dc)
                        formulas[(row, col)] = "=" + expanded
                    # else: shared formula not seen yet — will be missing

            else:
                # Normal formula (no t attr or t="normal")
                formula_text = f_text
                # Strip _xludf. prefix if ca=1 (UDF marker)
                if f_ca == "1":
                    formula_text = re.sub(r"_xludf\.", "", formula_text)
                formulas[(row, col)] = "=" + formula_text

    return formulas


# ---------------------------------------------------------------------------
# Worksheet XML parsing — values
# ---------------------------------------------------------------------------


def _parse_worksheet_values(
    data: bytes,
    sst: List[str],
) -> Dict[Tuple[int, int], object]:
    """
    Parse worksheet XML and return {(row, col): value}.
    Row/col are 0-based.
    """
    try:
        root = ET.fromstring(data)
    except ET.ParseError:
        return {}

    values: Dict[Tuple[int, int], object] = {}

    sheet_data = root.find(_q("sheetData"))
    if sheet_data is None:
        return {}

    for row_el in sheet_data:
        if _ns(row_el.tag) != "row":
            continue
        for c_el in row_el:
            if _ns(c_el.tag) != "c":
                continue
            ref = c_el.get("r", "")
            if not ref:
                continue
            row, col = _cell_ref_to_row_col(ref)
            cell_type = c_el.get("t", "")

            if cell_type == "inlineStr":
                # Inline string: read from <is><t>
                is_el = c_el.find(_q("is"))
                if is_el is not None:
                    t_el = is_el.find(_q("t"))
                    values[(row, col)] = t_el.text if t_el is not None else ""
                else:
                    values[(row, col)] = ""
                continue

            v_el = c_el.find(_q("v"))
            v_text = v_el.text if v_el is not None else None

            if v_text is None:
                # No value element — could be blank
                continue

            if cell_type == "s":
                # Shared string index
                try:
                    idx = int(v_text)
                    values[(row, col)] = sst[idx] if 0 <= idx < len(sst) else ""
                except (ValueError, IndexError):
                    values[(row, col)] = ""

            elif cell_type == "b":
                # Boolean
                values[(row, col)] = bool(int(v_text))

            elif cell_type == "e":
                # Error string
                values[(row, col)] = v_text

            elif cell_type in ("str", ""):
                # formula string result or number (no t attr means number)
                if cell_type == "str":
                    values[(row, col)] = v_text
                else:
                    # Number: parse as float; convert to int if whole
                    try:
                        f = float(v_text)
                        if f == int(f) and abs(f) < 1e15:
                            values[(row, col)] = int(f)
                        else:
                            values[(row, col)] = f
                    except ValueError:
                        values[(row, col)] = v_text

            else:
                # t="n" explicit number, or other unrecognised type -> try numeric
                try:
                    f = float(v_text)
                    if f == int(f) and abs(f) < 1e15:
                        values[(row, col)] = int(f)
                    else:
                        values[(row, col)] = f
                except ValueError:
                    values[(row, col)] = v_text

    return values


# ---------------------------------------------------------------------------
# Pivot table XML parsing
# ---------------------------------------------------------------------------


def _parse_pivot_table_xml(data: bytes) -> Dict[str, object]:
    """
    Parse xl/pivotTables/pivotTableN.xml and return metadata dict.
    """
    try:
        root = ET.fromstring(data)
    except ET.ParseError:
        return {}

    meta: Dict[str, object] = {}

    # Root attributes
    meta["name"] = root.get("name", "")
    try:
        meta["cache_id"] = int(root.get("cacheId", "0"))
    except ValueError:
        meta["cache_id"] = 0
    meta["data_caption"] = root.get("dataCaption", "")

    # Location
    loc_el = root.find(_q("location"))
    if loc_el is not None:
        ref = loc_el.get("ref", "")
        top_left = ref.split(":")[0] if ":" in ref else ref
        bottom_right = ref.split(":")[1] if ":" in ref else ref
        try:
            rw_first_head = int(loc_el.get("firstHeaderRow", "0"))
        except ValueError:
            rw_first_head = 0
        try:
            rw_first_data = int(loc_el.get("firstDataRow", "0"))
        except ValueError:
            rw_first_data = 0
        try:
            col_first_data_int = int(loc_el.get("firstDataCol", "0"))
        except ValueError:
            col_first_data_int = 0
        # col_first_data: use the actual column letter from top_left + offset
        # Per spec: col_first_data attr is the 0-based column offset within the pivot range
        # The actual column = top_left_col + firstDataCol
        if top_left:
            tl_row, tl_col = _cell_ref_to_row_col(top_left)
            actual_col = tl_col + col_first_data_int
            col_first_data_letter = col_to_letter(actual_col)
        else:
            col_first_data_letter = col_to_letter(col_first_data_int)

        meta["location"] = {
            "rfx_geom": {
                "top_left": top_left,
                "bottom_right": bottom_right,
            },
            "rw_first_head": rw_first_head,
            "rw_first_data": rw_first_data,
            "col_first_data": col_first_data_letter,
            "page_rows": 0,
            "page_cols": 0,
        }
    else:
        meta["location"] = None

    # Pivot fields
    pf_el = root.find(_q("pivotFields"))
    pivot_fields_count = 0
    pivot_items_count = 0
    if pf_el is not None:
        for pf in pf_el:
            if _ns(pf.tag) == "pivotField":
                pivot_fields_count += 1
                items_el = pf.find(_q("items"))
                if items_el is not None:
                    for item in items_el:
                        if _ns(item.tag) == "item":
                            pivot_items_count += 1
    meta["pivot_fields"] = pivot_fields_count
    meta["pivot_items"] = pivot_items_count

    # Row fields
    row_fields: List[int] = []
    rf_el = root.find(_q("rowFields"))
    if rf_el is not None:
        for field_el in rf_el:
            if _ns(field_el.tag) == "field":
                try:
                    x = int(field_el.get("x", "0"))
                    row_fields.append(x)
                except ValueError:
                    pass
    meta["row_fields"] = row_fields

    # Col fields (skip x="-2" which is the data placeholder)
    col_fields: List[int] = []
    cf_el = root.find(_q("colFields"))
    if cf_el is not None:
        for field_el in cf_el:
            if _ns(field_el.tag) == "field":
                try:
                    x = int(field_el.get("x", "0"))
                    if x != -2:
                        col_fields.append(x)
                except ValueError:
                    pass
    meta["col_fields"] = col_fields

    # Data fields
    data_fields: List[Dict[str, object]] = []
    df_el = root.find(_q("dataFields"))
    if df_el is not None:
        for df in df_el:
            if _ns(df.tag) == "dataField":
                df_name = df.get("name", "")
                try:
                    fld = int(df.get("fld", "0"))
                except ValueError:
                    fld = 0
                subtotal = df.get("subtotal", "sum")
                data_fields.append({"name": df_name, "fld": fld, "subtotal": subtotal})
    meta["data_fields"] = data_fields

    # Report filters (from <filters> element inside pivot table definition)
    report_filters: List[Dict[str, object]] = []
    filters_el = root.find(_q("filters"))
    if filters_el is not None:
        for filt_el in filters_el:
            if _ns(filt_el.tag) == "filter":
                try:
                    fld_idx = int(filt_el.get("fld", "0"))
                except ValueError:
                    fld_idx = 0
                filt_type = filt_el.get("type", "")
                report_filters.append({"fld": fld_idx, "type": filt_type})
    meta["report_filters"] = report_filters

    return meta


# ---------------------------------------------------------------------------
# Auto-filter XML parsing
# ---------------------------------------------------------------------------


def _parse_auto_filter(af_el: ET.Element, sheet_name: str) -> Dict[str, object]:
    """
    Parse an <autoFilter> element and return a filter metadata dict.
    """
    ref = af_el.get("ref", "")
    columns: List[Dict[str, object]] = []

    for fc_el in af_el:
        if _ns(fc_el.tag) != "filterColumn":
            continue
        try:
            col_id = int(fc_el.get("colId", "0"))
        except ValueError:
            col_id = 0

        col_info: Dict[str, object] = {"col_id": col_id}

        # Determine filter type from child elements
        for filter_child in fc_el:
            child_tag = _ns(filter_child.tag)

            if child_tag == "customFilters":
                col_info["type"] = "custom"
                conditions: List[Dict[str, str]] = []
                for cf in filter_child:
                    if _ns(cf.tag) == "customFilter":
                        conditions.append(
                            {
                                "operator": cf.get("operator", "equal"),
                                "val": cf.get("val", ""),
                            }
                        )
                col_info["conditions"] = conditions

            elif child_tag == "filters":
                col_info["type"] = "discrete"
                vals: List[str] = []
                for f in filter_child:
                    if _ns(f.tag) == "filter":
                        vals.append(f.get("val", ""))
                col_info["values"] = vals

            elif child_tag == "top10":
                col_info["type"] = "top10"
                col_info["attrs"] = dict(filter_child.attrib)

            elif child_tag == "dynamicFilter":
                col_info["type"] = "dynamic"
                col_info["attrs"] = dict(filter_child.attrib)

            break  # Only process the first filter type child

        columns.append(col_info)

    return {
        "sheet": sheet_name,
        "ref": ref,
        "columns": columns,
    }


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------


class XlsxWorkbook:
    """
    Open an .xlsx workbook and iterate worksheet formulas/values,
    PivotTable metadata, and auto-filter metadata.

    Parameters
    ----------
    path : str or path-like
        Path to the .xlsx file.

    Example
    -------
    >>> with XlsxWorkbook("data.xlsx") as wb:
    ...     for sheet, formulas in wb.iter_formulas():
    ...         for (row, col), f in sorted(formulas.items()):
    ...             print(f"{sheet}!{col_to_letter(col)}{row+1}: {f}")
    """

    def __init__(self, path: "os.PathLike"):
        self._zf = zipfile.ZipFile(str(path), "r")
        self._sst: List[str] = []
        self._sheets: List[Tuple[str, str]] = []  # [(name, rId), ...]
        self._paths: Dict[str, str] = {}  # rId -> "xl/worksheets/sheetN.xml"
        self._defined_names: Dict[str, str] = {}
        self._init_workbook()
        self._init_sst()

    def _read_part(self, path: str) -> bytes:
        """Read a part from the ZIP, with case-insensitive fallback."""
        path = path.replace("\\", "/").lstrip("/")
        try:
            return self._zf.read(path)
        except KeyError:
            lo = path.lower()
            for n in self._zf.namelist():
                if n.lower() == lo:
                    return self._zf.read(n)
            raise FileNotFoundError(f"Not found in XLSX zip: {path}")

    def _init_workbook(self):
        wb_data = self._read_part("xl/workbook.xml")
        self._sheets, self._defined_names = _parse_workbook(wb_data)

        try:
            rels = _parse_rels(self._read_part("xl/_rels/workbook.xml.rels"))
        except FileNotFoundError:
            rels = {}

        for _name, rid in self._sheets:
            target = rels.get(rid, "")
            if target:
                if target.startswith("/"):
                    full = target.lstrip("/")
                elif target.startswith("xl/"):
                    full = target
                else:
                    full = "xl/" + target
                self._paths[rid] = full.replace("//", "/")

    def _init_sst(self):
        try:
            self._sst = _parse_shared_strings(self._read_part("xl/sharedStrings.xml"))
        except FileNotFoundError:
            self._sst = []

    @property
    def sheet_names(self) -> List[str]:
        """Ordered list of worksheet names."""
        return [n for n, _ in self._sheets]

    def iter_formulas(
        self,
    ) -> Iterator[Tuple[str, Dict[Tuple[int, int], str]]]:
        """
        Yield (sheet_name, formulas) for every sheet.
        formulas maps (row, col) -> formula string (starts with '=').
        Row and col are 0-based.
        """
        for sheet_name, rid in self._sheets:
            zpath = self._paths.get(rid)
            if not zpath:
                yield sheet_name, {}
                continue
            try:
                ws_data = self._read_part(zpath)
            except FileNotFoundError:
                yield sheet_name, {}
                continue
            yield sheet_name, _parse_worksheet_formulas(ws_data, self._sst)

    def iter_values(
        self,
    ) -> Iterator[Tuple[str, Dict[Tuple[int, int], object]]]:
        """
        Yield (sheet_name, values) for every sheet.
        values maps (row, col) -> cached/constant value.
        Row and col are 0-based.
        """
        for sheet_name, rid in self._sheets:
            zpath = self._paths.get(rid)
            if not zpath:
                yield sheet_name, {}
                continue
            try:
                ws_data = self._read_part(zpath)
            except FileNotFoundError:
                yield sheet_name, {}
                continue
            yield sheet_name, _parse_worksheet_values(ws_data, self._sst)

    def iter_pivot_tables(self) -> Iterator[Dict[str, object]]:
        """
        Yield parsed PivotTable metadata dicts, one per pivot table.
        Each dict includes 'sheet', 'part', and 'pivot_cache_definition' keys.
        """
        # Map sheet part path -> sheet name
        sheet_by_path: Dict[str, str] = {}
        for sheet_name, rid in self._sheets:
            zpath = self._paths.get(rid)
            if zpath:
                sheet_by_path[zpath] = sheet_name

        seen_parts: set = set()

        for sheet_path, sheet_name in sheet_by_path.items():
            # Build path to worksheet rels file
            dirname = sheet_path.rsplit("/", 1)[0] if "/" in sheet_path else ""
            basename = sheet_path.rsplit("/", 1)[-1]
            rel_path = f"{dirname}/_rels/{basename}.rels"

            try:
                rels_data = self._read_part(rel_path)
            except FileNotFoundError:
                continue

            rels = _parse_rels(rels_data)
            for _, target in rels.items():
                resolved = _resolve_rel_target(sheet_path, target)
                if "/pivotTables/" not in resolved:
                    continue
                if resolved in seen_parts:
                    continue
                seen_parts.add(resolved)

                try:
                    pt_data = self._read_part(resolved)
                except FileNotFoundError:
                    continue

                meta = _parse_pivot_table_xml(pt_data)
                meta["sheet"] = sheet_name
                meta["part"] = resolved

                # Resolve pivot cache definition via pivot table rels
                pt_dirname = resolved.rsplit("/", 1)[0] if "/" in resolved else ""
                pt_basename = resolved.rsplit("/", 1)[-1]
                pt_rel_path = f"{pt_dirname}/_rels/{pt_basename}.rels"

                try:
                    pt_rels_data = self._read_part(pt_rel_path)
                    pt_rels = _parse_rels(pt_rels_data)
                except FileNotFoundError:
                    pt_rels = {}

                cache_def: Optional[str] = None
                for _, t in pt_rels.items():
                    rr = _resolve_rel_target(resolved, t)
                    if "/pivotCache/pivotCacheDefinition" in rr:
                        cache_def = rr
                        break
                if cache_def:
                    meta["pivot_cache_definition"] = cache_def

                yield meta

    def iter_filters(self) -> Iterator[Dict[str, object]]:
        """
        Yield auto-filter metadata dicts, one per sheet that has an <autoFilter>.

        Each dict has:
          - 'sheet': sheet name
          - 'ref': the autoFilter range reference (e.g. 'A1:M241')
          - 'columns': list of filter column dicts
        """
        for sheet_name, rid in self._sheets:
            zpath = self._paths.get(rid)
            if not zpath:
                continue
            try:
                ws_data = self._read_part(zpath)
            except FileNotFoundError:
                continue

            try:
                root = ET.fromstring(ws_data)
            except ET.ParseError:
                continue

            af_el = root.find(_q("autoFilter"))
            if af_el is None:
                continue

            yield _parse_auto_filter(af_el, sheet_name)

    def iter_vba_modules(self) -> Dict[str, str]:
        """Extract VBA module source code from the embedded vbaProject.bin.

        Returns a dict mapping module_name -> plain-text VBA source.
        Returns an empty dict if no VBA project is present.

        Requires no third-party dependencies — uses the bundled
        xlsb_reader._vba_reader module (stdlib only).
        """
        from xlsb_reader._vba_reader import read_vba_modules

        try:
            vba_bin = self._read_part("xl/vbaProject.bin")
        except (KeyError, FileNotFoundError):
            return {}

        try:
            return read_vba_modules(vba_bin)
        except Exception:
            return {}

    def close(self):
        self._zf.close()

    def __enter__(self):
        return self

    def __exit__(self, *_):
        self.close()
