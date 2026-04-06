"""
strip_xlsb_metadata.py
======================
Strips identifying metadata from XLSB (and XLSX/XLSM) files in-place:
  - docProps/core.xml : <cp:lastModifiedBy>
  - docProps/app.xml  : <Application>, <AppVersion>

Intended to be called from a pre-commit git hook.

Usage:
    python strip_xlsb_metadata.py <file1.xlsb> [file2.xlsb ...]

Exit code 0 always — failures are reported but don't block the commit.
"""

import io
import sys
import zipfile
import xml.etree.ElementTree as ET

CP_NS = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
AP_NS = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"

ET.register_namespace("cp", CP_NS)
ET.register_namespace("dc", "http://purl.org/dc/elements/1.1/")
ET.register_namespace("dcterms", "http://purl.org/dc/terms/")
ET.register_namespace("xsi", "http://www.w3.org/2001/XMLSchema-instance")
ET.register_namespace("", AP_NS)


def _blank_tags(tree: ET.Element, ns: str, tag_names: list[str]) -> bool:
    """Blank the text of each named tag. Returns True if any were changed."""
    changed = False
    for name in tag_names:
        el = tree.find(f"{{{ns}}}{name}")
        if el is not None and el.text:
            el.text = ""
            changed = True
    return changed


def _serialise(tree: ET.Element) -> bytes:
    xml_body = ET.tostring(tree, encoding="unicode", xml_declaration=False)
    return (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        + xml_body.encode()
    )


def strip_metadata(path: str) -> bool:
    """Return True if the file was modified."""
    try:
        with zipfile.ZipFile(path, "r") as zf:
            names = zf.namelist()
            all_items = [(n, zf.read(n), zf.getinfo(n)) for n in names]
    except Exception as e:
        print(f"  [warn] could not read {path}: {e}", file=sys.stderr)
        return False

    changed_files: dict[str, bytes] = {}

    # --- core.xml ---
    if "docProps/core.xml" in names:
        tree = ET.fromstring(dict((n, d) for n, d, _ in all_items)["docProps/core.xml"])
        if _blank_tags(tree, CP_NS, ["lastModifiedBy"]):
            changed_files["docProps/core.xml"] = _serialise(tree)

    # --- app.xml ---
    if "docProps/app.xml" in names:
        tree = ET.fromstring(dict((n, d) for n, d, _ in all_items)["docProps/app.xml"])
        if _blank_tags(tree, AP_NS, ["Application", "AppVersion"]):
            changed_files["docProps/app.xml"] = _serialise(tree)

    if not changed_files:
        return False

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as out:
        for name, data, info in all_items:
            out.writestr(info, changed_files.get(name, data))

    with open(path, "wb") as f:
        f.write(buf.getvalue())

    return True


def main() -> None:
    paths = sys.argv[1:]
    if not paths:
        return

    for path in paths:
        if strip_metadata(path):
            print(f"  stripped metadata from {path}")


if __name__ == "__main__":
    main()
