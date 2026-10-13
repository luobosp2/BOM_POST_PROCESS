"""Patch only note cells inside an XLSX, retaining all other ZIP members."""

from __future__ import annotations

from io import BytesIO
import posixpath
from xml.dom import Node, minidom
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from openpyxl.utils import get_column_letter, range_boundaries

from .errors import BomError


SHEET_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def _worksheet_part(archive: ZipFile, title: str) -> str:
    workbook = ET.fromstring(archive.read("xl/workbook.xml"))
    sheet = next((s for s in workbook.findall(f"{{{SHEET_NS}}}sheets/{{{SHEET_NS}}}sheet")
                  if s.get("name") == title), None)
    if sheet is None:
        raise BomError(f"找不到工作表：{title}")
    relation_id = sheet.get(f"{{{REL_NS}}}id")
    relations = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    for relation in relations:
        if relation.get("Id") == relation_id:
            target = relation.get("Target", "")
            return posixpath.normpath(target.lstrip("/") if target.startswith("/") else "xl/" + target)
    raise BomError(f"找不到工作表的数据文件：{title}")


def _children(node: Node, name: str) -> list:
    return [child for child in node.childNodes
            if child.nodeType == Node.ELEMENT_NODE
            and child.namespaceURI == SHEET_NS and child.localName == name]


def _patch_sheet(xml: bytes, column: int, values: dict[int, str],
                 reference_column: int, new_column: bool) -> bytes:
    with minidom.parseString(xml) as document:
        root = document.documentElement
        prefix = f"{root.prefix}:" if root.prefix else ""

        def element(name: str):
            return document.createElementNS(SHEET_NS, prefix + name)

        data = _children(root, "sheetData")[0]
        rows = {int(row.getAttribute("r")): row for row in _children(data, "row")}
        letter = get_column_letter(column)
        reference_letter = get_column_letter(reference_column)
        for number, value in values.items():
            row = rows[number]
            cells = _children(row, "c")
            address = f"{letter}{number}"
            cell = next((c for c in cells if c.getAttribute("r") == address), None)
            if cell is None:
                cell = element("c")
                cell.setAttribute("r", address)
                donor = next((c for c in cells if c.getAttribute("r") == f"{reference_letter}{number}"), None)
                if donor is not None and donor.hasAttribute("s"):
                    cell.setAttribute("s", donor.getAttribute("s"))
                following = next((c for c in cells
                                  if range_boundaries(c.getAttribute("r"))[0] > column), None)
                row.insertBefore(cell, following if following is not None else (cells[-1].nextSibling if cells else row.firstChild))
            for child in list(cell.childNodes):
                if (child.nodeType == Node.ELEMENT_NODE and child.namespaceURI == SHEET_NS
                        and child.localName in {"f", "v", "is"}):
                    cell.removeChild(child)
                    child.unlink()
            cell.setAttribute("t", "inlineStr")
            inline = element("is")
            text = element("t")
            text.setAttribute("xml:space", "preserve")
            text.appendChild(document.createTextNode(value))
            inline.appendChild(text)
            cell.insertBefore(inline, cell.firstChild)
            if row.hasAttribute("spans"):
                bounds = [int(part) for span in row.getAttribute("spans").split() for part in span.split(":")]
                row.setAttribute("spans", f"{min(*bounds, column)}:{max(*bounds, column)}")

        if new_column:
            dimensions = _children(root, "dimension")
            if dimensions:
                left, top, right, bottom = range_boundaries(dimensions[0].getAttribute("ref"))
                dimensions[0].setAttribute("ref", f"{get_column_letter(left)}{top}:{get_column_letter(max(right, column))}{bottom}")
            # Only size the added column. Existing widths, row heights and styles stay intact.
            existing_cols = _children(root, "cols")
            cols = existing_cols[0] if existing_cols else element("cols")
            if not existing_cols:
                root.insertBefore(cols, data)
            if not any(int(c.getAttribute("min")) <= column <= int(c.getAttribute("max")) for c in _children(cols, "col")):
                added = element("col")
                for key, value in {"min": str(column), "max": str(column), "width": "42", "customWidth": "1"}.items():
                    added.setAttribute(key, value)
                cols.appendChild(added)
        return document.toxml(encoding="UTF-8", standalone=document.standalone)


def update_notes(source: bytes, sheet: str, column: int, values: dict[int, str],
                 reference_column: int, *, new_column: bool) -> bytes:
    if not values:
        return source
    result = BytesIO()
    with ZipFile(BytesIO(source)) as original, ZipFile(result, "w") as output:
        target = _worksheet_part(original, sheet)
        for entry in original.infolist():
            data = original.read(entry.filename)
            if entry.filename == target:
                data = _patch_sheet(data, column, values, reference_column, new_column)
            output.writestr(entry, data)
        output.comment = original.comment
    return result.getvalue()
