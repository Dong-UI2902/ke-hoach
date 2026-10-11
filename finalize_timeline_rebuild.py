from __future__ import annotations

import os
import tempfile
import zipfile
from lxml import etree as ET


SOURCE = r"D:\kehoach\InviBrowser_Business_Plan_Giu_Nguyen_Format_v2.timeline_labels.xlsx"
TARGET = r"D:\kehoach\InviBrowser_Business_Plan_Giu_Nguyen_Format_v2.xlsx"
BACKUP = r"D:\kehoach\InviBrowser_Business_Plan_Giu_Nguyen_Format_v2.pre_timeline_labels_backup.xlsx"
MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"


def sheet_parts(archive: zipfile.ZipFile) -> dict[str, str]:
    workbook = ET.fromstring(archive.read("xl/workbook.xml"))
    relationships = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    targets = {
        relationship.attrib["Id"]: relationship.attrib["Target"]
        for relationship in relationships.findall(f"{{{PKG_REL}}}Relationship")
    }
    result = {}
    for sheet in workbook.findall(f".//{{{MAIN}}}sheet"):
        target = targets[sheet.attrib[f"{{{REL}}}id"]].lstrip("/")
        result[sheet.attrib["name"]] = target if target.startswith("xl/") else f"xl/{target}"
    return result


def clone(element: ET._Element) -> ET._Element:
    return ET.fromstring(ET.tostring(element, encoding="utf-8"))


def restore_layout(data: bytes, backup_data: bytes) -> bytes:
    root = ET.fromstring(data)
    backup_root = ET.fromstring(backup_data)
    for name in (
        "sheetPr", "sheetViews", "sheetFormatPr", "cols", "mergeCells",
        "printOptions", "pageMargins", "pageSetup", "headerFooter",
    ):
        current = root.find(f"{{{MAIN}}}{name}")
        original = backup_root.find(f"{{{MAIN}}}{name}")
        if current is not None and original is not None:
            index = list(root).index(current)
            root.remove(current)
            root.insert(index, clone(original))

    backup_rows = {
        row.attrib.get("r"): row
        for row in backup_root.findall(f".//{{{MAIN}}}sheetData/{{{MAIN}}}row")
    }
    backup_cells = {
        cell.attrib.get("r"): cell
        for cell in backup_root.findall(f".//{{{MAIN}}}sheetData/{{{MAIN}}}row/{{{MAIN}}}c")
    }
    for row in root.findall(f".//{{{MAIN}}}sheetData/{{{MAIN}}}row"):
        original_row = backup_rows.get(row.attrib.get("r"))
        if original_row is not None:
            row.attrib.clear()
            row.attrib.update(original_row.attrib)
        for cell in row.findall(f"{{{MAIN}}}c"):
            original_cell = backup_cells.get(cell.attrib.get("r"))
            if original_cell is not None and "s" in original_cell.attrib:
                cell.set("s", original_cell.attrib["s"])
            elif original_cell is not None:
                cell.attrib.pop("s", None)
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def normalize_timeline_view(data: bytes) -> bytes:
    root = ET.fromstring(data)
    sheet_view = root.find(f".//{{{MAIN}}}sheetView")
    if sheet_view is None:
        return data
    sheet_view.set("topLeftCell", "A1")
    for selection in list(sheet_view.findall(f"{{{MAIN}}}selection")):
        sheet_view.remove(selection)
    pane = sheet_view.find(f"{{{MAIN}}}pane")
    if pane is None:
        pane = ET.SubElement(sheet_view, f"{{{MAIN}}}pane")
    pane.attrib.clear()
    pane.attrib.update({
        "xSplit": "3",
        "ySplit": "2",
        "topLeftCell": "D3",
        "activePane": "bottomRight",
        "state": "frozen",
    })
    for attributes in (
        {"pane": "topRight", "activeCell": "D1", "sqref": "D1"},
        {"pane": "bottomLeft", "activeCell": "A3", "sqref": "A3"},
        {"pane": "bottomRight", "activeCell": "D3", "sqref": "D3"},
    ):
        ET.SubElement(sheet_view, f"{{{MAIN}}}selection", **attributes)
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


handle, temp_path = tempfile.mkstemp(suffix=".xlsx", dir=os.path.dirname(TARGET))
os.close(handle)
try:
    with zipfile.ZipFile(SOURCE, "r") as source, zipfile.ZipFile(
        BACKUP, "r"
    ) as backup, zipfile.ZipFile(temp_path, "w", compression=zipfile.ZIP_DEFLATED) as target:
        source_parts = sheet_parts(source)
        backup_parts = sheet_parts(backup)
        part_to_sheet = {part: name for name, part in source_parts.items()}
        for info in source.infolist():
            data = source.read(info.filename)
            if info.filename in part_to_sheet:
                sheet_name = part_to_sheet[info.filename]
                if sheet_name == "Timeline GTM":
                    data = normalize_timeline_view(data)
                elif sheet_name in backup_parts:
                    data = restore_layout(data, backup.read(backup_parts[sheet_name]))
            target.writestr(info, data)
    os.replace(temp_path, TARGET)
finally:
    if os.path.exists(temp_path):
        os.remove(temp_path)

print(f"FINALIZED {TARGET}")
