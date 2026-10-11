from __future__ import annotations

import os
import tempfile
import zipfile
from lxml import etree as ET


WORKBOOK = r"D:\kehoach\InviBrowser_Business_Plan_Giu_Nguyen_Format.xlsx"
TEMPLATE = r"C:\Users\Admin\Desktop\Bussiness Plan mẫu.xlsx"
HIDDEN_SHEETS = {
    "Kế hoạch tài chính Needly（Copy）",
    "KH doanh thu và P&L",
    "Phân tích từ khóa_Đối thủ cạnh tranh",
}
MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"


def sheet_part_map(source: zipfile.ZipFile) -> dict[str, str]:
    workbook_root = ET.fromstring(source.read("xl/workbook.xml"))
    rels_root = ET.fromstring(source.read("xl/_rels/workbook.xml.rels"))
    rels = {
        rel.attrib["Id"]: rel.attrib["Target"]
        for rel in rels_root.findall(f"{{{PKG_REL_NS}}}Relationship")
    }
    result: dict[str, str] = {}
    for sheet in workbook_root.findall(f".//{{{MAIN_NS}}}sheet"):
        target = rels[sheet.attrib[f"{{{REL_NS}}}id"]].lstrip("/")
        result[sheet.attrib["name"]] = target if target.startswith("xl/") else f"xl/{target}"
    return result


def remove_all_formulas(data: bytes, sheet_name: str) -> bytes:
    root = ET.fromstring(data)
    for cell in root.findall(f".//{{{MAIN_NS}}}c"):
        formula = cell.find(f"{{{MAIN_NS}}}f")
        if formula is None:
            continue
        cell.remove(formula)
        value = cell.find(f"{{{MAIN_NS}}}v")
        if value is not None:
            cell.remove(value)
        cell.attrib.pop("t", None)

    # J2 in the pricing sheet is a source-backed note. The imported template's
    # formula occupied that cell, so restore the text after removing formulas.
    if sheet_name == "Giá Needly đơn đầu":
        for cell in root.findall(f".//{{{MAIN_NS}}}c"):
            if cell.attrib.get("r") != "J2":
                continue
            for child in list(cell):
                if child.tag in {
                    f"{{{MAIN_NS}}}v",
                    f"{{{MAIN_NS}}}is",
                    f"{{{MAIN_NS}}}f",
                }:
                    cell.remove(child)
            cell.set("t", "inlineStr")
            inline = ET.SubElement(cell, f"{{{MAIN_NS}}}is")
            text = ET.SubElement(inline, f"{{{MAIN_NS}}}t")
            text.text = "Cơ chế trial/activation"

    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def clone(element: ET.Element) -> ET.Element:
    return ET.fromstring(ET.tostring(element, encoding="utf-8"))


def restore_template_format(data: bytes, template_data: bytes) -> bytes:
    """Restore layout and cell style references while retaining new values."""
    root = ET.fromstring(data)
    template_root = ET.fromstring(template_data)

    for name in ("sheetPr", "sheetViews", "sheetFormatPr", "cols", "mergeCells",
                 "printOptions", "pageMargins", "pageSetup", "headerFooter"):
        current = root.find(f"{{{MAIN_NS}}}{name}")
        original = template_root.find(f"{{{MAIN_NS}}}{name}")
        if current is not None and original is not None:
            index = list(root).index(current)
            root.remove(current)
            root.insert(index, clone(original))

    original_rows = {
        row.attrib.get("r"): row
        for row in template_root.findall(f".//{{{MAIN_NS}}}sheetData/{{{MAIN_NS}}}row")
    }
    original_cells = {
        cell.attrib.get("r"): cell
        for cell in template_root.findall(f".//{{{MAIN_NS}}}sheetData/{{{MAIN_NS}}}row/{{{MAIN_NS}}}c")
    }
    for row in root.findall(f".//{{{MAIN_NS}}}sheetData/{{{MAIN_NS}}}row"):
        original = original_rows.get(row.attrib.get("r"))
        if original is not None:
            row.attrib.clear()
            row.attrib.update(original.attrib)
        for cell in row.findall(f"{{{MAIN_NS}}}c"):
            original_cell = original_cells.get(cell.attrib.get("r"))
            if original_cell is not None and "s" in original_cell.attrib:
                cell.set("s", original_cell.attrib["s"])
            else:
                cell.attrib.pop("s", None)

    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def force_explicit_white_for_no_fill(data: bytes) -> bytes:
    """Make Excel's implicit white background explicit for reliable previewing."""
    root = ET.fromstring(data)
    fills = root.find(f"{{{MAIN_NS}}}fills")
    if fills is None:
        return data

    white_fill_id = None
    for index, fill in enumerate(list(fills)):
        pattern = fill.find(f"{{{MAIN_NS}}}patternFill")
        if pattern is None or pattern.attrib.get("patternType") != "solid":
            continue
        fg = pattern.find(f"{{{MAIN_NS}}}fgColor")
        if fg is not None and fg.attrib.get("rgb") == "FFFFFFFF":
            white_fill_id = index
            break

    if white_fill_id is None:
        fill = ET.SubElement(fills, f"{{{MAIN_NS}}}fill")
        pattern = ET.SubElement(fill, f"{{{MAIN_NS}}}patternFill", {"patternType": "solid"})
        ET.SubElement(pattern, f"{{{MAIN_NS}}}fgColor", {"rgb": "FFFFFFFF"})
        ET.SubElement(pattern, f"{{{MAIN_NS}}}bgColor", {"indexed": "64"})
        white_fill_id = len(list(fills)) - 1
        fills.set("count", str(len(list(fills))))

    for group_name in ("cellStyleXfs", "cellXfs"):
        group = root.find(f"{{{MAIN_NS}}}{group_name}")
        if group is None:
            continue
        for xf in group.findall(f"{{{MAIN_NS}}}xf"):
            if xf.attrib.get("fillId", "0") == "0":
                xf.set("fillId", str(white_fill_id))
                xf.set("applyFill", "1")

    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def main() -> None:
    handle, temp_path = tempfile.mkstemp(suffix=".xlsx", dir=os.path.dirname(WORKBOOK))
    os.close(handle)
    try:
        with zipfile.ZipFile(WORKBOOK, "r") as source, zipfile.ZipFile(
            TEMPLATE, "r"
        ) as template_source, zipfile.ZipFile(temp_path, "w", compression=zipfile.ZIP_DEFLATED) as target:
            parts = sheet_part_map(source)
            part_to_name = {part: name for name, part in parts.items()}
            template_parts = sheet_part_map(template_source)
            for info in source.infolist():
                data = source.read(info.filename)
                if info.filename == "xl/workbook.xml":
                    root = ET.fromstring(data)
                    for sheet in root.findall(f".//{{{MAIN_NS}}}sheet"):
                        if sheet.attrib.get("name") in HIDDEN_SHEETS:
                            sheet.set("state", "hidden")
                        else:
                            sheet.attrib.pop("state", None)
                    data = ET.tostring(root, encoding="utf-8", xml_declaration=True)
                elif info.filename in part_to_name:
                    sheet_name = part_to_name[info.filename]
                    data = remove_all_formulas(data, sheet_name)
                    data = restore_template_format(
                        data, template_source.read(template_parts[sheet_name])
                    )
                elif info.filename == "xl/styles.xml":
                    data = force_explicit_white_for_no_fill(
                        template_source.read("xl/styles.xml")
                    )
                target.writestr(info, data)
        os.replace(temp_path, WORKBOOK)
    finally:
        if os.path.exists(temp_path):
            os.remove(temp_path)


if __name__ == "__main__":
    main()
