from __future__ import annotations

import os
import re
import tempfile
import zipfile
import xml.etree.ElementTree as ET


WORKBOOK = r"D:\kehoach\InviBrowser_Business_Plan_Theo_Mau.xlsx"
HIDDEN_SHEETS = {
    "Tài chính chi tiết",
    "Doanh thu và P&L",
    "Từ khóa & đối thủ",
}
MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
PRICING_NOTES = {
    "J2": "Cơ chế trial và activation",
    "J3": "Chưa có pricing trong nguồn",
    "J4": "Ưu tiên permission và handoff",
    "J5": "Không dùng làm trục định vị",
    "J6": "Cần reliability, audit, SLA và support",
}
KEYWORD_WIDTHS = {
    1: 7,
    2: 24,
    3: 18,
    4: 18,
    5: 22,
    6: 18,
    7: 22,
    8: 22,
    9: 18,
    10: 16,
    11: 16,
    12: 20,
    13: 24,
}


def sheet_part_map(source: zipfile.ZipFile) -> dict[str, str]:
    workbook_root = ET.fromstring(source.read("xl/workbook.xml"))
    rels_root = ET.fromstring(source.read("xl/_rels/workbook.xml.rels"))
    rels = {
        rel.attrib["Id"]: rel.attrib["Target"]
        for rel in rels_root.findall(f"{{{PKG_REL_NS}}}Relationship")
    }
    result = {}
    for sheet in workbook_root.findall(f".//{{{MAIN_NS}}}sheet"):
        rel_id = sheet.attrib[f"{{{REL_NS}}}id"]
        target = rels[rel_id].lstrip("/")
        result[sheet.attrib["name"]] = (
            target if target.startswith("xl/") else f"xl/{target}"
        )
    return result


def remove_formula_and_value(cell: ET.Element) -> None:
    for child_name in ("f", "v", "is"):
        child = cell.find(f"{{{MAIN_NS}}}{child_name}")
        if child is not None:
            cell.remove(child)
    cell.attrib.pop("t", None)


def set_inline_string(cell: ET.Element, value: str) -> None:
    remove_formula_and_value(cell)
    cell.set("t", "inlineStr")
    inline = ET.SubElement(cell, f"{{{MAIN_NS}}}is")
    text = ET.SubElement(inline, f"{{{MAIN_NS}}}t")
    text.text = value


def clean_imported_formulas(data: bytes, sheet_name: str) -> bytes:
    root = ET.fromstring(data)
    for cell in root.findall(f".//{{{MAIN_NS}}}c"):
        coord = cell.attrib.get("r", "")
        match = re.fullmatch(r"([A-Z]+)(\d+)", coord)
        if not match:
            continue
        col, row_text = match.groups()
        row = int(row_text)
        if sheet_name == "Kế hoạch tài chính" and coord in {"D7", "D8"}:
            remove_formula_and_value(cell)
        elif sheet_name == "Bảng giá & gói dịch vụ" and 2 <= row <= 46:
            if coord in PRICING_NOTES:
                set_inline_string(cell, PRICING_NOTES[coord])
            elif col == "J" or (col == "I" and row >= 7):
                remove_formula_and_value(cell)
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def clean_keyword_columns(data: bytes) -> bytes:
    root = ET.fromstring(data)
    cols = root.find(f"{{{MAIN_NS}}}cols")
    if cols is None:
        sheet_data = root.find(f"{{{MAIN_NS}}}sheetData")
        cols = ET.Element(f"{{{MAIN_NS}}}cols")
        root.insert(list(root).index(sheet_data), cols)
    for col in list(cols):
        start = int(col.attrib.get("min", "0"))
        end = int(col.attrib.get("max", "0"))
        if start <= 13 and end >= 1:
            cols.remove(col)
    for index, width in KEYWORD_WIDTHS.items():
        ET.SubElement(
            cols,
            f"{{{MAIN_NS}}}col",
            {
                "min": str(index),
                "max": str(index),
                "width": str(width),
                "customWidth": "1",
            },
        )
    ordered = sorted(list(cols), key=lambda col: int(col.attrib.get("min", "0")))
    for col in list(cols):
        cols.remove(col)
    cols.extend(ordered)
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def force_white_default_fill(data: bytes) -> bytes:
    root = ET.fromstring(data)
    fills = root.find(f"{{{MAIN_NS}}}fills")
    if fills is None:
        return data
    white_fill_id = None
    for index, fill in enumerate(list(fills)):
        pattern = fill.find(f"{{{MAIN_NS}}}patternFill")
        if pattern is None or pattern.attrib.get("patternType") != "solid":
            continue
        foreground = pattern.find(f"{{{MAIN_NS}}}fgColor")
        if foreground is not None and foreground.attrib.get("rgb") == "FFFFFFFF":
            white_fill_id = index
            break
    if white_fill_id is None:
        fill = ET.SubElement(fills, f"{{{MAIN_NS}}}fill")
        pattern = ET.SubElement(
            fill, f"{{{MAIN_NS}}}patternFill", {"patternType": "solid"}
        )
        ET.SubElement(
            pattern, f"{{{MAIN_NS}}}fgColor", {"rgb": "FFFFFFFF"}
        )
        ET.SubElement(
            pattern, f"{{{MAIN_NS}}}bgColor", {"indexed": "64"}
        )
        white_fill_id = len(list(fills)) - 1
        fills.set("count", str(len(list(fills))))
    for container_name in ("cellStyleXfs", "cellXfs"):
        container = root.find(f"{{{MAIN_NS}}}{container_name}")
        if container is None:
            continue
        for xf in container.findall(f"{{{MAIN_NS}}}xf"):
            if xf.attrib.get("fillId", "0") == "0":
                xf.set("fillId", str(white_fill_id))
                xf.set("applyFill", "1")
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def main() -> None:
    ET.register_namespace("", MAIN_NS)
    fd, temp_path = tempfile.mkstemp(suffix=".xlsx", dir=os.path.dirname(WORKBOOK))
    os.close(fd)
    try:
        with zipfile.ZipFile(WORKBOOK, "r") as source, zipfile.ZipFile(
            temp_path, "w", compression=zipfile.ZIP_DEFLATED
        ) as target:
            sheet_parts = sheet_part_map(source)
            finance_part = sheet_parts["Kế hoạch tài chính"]
            pricing_part = sheet_parts["Bảng giá & gói dịch vụ"]
            keyword_part = sheet_parts["Từ khóa & đối thủ"]
            for info in source.infolist():
                data = source.read(info.filename)
                if info.filename == "xl/workbook.xml":
                    root = ET.fromstring(data)
                    for sheet in root.findall(f".//{{{MAIN_NS}}}sheet"):
                        if sheet.attrib.get("name") in HIDDEN_SHEETS:
                            sheet.set("state", "hidden")
                        else:
                            sheet.attrib.pop("state", None)
                    data = ET.tostring(
                        root, encoding="utf-8", xml_declaration=True
                    )
                elif info.filename == finance_part:
                    data = clean_imported_formulas(data, "Kế hoạch tài chính")
                elif info.filename == pricing_part:
                    data = clean_imported_formulas(
                        data, "Bảng giá & gói dịch vụ"
                    )
                elif info.filename == keyword_part:
                    data = clean_keyword_columns(data)
                elif info.filename == "xl/styles.xml":
                    data = force_white_default_fill(data)
                target.writestr(info, data)
        os.replace(temp_path, WORKBOOK)
    finally:
        if os.path.exists(temp_path):
            os.remove(temp_path)


if __name__ == "__main__":
    main()
