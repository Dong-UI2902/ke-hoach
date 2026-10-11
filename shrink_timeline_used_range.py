import os
import re
import tempfile
import zipfile
from lxml import etree as ET

PATH = r"D:\kehoach\InviBrowser_Business_Plan_Giu_Nguyen_Format_v2.xlsx"
SHEET_NAME = "Timeline GTM"
MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"


def column_number(reference: str) -> int:
    letters = re.match(r"[A-Z]+", reference).group(0)
    result = 0
    for letter in letters:
        result = result * 26 + ord(letter) - 64
    return result


handle, temp_path = tempfile.mkstemp(suffix=".xlsx", dir=os.path.dirname(PATH))
os.close(handle)
try:
    with zipfile.ZipFile(PATH, "r") as source:
        workbook = ET.fromstring(source.read("xl/workbook.xml"))
        relationships = ET.fromstring(source.read("xl/_rels/workbook.xml.rels"))
        targets = {
            relation.attrib["Id"]: relation.attrib["Target"]
            for relation in relationships.findall(f"{{{PKG_REL}}}Relationship")
        }
        sheet_node = next(
            sheet
            for sheet in workbook.findall(f".//{{{MAIN}}}sheet")
            if sheet.attrib["name"] == SHEET_NAME
        )
        relationship_id = sheet_node.attrib[f"{{{REL}}}id"]
        target = targets[relationship_id].lstrip("/")
        sheet_path = target if target.startswith("xl/") else f"xl/{target}"

        with zipfile.ZipFile(temp_path, "w", compression=zipfile.ZIP_DEFLATED) as output:
            for info in source.infolist():
                data = source.read(info.filename)
                if info.filename == sheet_path:
                    root = ET.fromstring(data)
                    dimension = root.find(f"{{{MAIN}}}dimension")
                    if dimension is not None:
                        dimension.set("ref", "A1:AQ9")
                    sheet_data = root.find(f"{{{MAIN}}}sheetData")
                    for row in list(sheet_data):
                        row_number = int(row.attrib["r"])
                        if row_number > 9:
                            sheet_data.remove(row)
                            continue
                        for cell in list(row.findall(f"{{{MAIN}}}c")):
                            if column_number(cell.attrib["r"]) > 43:
                                row.remove(cell)
                    columns = root.find(f"{{{MAIN}}}cols")
                    if columns is not None:
                        for column in list(columns):
                            minimum = int(column.attrib["min"])
                            maximum = int(column.attrib["max"])
                            if minimum > 43:
                                columns.remove(column)
                            elif maximum > 43:
                                column.set("max", "43")
                    data = ET.tostring(root, encoding="utf-8", xml_declaration=True)
                output.writestr(info, data)
    os.replace(temp_path, PATH)
finally:
    if os.path.exists(temp_path):
        os.remove(temp_path)

print("SHRUNK Timeline GTM TO A1:AQ9")
