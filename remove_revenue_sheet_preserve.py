from pathlib import PurePosixPath
from zipfile import ZIP_DEFLATED, ZipFile

from lxml import etree as ET

SOURCE = r"D:\kehoach\InviBrowser_Business_Plan_Giu_Nguyen_Format_v2.pre_remove_revenue_backup.xlsx"
OUTPUT = r"D:\kehoach\InviBrowser_Business_Plan_Giu_Nguyen_Format_v2.xlsx"
SHEET_NAME = "KH doanh thu và P&L"

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
DOC_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
NS = {"m": MAIN_NS, "r": DOC_REL_NS, "pr": PKG_REL_NS, "ct": CT_NS}

with ZipFile(SOURCE, "r") as zin:
    payload = {name: zin.read(name) for name in zin.namelist()}

workbook = ET.fromstring(payload["xl/workbook.xml"])
workbook_rels = ET.fromstring(payload["xl/_rels/workbook.xml.rels"])
content_types = ET.fromstring(payload["[Content_Types].xml"])

sheets = workbook.xpath(".//m:sheets/m:sheet", namespaces=NS)
target_index = next(i for i, sheet in enumerate(sheets) if sheet.get("name") == SHEET_NAME)
target_sheet = sheets[target_index]
relationship_id = target_sheet.get(f"{{{DOC_REL_NS}}}id")
target_sheet.getparent().remove(target_sheet)

relationship = workbook_rels.xpath(
    f'.//pr:Relationship[@Id="{relationship_id}"]', namespaces=NS
)[0]
relationship_target = relationship.get("Target")
workbook_rels.remove(relationship)

if relationship_target.startswith("/"):
    worksheet_path = relationship_target.lstrip("/")
else:
    worksheet_path = str(PurePosixPath("xl") / relationship_target)

part_name = "/" + worksheet_path
for override in content_types.xpath(
    f'.//ct:Override[@PartName="{part_name}"]', namespaces=NS
):
    override.getparent().remove(override)

defined_names = workbook.find(f"{{{MAIN_NS}}}definedNames")
if defined_names is not None:
    for defined_name in list(defined_names):
        local_id = defined_name.get("localSheetId")
        text = defined_name.text or ""
        if SHEET_NAME in text or (local_id is not None and int(local_id) == target_index):
            defined_names.remove(defined_name)
        elif local_id is not None and int(local_id) > target_index:
            defined_name.set("localSheetId", str(int(local_id) - 1))
    if len(defined_names) == 0:
        workbook.remove(defined_names)

payload["xl/workbook.xml"] = ET.tostring(
    workbook, xml_declaration=True, encoding="UTF-8", standalone=True
)
payload["xl/_rels/workbook.xml.rels"] = ET.tostring(
    workbook_rels, xml_declaration=True, encoding="UTF-8", standalone=True
)
payload["[Content_Types].xml"] = ET.tostring(
    content_types, xml_declaration=True, encoding="UTF-8", standalone=True
)
payload.pop(worksheet_path, None)
payload.pop(
    str(PurePosixPath(worksheet_path).parent / "_rels" / (PurePosixPath(worksheet_path).name + ".rels")),
    None,
)

with ZipFile(OUTPUT, "w", compression=ZIP_DEFLATED, compresslevel=6) as zout:
    for name, data in payload.items():
        zout.writestr(name, data)

print(f"REMOVED {SHEET_NAME} -> {OUTPUT}")
