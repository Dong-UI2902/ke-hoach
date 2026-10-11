from zipfile import ZipFile
from lxml import etree as ET

path = r"D:\kehoach\InviBrowser_Business_Plan_Giu_Nguyen_Format_v2.pre_remove_revenue_backup.xlsx"
ns = {
    "m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "pr": "http://schemas.openxmlformats.org/package/2006/relationships",
}

with ZipFile(path) as z:
    wb = ET.fromstring(z.read("xl/workbook.xml"))
    rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    sheet = wb.xpath('.//m:sheet[@name="KH doanh thu và P&L"]', namespaces=ns)[0]
    rid = sheet.get("{%s}id" % ns["r"])
    rel = rels.xpath(f'.//pr:Relationship[@Id="{rid}"]', namespaces=ns)[0]
    target = rel.get("Target")
    print("RID", rid, "TARGET", target)
    sheet_path = "xl/" + target.lstrip("/")
    rels_path = "xl/worksheets/_rels/" + sheet_path.rsplit("/", 1)[1] + ".rels"
    print("SHEET_PATH", sheet_path)
    print("HAS_SHEET_RELS", rels_path in z.namelist(), rels_path)
    if rels_path in z.namelist():
        print(z.read(rels_path).decode("utf-8"))
