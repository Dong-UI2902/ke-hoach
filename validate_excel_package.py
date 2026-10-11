from __future__ import annotations

import posixpath
import zipfile
from lxml import etree as ET


PATH = r"D:\kehoach\InviBrowser_Business_Plan_Giu_Nguyen_Format_v2.xlsx"
MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"


with zipfile.ZipFile(PATH) as archive:
    names = set(archive.namelist())
    xml_names = [name for name in names if name.endswith((".xml", ".rels"))]

    for name in xml_names:
        root = ET.fromstring(archive.read(name))
        ignorable = root.attrib.get(f"{{{MC}}}Ignorable", "")
        for prefix in ignorable.split():
            assert prefix in root.nsmap, f"{name}: undeclared Ignorable prefix {prefix}"

    styles = ET.fromstring(archive.read("xl/styles.xml"))
    count_names = ("numFmts", "fonts", "fills", "borders", "cellStyleXfs", "cellXfs", "cellStyles", "dxfs")
    for name in count_names:
        node = styles.find(f"{{{MAIN}}}{name}")
        if node is not None and "count" in node.attrib:
            assert int(node.attrib["count"]) == len(node), f"styles.xml: wrong {name} count"

    fonts = styles.find(f"{{{MAIN}}}fonts")
    fills = styles.find(f"{{{MAIN}}}fills")
    borders = styles.find(f"{{{MAIN}}}borders")
    xfs = styles.find(f"{{{MAIN}}}cellXfs")
    assert fonts is not None and fills is not None and borders is not None and xfs is not None
    for xf in xfs:
        assert int(xf.attrib.get("fontId", "0")) < len(fonts)
        assert int(xf.attrib.get("fillId", "0")) < len(fills)
        assert int(xf.attrib.get("borderId", "0")) < len(borders)

    for name in names:
        if not name.startswith("xl/worksheets/") or not name.endswith(".xml"):
            continue
        root = ET.fromstring(archive.read(name))
        for cell in root.findall(f".//{{{MAIN}}}c"):
            assert int(cell.attrib.get("s", "0")) < len(xfs), f"{name}: invalid style index"

    for rel_name in [name for name in names if name.endswith(".rels")]:
        root = ET.fromstring(archive.read(rel_name))
        base_part = rel_name.replace("/_rels/", "/").removesuffix(".rels")
        base_dir = posixpath.dirname(base_part)
        for rel in root.findall(f"{{{PKG_REL}}}Relationship"):
            if rel.attrib.get("TargetMode") == "External":
                continue
            target = rel.attrib.get("Target", "")
            if target.startswith("/"):
                resolved = target.lstrip("/")
            else:
                resolved = posixpath.normpath(posixpath.join(base_dir, target))
            assert resolved in names, f"{rel_name}: missing target {resolved}"

print(f"PACKAGE_VALID xml_parts={len(xml_names)} styles={len(xfs)} files={len(names)}")
