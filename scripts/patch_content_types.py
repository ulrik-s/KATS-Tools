import sys
import xml.etree.ElementTree as ET
from pathlib import Path

ct_path = Path(sys.argv[1])

ns_uri = "http://schemas.openxmlformats.org/package/2006/content-types"
ET.register_namespace("", ns_uri)

tree = ET.parse(ct_path)
root = tree.getroot()

has_xml_default = False
for child in root.findall(f"{{{ns_uri}}}Default"):
    if child.get("Extension") == "xml":
        has_xml_default = True
        break

if not has_xml_default:
    new_default = ET.Element(f"{{{ns_uri}}}Default")
    new_default.set("Extension", "xml")
    new_default.set("ContentType", "application/xml")
    root.append(new_default)
    print(f"Added xml default content type to {ct_path}")
else:
    print(f"xml default content type already present in {ct_path}")

tree.write(ct_path, encoding="utf-8", xml_declaration=True)

