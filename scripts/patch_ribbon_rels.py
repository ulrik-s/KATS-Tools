import sys
import xml.etree.ElementTree as ET
from pathlib import Path

rels_path = Path(sys.argv[1])

ns_uri = "http://schemas.openxmlformats.org/package/2006/relationships"
ET.register_namespace("", ns_uri)

tree = ET.parse(rels_path)
root = tree.getroot()

wanted_type = "http://schemas.microsoft.com/office/2006/relationships/ui/extensibility"
wanted_target = "customUI/customUI.xml"

for rel in root.findall(f"{{{ns_uri}}}Relationship"):
    if rel.get("Type") == wanted_type and rel.get("Target") == wanted_target:
        tree.write(rels_path, encoding="utf-8", xml_declaration=True)
        print(f"Ribbon relationship already present in {rels_path}")
        sys.exit(0)

used_ids = {rel.get("Id") for rel in root.findall(f"{{{ns_uri}}}Relationship")}
rid_num = 1
while f"rId{rid_num}" in used_ids:
    rid_num += 1

new_rel = ET.Element(f"{{{ns_uri}}}Relationship")
new_rel.set("Id", f"rId{rid_num}")
new_rel.set("Type", wanted_type)
new_rel.set("Target", wanted_target)
root.append(new_rel)

tree.write(rels_path, encoding="utf-8", xml_declaration=True)
print(f"Added ribbon relationship to {rels_path}")

