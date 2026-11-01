import json
import re
import sys
from xml.sax.saxutils import escape

path, data_path = sys.argv[1:3]
data = json.load(open(data_path))
xml = open(path, encoding="utf-8").read()
for key in ("client", "date", "total"):
    xml = xml.replace("{{%s}}" % key, escape(data[key]))
row = re.search(r"<w:tr[ >](?:(?!<w:tr[ >]).)*?\{\{name\}\}.*?</w:tr>", xml, re.S).group(0)
rows = ""
for item in data["items"]:
    new = row
    for key in ("name", "qty", "price"):
        new = new.replace("{{%s}}" % key, escape(item[key]))
    rows += new
xml = xml.replace(row, rows)
if "{{" in xml:
    sys.exit("placeholders left: " + ", ".join(sorted(set(re.findall(r"\{\{[^}]*\}\}", xml)))))
open(path, "w", encoding="utf-8").write(xml)
