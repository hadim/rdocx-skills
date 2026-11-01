import os
import re
import shutil
import sys

from PIL import Image

root, new_png, replacement_png = sys.argv[1:4]
doc_path = os.path.join(root, "word", "document.xml")
rels_path = os.path.join(root, "word", "_rels", "document.xml.rels")
doc = open(doc_path, encoding="utf-8").read()
rels = open(rels_path, encoding="utf-8").read()

# replace the existing picture's image: overwrite the media part its r:embed points at (the extent is kept)
rid = re.search(r'<a:blip r:embed="([^"]+)"', doc).group(1)
target = re.search(rf'Id="{rid}"[^>]*Target="([^"]+)"|Target="([^"]+)"[^>]*Id="{rid}"', rels)
shutil.copy(replacement_png, os.path.join(root, "word", target.group(1) or target.group(2)))

# insert new.png in a new paragraph after the anchor, 10 cm wide, height in proportion
shutil.copy(new_png, os.path.join(root, "word", "media", "inserted1.png"))
rels = rels.replace("</Relationships>", '<Relationship Id="rIdInserted1" Type="http://schemas.openxmlformats.org/'
                    'officeDocument/2006/relationships/image" Target="media/inserted1.png"/></Relationships>')
w, h = Image.open(new_png).size
cx, cy = 3600000, round(3600000 * h / w)
ids = [int(i) for i in re.findall(r'<wp:docPr id="(\d+)"', doc)]
pid = max(ids, default=0) + 1
pic = (f'<w:p><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="{cx}" cy="{cy}"/>'
       f'<wp:docPr id="{pid}" name="Picture {pid}"/><wp:cNvGraphicFramePr><a:graphicFrameLocks '
       'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/></wp:cNvGraphicFramePr>'
       '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData '
       'uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic '
       'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr>'
       f'<pic:cNvPr id="{pid}" name="inserted1.png"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill>'
       '<a:blip r:embed="rIdInserted1"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm>'
       f'<a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
       '</pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>')
anchor = re.search(r"<w:p[ >](?:(?!<w:p[ >]).)*?Figure placement .*?</w:p>", doc, re.S)
doc = doc[:anchor.end()] + pic + doc[anchor.end():]
open(doc_path, "w", encoding="utf-8").write(doc)
open(rels_path, "w", encoding="utf-8").write(rels)
