# bench: tools=add_slide.py, clean.py, zipfile, lxml, Pillow; processes=5; raw_xml=yes
import os
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

from lxml import etree
from PIL import Image

SCRIPTS = Path(os.environ["PPTX_SKILL"]) / "scripts"
NS = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main",
      "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
      "r": "http://schemas.openxmlformats.org/package/2006/relationships",
      "ct": "http://schemas.openxmlformats.org/package/2006/content-types"}
RT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
DECL = f'xmlns:a="{NS["a"]}" xmlns:p="{NS["p"]}" xmlns:r="{RT[:-1]}"'
template, photo, out = sys.argv[1], sys.argv[2], sys.argv[3]
u = Path(tempfile.mkdtemp())
zipfile.ZipFile(template).extractall(u)


def load(path):
    return etree.parse(str(path))


def save(tree, path):
    tree.write(str(path), xml_declaration=True, encoding="UTF-8", standalone=True)


layouts = {load(p).find("p:cSld", NS).get("name"): p.name for p in (u / "ppt/slideLayouts").glob("*.xml")}
# structure first: three slides from their layouts, then the instruction slide out of <p:sldIdLst>, then clean.py
made = []
for name in ("Title Slide", "Title and Content", "Picture with Caption"):
    res = subprocess.run([sys.executable, SCRIPTS / "add_slide.py", u, layouts[name]], capture_output=True, text=True,
                         check=True)
    made.append(u / re.search(r"Created (ppt/slides/\S+\.xml)", res.stdout).group(1))
pres = load(u / "ppt/presentation.xml")
lst = pres.find("p:sldIdLst", NS)
lst.remove(lst[0])
save(pres, u / "ppt/presentation.xml")
subprocess.run([sys.executable, SCRIPTS / "clean.py", u], capture_output=True, check=True)


def placeholder(sid, ph, lines, name):
    paras = "".join(f'<a:p><a:r><a:rPr lang="en-US"/><a:t>{line}</a:t></a:r></a:p>' for line in lines)
    return etree.fromstring(
        f'<p:sp {DECL}><p:nvSpPr><p:cNvPr id="{sid}" name="{name}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>'
        f'<p:nvPr>{ph}</p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/>{paras}</p:txBody></p:sp>')


def fill(slide, shapes):
    tree = load(slide)
    tree.find(".//p:spTree", NS).extend(shapes)
    save(tree, slide)


fill(made[0], [placeholder(2, '<p:ph type="ctrTitle"/>', ["Northwind annual review"], "Title 1"),
               placeholder(3, '<p:ph type="subTitle" idx="1"/>', ["Fiscal year 2026"], "Subtitle 2")])
fill(made[1], [placeholder(2, '<p:ph type="title"/>', ["Highlights"], "Title 1"),
               placeholder(3, '<p:ph idx="1"/>', ["Revenue up 12 %", "Two new regions", "Customer churn down to 4 %"],
                           "Content Placeholder 2")])
# the picture placeholder: the image as a new media part, cropped to the placeholder's aspect ratio
layout = load(u / "ppt/slideLayouts" / layouts["Picture with Caption"])
ext = next(sp for sp in layout.iterfind(".//p:sp", NS) if sp.find(".//p:ph[@type='pic']", NS) is not None) \
    .find("p:spPr/a:xfrm/a:ext", NS)
box = int(ext.get("cx")) / int(ext.get("cy"))
w, h = Image.open(photo).size
cut = int(50000 * (1 - min(box, w / h) / max(box, w / h)))
crop = f'l="{cut}" r="{cut}"' if w / h > box else f't="{cut}" b="{cut}"'
media = u / "ppt/media"
media.mkdir(exist_ok=True)
shutil.copy(photo, media / f"image{len(list(media.iterdir())) + 1}.png")
target = sorted(media.iterdir())[-1].name
rels_path = made[2].parent / "_rels" / (made[2].name + ".rels")
rels = load(rels_path)
etree.SubElement(rels.getroot(), f"{{{NS['r']}}}Relationship", Id="rId2", Type=RT + "image",
                 Target=f"../media/{target}")
save(rels, rels_path)
pic = etree.fromstring(
    f'<p:pic {DECL}><p:nvPicPr><p:cNvPr id="3" name="Picture Placeholder 2"/><p:cNvPicPr><a:picLocks noGrp="1" '
    f'noChangeAspect="1"/></p:cNvPicPr><p:nvPr><p:ph type="pic" idx="1"/></p:nvPr></p:nvPicPr><p:blipFill>'
    f'<a:blip r:embed="rId2"/><a:srcRect {crop}/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr/></p:pic>')
fill(made[2], [placeholder(2, '<p:ph type="title"/>', ["Our new headquarters"], "Title 1"), pic,
               placeholder(4, '<p:ph type="body" sz="half" idx="2"/>', ["Opening March 2027"], "Text Placeholder 3")])
ct = load(u / "[Content_Types].xml")  # a .pptx: the presentation content type, and PNG declared
main = ct.find("ct:Override[@PartName='/ppt/presentation.xml']", NS)
main.set("ContentType", main.get("ContentType").replace("template.main", "presentation.main"))
if ct.find("ct:Default[@Extension='png']", NS) is None:
    ct.getroot().insert(0, etree.Element(f"{{{NS['ct']}}}Default", Extension="png", ContentType="image/png"))
save(ct, u / "[Content_Types].xml")
with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
    z.write(u / "[Content_Types].xml", "[Content_Types].xml")
    for f in sorted(u.rglob("*")):
        if f.is_file() and f.name != "[Content_Types].xml":
            z.write(f, f.relative_to(u).as_posix())
shutil.rmtree(u)
