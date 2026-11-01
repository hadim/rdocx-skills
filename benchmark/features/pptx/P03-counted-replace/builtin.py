# bench: tools=zipfile, lxml; processes=1; raw_xml=yes
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path

from lxml import etree

src, out, old, new, expect = sys.argv[1], sys.argv[2], sys.argv[3], sys.argv[4], int(sys.argv[5])
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
unpacked = Path(tempfile.mkdtemp())
zipfile.ZipFile(src).extractall(unpacked)
parts = sorted(unpacked.glob("ppt/slides/slide*.xml")) + sorted(unpacked.glob("ppt/notesSlides/notesSlide*.xml"))


def replace(p):
    runs = p.findall(A + "r")
    chars = [(k, c) for k, r in enumerate(runs) for c in (r.findtext(A + "t") or "")]
    text = "".join(c for _, c in chars)
    n, i, kept = 0, 0, []
    while i < len(chars):
        if text.startswith(old, i):  # the replacement takes the run where the match starts
            kept += [(chars[i][0], c) for c in new]
            i, n = i + len(old), n + 1
        else:
            kept.append(chars[i])
            i += 1
    if n:
        for k, r in enumerate(runs):
            t = "".join(c for j, c in kept if j == k)
            if not t:
                p.remove(r)
                continue
            r.find(A + "t").text = t
            if t != t.strip():
                r.find(A + "t").set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    return n


trees = {part: etree.parse(str(part)) for part in parts}
count = sum(replace(p) for tree in trees.values() for p in tree.iter(A + "p"))
if count != expect:
    sys.exit(f"refused: {count} occurrences of {old!r}, expected {expect}; nothing written")
for part, tree in trees.items():
    tree.write(str(part), xml_declaration=True, encoding="UTF-8", standalone=True)
with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
    z.write(unpacked / "[Content_Types].xml", "[Content_Types].xml")
    for f in sorted(unpacked.rglob("*")):
        if f.is_file() and f.name != "[Content_Types].xml":
            z.write(f, f.relative_to(unpacked).as_posix())
shutil.rmtree(unpacked)
