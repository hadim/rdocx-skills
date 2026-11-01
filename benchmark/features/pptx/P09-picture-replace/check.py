"""P09 oracle: on slide 3 the picture "Site photo" shows the bytes of new.png and keeps its xfrm, srcRect (crop),
outline, name and alt text; slide 2's picture still shows old.png; no part is orphaned. A stale relationship to the
old image left in slide 3's rels is reported in info (harmless)."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "lib"))
import oracle  # noqa: E402
from oracle import NS  # noqa: E402
from lxml import etree  # noqa: E402

EMBED = oracle.q("r:embed")


def pic(pkg, slide, name):
    return oracle.find_one(pkg, slide, lambda sh: sh.tag == oracle.q("p:pic") and pkg.shape_name(sh) == name)


def frame(p):
    keep = {}
    for path in ("p:spPr/a:xfrm", "p:blipFill/a:srcRect", "p:spPr/a:ln", "p:nvPicPr/p:cNvPr"):
        el = p.find(path, NS)
        keep[path] = oracle.canon(el)
    return keep


def check(out, inputs):
    src, pkg = oracle.Pkg(inputs / "deck.pptx"), oracle.Pkg(out)
    errors = pkg.integrity()
    old, new = (inputs / "old.png").read_bytes(), (inputs / "new.png").read_bytes()
    s2, s3 = pkg.slides()[1], pkg.slides()[2]
    p3 = pic(pkg, s3, "Site photo")
    if p3 is None:
        return errors + ["no picture named 'Site photo' on slide 3"]
    rid = p3.find("p:blipFill/a:blip", NS).get(EMBED)
    if pkg.read(pkg.rel(s3, rid)) != new:
        errors.append("slide 3 picture does not show new.png")
    before = frame(pic(src, src.slides()[2], "Site photo"))
    for k, v in frame(p3).items():
        if v != before[k]:
            errors.append(f"slide 3 picture {k} changed")
    p2 = pic(pkg, s2, "Overview photo")
    if pkg.read(pkg.rel(s2, p2.find("p:blipFill/a:blip", NS).get(EMBED))) != old:
        errors.append("slide 2 picture no longer shows the old image")
    orphans = pkg.orphans()
    if orphans:
        errors.append(f"orphan parts: {orphans}")
    used = {e.get(EMBED) for e in pkg.xml(s3).iter() if e.get(EMBED)}
    stale = [r["id"] for r in pkg.rels(s3) if r["type"].endswith("/image") and r["id"] not in used]
    return errors, {"stale_image_rels_on_slide3": stale}


oracle.main(check)
