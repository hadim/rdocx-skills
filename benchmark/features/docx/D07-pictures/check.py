"""D07 oracle: the new picture after the anchor at 10 cm wide, the old picture showing the replacement at its
old size, no picture showing the original image, unique drawing ids."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from authoring_oracle import NS, W, Check, Pkg, picture_targets, png_size  # noqa: E402

pkg, ck = Pkg(sys.argv[1]), Check()
body = [p for p in pkg.body if p.tag == W + "p"]
texts = [pkg.text(p) for p in body]
pics = picture_targets(pkg, pkg.body)
sizes = [(png_size(pkg.read(t)) if t else None, ext) for _, t, ext in pics]
ck.require(len(pics) == 2, f"{len(pics)} pictures in the body, expected 2: {sizes}")
k = texts.index("Figure placement anchor.") if "Figure placement anchor." in texts else None
if ck.require(k is not None and k + 1 < len(body), "anchor paragraph missing"):
    new = picture_targets(pkg, body[k + 1])
    if ck.require(len(new) == 1, "the paragraph after the anchor holds no picture"):
        _, target, (cx, cy) = new[0]
        ck.require(png_size(pkg.read(target)) == (300, 200), "the inserted picture is not new.png")
        ck.require(abs(cx - 3600000) <= 36000 and abs(cy - 2400000) <= 24000,
                   f"inserted picture is {cx} x {cy} EMU, expected 3600000 x 2400000 (10 x 6.67 cm)")
c = texts.index("Figure 1: site plan.") if "Figure 1: site plan." in texts else None
if ck.require(c is not None, "caption missing"):
    old = picture_targets(pkg, body[c - 1])
    if ck.require(len(old) == 1, "the picture before the caption is gone"):
        _, target, ext = old[0]
        ck.require(png_size(pkg.read(target)) == (200, 150), "the old picture does not show replacement.png")
        ck.require(ext == (2880000, 2160000), f"the old picture's size changed: {ext}")
ck.require(not any(s == (400, 300) for s, _ in sizes), "a picture still shows the original image")
ids = [d.get("id") for d in pkg.body.iter("{%s}docPr" % NS["wp"])]
ck.require(len(ids) == len(set(ids)), f"duplicate drawing ids {ids}")
orphan = [n for n in pkg.names if n.startswith("word/media/") and png_size(pkg.read(n)) == (400, 300)]
ck.info(f"original image part left in the package: {bool(orphan)}")
ck.done()
