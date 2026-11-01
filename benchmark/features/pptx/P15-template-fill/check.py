"""P15 oracle: the main part has the presentation (not template) content type; 3 slides on the layouts Title
Slide, Title and Content, Picture with Caption; every text sits in a placeholder of the slide's layout (title,
subtitle, body, caption: matched by type and idx), none in a new text box; the picture slide's picture is the
layout's picture placeholder (ph type pic) showing photo.png, cropped to the placeholder's aspect ratio (within
2 %, no distortion); the instruction slide is gone (no slide text, no orphan part)."""
import struct
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "lib"))
import oracle  # noqa: E402
from oracle import NS, q  # noqa: E402

WANT = [("Title Slide", {"title": "Northwind annual review", "sub": "Fiscal year 2026"}),
        ("Title and Content", {"title": "Highlights",
                               "body": "Revenue up 12 %\nTwo new regions\nCustomer churn down to 4 %"}),
        ("Picture with Caption", {"title": "Our new headquarters", "body": "Opening March 2027"})]
PRES_CT = "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"


def check(out, inputs):
    pkg = oracle.Pkg(out)
    errors = pkg.integrity()
    if pkg.content_type(pkg.main) != PRES_CT:
        errors.append(f"main part content type {pkg.content_type(pkg.main)}")
    slides = pkg.slides()
    if len(slides) != 3:
        return errors + [f"{len(slides)} slides"]
    for k, (slide, (layout_name, texts)) in enumerate(zip(slides, WANT), 1):
        layout = pkg.layout_of(slide)
        if pkg.layout_name(layout) != layout_name:
            errors.append(f"slide {k} on layout {pkg.layout_name(layout)!r}")
            continue
        lay = {(pkg.ph(sh)["type"], pkg.ph(sh)["idx"]): sh for sh in pkg.shapes(layout) if pkg.ph(sh)}
        found = {}
        for sh in pkg.shapes(slide):
            text = pkg.text(slide, sh).strip()
            if not text:
                continue
            ph = pkg.ph(sh)
            if ph is None:
                errors.append(f"slide {k}: text {text[:30]!r} outside any placeholder")
                continue
            if (ph["type"], ph["idx"]) not in lay and not (ph["type"] == "obj" and ("body", ph["idx"]) in lay) \
                    and not any(t == ph["type"] for t, _ in lay if ph["idx"] == "0"):
                errors.append(f"slide {k}: placeholder {ph} has no match on the layout")
            kind = "title" if ph["type"] in ("title", "ctrTitle") else "sub" if ph["type"] == "subTitle" else "body"
            found[kind] = text
        for kind, want in texts.items():
            if found.get(kind) != want:
                errors.append(f"slide {k}: {kind} {found.get(kind)!r}, expected {want!r}")
        if k == 3:
            pics = [sh for sh in pkg.shapes(slide) if sh.tag == q("p:pic")]
            if len(pics) != 1 or (pkg.ph(pics[0]) or {}).get("type") != "pic":
                errors.append("slide 3: no picture filling the picture placeholder")
                continue
            pic = pics[0]
            photo = (inputs / "photo.png").read_bytes()
            if pkg.read(pkg.rel(slide, pic.find("p:blipFill/a:blip", NS).get(q("r:embed")))) != photo:
                errors.append("slide 3: the placeholder does not show photo.png")
            ext = pic.find("p:spPr/a:xfrm/a:ext", NS)
            if ext is None:
                ext = lay[("pic", pkg.ph(pic)["idx"])].find("p:spPr/a:xfrm/a:ext", NS)
            w, h = struct.unpack(">II", photo[16:24])
            src = pic.find("p:blipFill/a:srcRect", NS)
            crop = {s: int(src.get(s, 0)) / 100000 if src is not None else 0 for s in "lrtb"}
            seen = (w * (1 - crop["l"] - crop["r"])) / (h * (1 - crop["t"] - crop["b"]))
            box = int(ext.get("cx")) / int(ext.get("cy"))
            if abs(seen - box) / box > 0.02:
                errors.append(f"slide 3: picture distorted (image shows {seen:.3f}, frame {box:.3f})")
    for n in pkg.names:
        if n.startswith("ppt/slides/") and n.endswith(".xml") and b"How to use this template" in pkg.read(n):
            errors.append(f"instruction slide still in the package ({n})")
    if pkg.orphans():
        errors.append(f"orphan parts: {pkg.orphans()}")
    return errors


oracle.main(check)
