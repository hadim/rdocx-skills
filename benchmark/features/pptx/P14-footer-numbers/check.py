"""P14 oracle: slides 2-6 each hold a slide-number placeholder with a slidenum field and a footer placeholder
reading "Northwind Confidential", each matching a placeholder of the slide's layout by type and idx (so it
inherits the layout's position and style); the title slide (slide 1) has neither and no slidenum field."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "lib"))
import oracle  # noqa: E402
from oracle import NS  # noqa: E402

FOOTER = "Northwind Confidential"


def check(out, inputs):
    pkg = oracle.Pkg(out)
    errors = pkg.integrity()
    slides = pkg.slides()
    if len(slides) != 6:
        return errors + [f"{len(slides)} slides"]
    for n, slide in enumerate(slides, 1):
        layout = pkg.layout_of(slide)
        lay_ph = {(pkg.ph(sh)["type"], pkg.ph(sh)["idx"]) for sh in pkg.shapes(layout) if pkg.ph(sh)}
        shapes = pkg.shapes(slide)
        num = [sh for sh in shapes if (pkg.ph(sh) or {}).get("type") == "sldNum"]
        ftr = [sh for sh in shapes if (pkg.ph(sh) or {}).get("type") == "ftr"]
        fields = pkg.xml(slide).findall(".//a:fld[@type='slidenum']", NS)
        if n == 1:
            if num or ftr or fields:
                errors.append("title slide has a slide number or a footer")
            continue
        if len(num) != 1 or num[0].find(".//a:fld[@type='slidenum']", NS) is None:
            errors.append(f"slide {n}: no slide-number placeholder with a slidenum field")
        elif (pkg.ph(num[0])["type"], pkg.ph(num[0])["idx"]) not in lay_ph:
            errors.append(f"slide {n}: slide-number placeholder idx {pkg.ph(num[0])['idx']} not on the layout")
        if len(ftr) != 1 or pkg.text(slide, ftr[0]).strip() != FOOTER:
            errors.append(f"slide {n}: no footer placeholder reading {FOOTER!r}")
        elif (pkg.ph(ftr[0])["type"], pkg.ph(ftr[0])["idx"]) not in lay_ph:
            errors.append(f"slide {n}: footer placeholder idx {pkg.ph(ftr[0])['idx']} not on the layout")
        ids = [pkg.shape_id(sh) for sh in shapes]
        if len(ids) != len(set(ids)):
            errors.append(f"slide {n}: repeated shape ids")
    return errors


oracle.main(check)
