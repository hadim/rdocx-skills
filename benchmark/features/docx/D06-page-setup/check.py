"""D06 oracle: section 2 landscape with 2 cm margins, section 1 unchanged, header and footer in both."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from authoring_oracle import W, Check, Pkg, norm  # noqa: E402

pkg, ck = Pkg(sys.argv[1]), Check()
sects = pkg.sect_prs()
if ck.require(len(sects) == 2, f"{len(sects)} sections, expected 2"):
    def page(s):
        pg, mar = s.find(W + "pgSz"), s.find(W + "pgMar")
        size = (int(pg.get(W + "w")), int(pg.get(W + "h")), pg.get(W + "orient") or "portrait")
        margins = tuple(int(mar.get(W + k)) for k in ("top", "right", "bottom", "left"))
        return size, margins

    (w1, h1, o1), m1 = page(sects[0])
    (w2, h2, o2), m2 = page(sects[1])
    ck.require((w1, h1, o1) == (11906, 16838, "portrait") and m1 == (1417, 1417, 1417, 1417),
               f"section 1 changed: {w1}x{h1} {o1}, margins {m1}")
    ck.require(o2 == "landscape" and abs(w2 - 16838) <= 2 and abs(h2 - 11906) <= 2,
               f"section 2 is not A4 landscape: {w2}x{h2} {o2}")
    ck.require(all(abs(m - 1134) <= 2 for m in m2), f"section 2 margins are {m2} twips, expected 1134 (2 cm)")
    for i in range(2):
        header = pkg.story_for(sects, i, "header")
        ck.require(header and "Confidential" in pkg.text(pkg.xml(header)), f"section {i + 1} shows no 'Confidential' header")
        footer = pkg.story_for(sects, i, "footer")
        if ck.require(footer, f"section {i + 1} shows no footer"):
            root = pkg.xml(footer)
            fields = [f.split()[0].upper() for f in pkg.fields(root) if f]
            ck.require(fields[:2] == ["PAGE", "NUMPAGES"], f"section {i + 1} footer fields are {fields}")
            text = norm(pkg.text(root))
            ck.require(text.startswith("Page ") and " of " in text, f"section {i + 1} footer reads {text!r}")
        ck.info(f"section {i + 1}: header {header}, footer {footer}")
ck.done()
