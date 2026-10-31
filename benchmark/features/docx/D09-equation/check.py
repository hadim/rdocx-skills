"""D09 oracle: a display equation (m:oMathPara) right after the anchor, built of a fraction whose numerator holds
a radical over b^2 - 4ac, with the quadratic formula's text."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from authoring_oracle import NS, W, Check, Pkg, q  # noqa: E402
from lxml import etree  # noqa: E402

pkg, ck = Pkg(sys.argv[1]), Check()
body = [p for p in pkg.body if p.tag == W + "p"]
texts = [pkg.text(p) for p in body]
k = texts.index("The roots are given by:") if "The roots are given by:" in texts else None
if ck.require(k is not None and k + 1 < len(body), "anchor paragraph missing"):
    math = [p for p in body[k:k + 2] if p.find(".//m:oMathPara", NS) is not None]
    if ck.require(math, "no display equation (m:oMathPara) in or right after the anchor"):
        para_ = math[0].find(".//m:oMathPara", NS)
        text = "".join(t.text or "" for t in para_.iter(q("m:t"))).replace("\u2212", "-").replace(" ", "")
        ck.require(text == "x=-b\u00b1b2-4ac2a", f"equation text is {text!r}")
        f = para_.find(".//m:f", NS)
        if ck.require(f is not None, "no fraction"):
            num, den = f.find("m:num", NS), f.find("m:den", NS)
            rad = num.find(".//m:rad", NS) if num is not None else None
            ck.require(rad is not None, "no radical in the numerator")
            if rad is not None:
                sup = rad.find(".//m:sSup", NS)
                ck.require(sup is not None and "".join(t.text for t in sup.find("m:sup", NS).iter(q("m:t"))) == "2",
                           "no b squared under the radical")
                deg = rad.find("m:deg", NS)
                hidden = rad.find("m:radPr/m:degHide", NS)
                ck.require(hidden is not None or deg is None or not "".join(deg.itertext()).strip(),
                           "the radical shows a degree")
                # ECMA-376 Part 1, 22.1.2.88: m:rad is (m:radPr?, m:deg, m:e); without m:deg, LibreOffice reads
                # the radicand as the degree and draws empty boxes
                kids = [etree.QName(c).localname for c in rad if etree.QName(c).localname != "radPr"]
                ck.require(kids == ["deg", "e"], f"m:rad holds {kids}, the schema requires deg then e")
            ck.require("".join(t.text or "" for t in den.iter(q("m:t"))) == "2a", "denominator is not 2a")
    ck.require(texts[k + 2 if math and math[0] is not body[k] else k + 1].startswith("where"), "text after moved")
ck.done()
