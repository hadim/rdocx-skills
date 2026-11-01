"""R09 oracle: the TOC field's result paragraphs, read in order, give (level, title, page) equal to the headings'
(level from the TOC1-3 style, title before the last tab, page after it); the headings themselves unchanged."""
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
sys.path.insert(0, str(Path(__file__).resolve().parent))
import review_oracle as o  # noqa: E402
from make_input import EXPECTED, HEADINGS  # noqa: E402


def toc_entries(d):
    """Paragraphs of the result of the first TOC field, as (level, title, page)."""
    paras = [p for p in d.document.iter(o.q("w:p"))]
    out, depth, inside = [], 0, False
    for p in paras:
        started_here = False
        for el in p.iter(o.q("w:fldChar"), o.q("w:instrText"), o.q("w:fldSimple")):
            if el.tag == o.q("w:instrText") and re.match(r"\s*TOC\b", el.text or "") and not inside:
                inside, depth, started_here = True, 1, True
            elif inside and el.tag == o.q("w:fldChar"):
                kind = el.get(o.q("w:fldCharType"))
                if kind == "begin" and not started_here:
                    depth += 1
                elif kind == "end":
                    depth -= 1
            started_here = False if el.tag == o.q("w:fldChar") and el.get(o.q("w:fldCharType")) == "separate" \
                else started_here
        if inside:
            text = o.run_text(p)
            if "\t" in text:
                title, page = text.rsplit("\t", 1)
                m = re.search(r"(\d+)\s*$", o.style_of(p) or "")
                out.append((int(m.group(1)) if m else None, title.split("\t")[-1].strip(), page.strip()))
            if depth == 0:
                break
    return out


def main(out):
    d = o.Docx(out)
    got = toc_entries(d)
    heads = [(int(s[-1]), t) for s, t in o.body(d) if s.startswith("Heading")]
    o.report({
        "entries_match_headings": [(lvl, t) for lvl, t, _ in got] == [(lvl, t) for lvl, t, _ in EXPECTED],
        "page_numbers": [p for _, _, p in got] == [str(p) for _, _, p in EXPECTED],
        "headings_unchanged": heads == [(lvl, t) for lvl, t, *_ in HEADINGS],
    }, {"toc": got})


o.guard(lambda: main(sys.argv[1]))
