"""Place comment CID's markers around PHRASE in an unpacked .docx (what comment.py asks for): find the body paragraph
whose runs hold the phrase, split the first and last runs at its edges (same formatting), then
commentRangeStart / commentRangeEnd around it and the reference run after it."""
import copy
import sys

from lxml import etree

root, cid, phrase = sys.argv[1:4]
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
q = lambda tag: f"{{{W}}}{tag}"  # noqa: E731
path = f"{root}/word/document.xml"
tree = etree.parse(path)


def text_of(r):
    return "".join(t.text or "" for t in r.iter(q("t")))


def split(run, at):
    """Split a run at a character offset of its w:t text (merge_runs.py can leave several w:t, a w:tab or a
    w:lastRenderedPageBreak in one run); return the right half. Same run properties on both sides."""
    right = copy.deepcopy(run)
    run.addnext(right)
    for side, r in (("left", run), ("right", right)):
        pos = 0
        for child in [c for c in r if c.tag != q("rPr")]:
            n = len(child.text or "") if child.tag == q("t") else 0
            keep = pos < at if side == "left" else (pos + n > at or (n == 0 and pos >= at))
            if not keep:
                r.remove(child)
            elif child.tag == q("t") and pos < at < pos + n:
                child.text = child.text[:at - pos] if side == "left" else child.text[at - pos:]
                child.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            pos += n
    return right


for p in tree.getroot().iter(q("p")):
    runs = [r for r in p.iterchildren(q("r"))]
    text = "".join(text_of(r) for r in runs)
    start = text.find(phrase)
    if start < 0:
        continue
    end, pos, covered = start + len(phrase), 0, []
    for r in runs:
        a, b = pos, pos + len(text_of(r))
        pos = b
        if b <= start or a >= end:
            continue
        if a < start:
            r, a = split(r, start - a), start
        if b > end:
            split(r, end - a)
        covered.append(r)
    covered[0].addprevious(etree.fromstring(f'<w:commentRangeStart xmlns:w="{W}" w:id="{cid}"/>'))
    covered[-1].addnext(etree.fromstring(
        f'<w:r xmlns:w="{W}"><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="{cid}"/></w:r>'))
    covered[-1].addnext(etree.fromstring(f'<w:commentRangeEnd xmlns:w="{W}" w:id="{cid}"/>'))
    tree.write(path, xml_declaration=True, encoding="UTF-8", standalone=True)
    break
else:
    sys.exit(f"phrase {phrase!r} not found in a paragraph's runs")
