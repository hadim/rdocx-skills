"""Neutral oracle helpers for the review tasks R01-R15: zipfile, lxml and the standard library only.

Text views follow the OOXML meaning of tracked changes: the accepted view keeps w:ins content and drops
w:del content, a deleted paragraph mark joins its paragraph to the next one; the rejected view does the
opposite. PNG pixels are decoded with zlib (no image library) and PDF pages are counted from the page
objects, object streams included.
"""
from __future__ import annotations

import json
import re
import sys
import zipfile
import zlib
from pathlib import Path

from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
W15 = "http://schemas.microsoft.com/office/word/2012/wordml"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS = {"w": W, "w14": W14, "w15": W15, "r": R,
      "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
      "rel": "http://schemas.openxmlformats.org/package/2006/relationships"}
REVISION_TAGS = ("ins", "del", "moveFrom", "moveTo", "rPrChange", "pPrChange", "sectPrChange", "tblPrChange",
                 "trPrChange", "tcPrChange", "numberingChange", "cellIns", "cellDel", "cellMerge")


def q(tag):
    pfx, local = tag.split(":")
    return f"{{{NS[pfx]}}}{local}"


class Docx:
    def __init__(self, path):
        self.path = Path(path)
        with zipfile.ZipFile(self.path) as z:
            self.files = {n: z.read(n) for n in z.namelist()}
        self._xml = {}

    def xml(self, name):
        name = name.lstrip("/")
        if name not in self._xml:
            self._xml[name] = etree.fromstring(self.files[name]) if name in self.files else None
        return self._xml[name]

    @property
    def document(self):
        return self.xml(self.main_part())

    def main_part(self):
        rels = etree.fromstring(self.files["_rels/.rels"])
        for r in rels:
            if r.get("Type", "").endswith("/officeDocument"):
                return r.get("Target").lstrip("/")
        return "word/document.xml"

    def rels(self, part="word/document.xml"):
        folder, base = part.rsplit("/", 1)
        name = f"{folder}/_rels/{base}.rels"
        out = {}
        if name in self.files:
            for r in etree.fromstring(self.files[name]):
                tgt = r.get("Target")
                if r.get("TargetMode") != "External":
                    tgt = tgt.lstrip("/") if tgt.startswith("/") else _join(folder, tgt)
                out[r.get("Id")] = (r.get("Type").rsplit("/", 1)[-1], tgt)
        return out

    def part_by_type(self, rel_suffix, part="word/document.xml"):
        for _rid, (typ, tgt) in self.rels(part).items():
            if typ == rel_suffix:
                return tgt
        return None


def _join(folder, target):
    parts = folder.split("/") + target.split("/")
    out = []
    for p in parts:
        if p == "..":
            out.pop()
        elif p and p != ".":
            out.append(p)
    return "/".join(out)


# ------------------------------------------------------------------------------------------- text
def _in(el, tag, stop):
    """True if `el` has an ancestor `tag` below `stop`."""
    a = el.getparent()
    while a is not None and a is not stop:
        if a.tag == q(tag):
            return True
        a = a.getparent()
    return False


def run_text(p, view="accepted"):
    """Visible text of paragraph `p` in a view (field results kept, field codes dropped)."""
    out = []
    for el in p.iter():
        tag = el.tag
        if tag == q("w:fldChar"):
            continue
        if tag in (q("w:t"), q("w:delText"), q("w:tab"), q("w:br"), q("w:cr"), q("w:noBreakHyphen")):
            if _in(el, "w:pPr", p) or _in(el, "w:rPr", p):
                continue
            deleted = _in(el, "w:del", p) or _in(el, "w:moveFrom", p)
            inserted = _in(el, "w:ins", p) or _in(el, "w:moveTo", p)
            if view == "accepted" and deleted:
                continue
            if view == "rejected" and inserted:
                continue
            if tag in (q("w:t"), q("w:delText")):
                out.append(el.text or "")
            elif tag == q("w:tab"):
                out.append("\t")
            elif tag == q("w:noBreakHyphen"):
                out.append("-")
            elif tag == q("w:br") and el.get(q("w:type")) in (None, "textWrapping"):
                out.append("\n")
    return "".join(out)


def mark_revision(p):
    """'ins', 'del' or None: the tracked state of the paragraph mark of `p`."""
    rpr = p.find("w:pPr/w:rPr", NS)
    if rpr is None:
        return None
    if rpr.find("w:del", NS) is not None or rpr.find("w:moveFrom", NS) is not None:
        return "del"
    if rpr.find("w:ins", NS) is not None or rpr.find("w:moveTo", NS) is not None:
        return "ins"
    return None


def style_of(p):
    s = p.find("w:pPr/w:pStyle", NS)
    return s.get(q("w:val")) if s is not None else None


def paragraphs(container, view="accepted", tables=True):
    """[(style, text)] of the paragraphs of a story element in document order, in a view. A paragraph
    whose mark is removed in that view joins the next paragraph (the next one's style survives)."""
    items = []

    def walk(el):
        for child in el:
            if child.tag == q("w:p"):
                items.append(child)
            elif child.tag == q("w:tbl"):
                if tables:
                    for tc in child.iter(q("w:tc")):
                        for cp in tc.findall("w:p", NS):
                            items.append(cp)
            elif child.tag in (q("w:sdt"), q("w:sdtContent"), q("w:customXml"), q("w:ins"), q("w:del")):
                walk(child)

    walk(container)
    out, carry = [], ""
    for p in items:
        text = carry + run_text(p, view)
        mr = mark_revision(p)
        gone = (view == "accepted" and mr == "del") or (view == "rejected" and mr == "ins")
        if gone:
            carry = text
            continue
        carry = ""
        out.append((style_of(p) or "Normal", text))
    if carry:
        out.append(("Normal", carry))
    return out


def norm(expected):
    """Expected (style, text) pairs with the default paragraph style named, as paragraphs() gives them."""
    return [(s or "Normal", t) for s, t in expected]


def body(doc, view="accepted", tables=True):
    return paragraphs(doc.document.find("w:body", NS), view, tables)


def texts(doc, view="accepted", tables=True):
    return [t for _, t in body(doc, view, tables)]


def revisions(root):
    """Every revision element under `root` (a part's XML root)."""
    if root is None:
        return []
    return [el for el in root.iter() if isinstance(el.tag, str) and el.tag.startswith(f"{{{W}}}")
            and el.tag.split("}")[1] in REVISION_TAGS]


def story_parts(doc):
    """Names of the parts of every story of the main document."""
    names = [doc.main_part()]
    for _rid, (typ, tgt) in doc.rels(doc.main_part()).items():
        if typ in ("header", "footer", "footnotes", "endnotes", "comments"):
            names.append(tgt)
    return names


def all_revisions(doc):
    out = []
    for name in story_parts(doc):
        out += [(name, el) for el in revisions(doc.xml(name))]
    return out


# ------------------------------------------------------------------------------------------ comments
def comments(doc):
    """{id: dict(author, date, initials, text, para_ids)} from the comments part."""
    name = doc.part_by_type("comments", doc.main_part())
    root = doc.xml(name) if name else None
    out = {}
    if root is None:
        return out
    for c in root.findall("w:comment", NS):
        ps = c.findall(".//w:p", NS)
        out[int(c.get(q("w:id")))] = {
            "author": c.get(q("w:author")), "date": c.get(q("w:date")), "initials": c.get(q("w:initials")),
            "text": "\n".join(run_text(p) for p in ps).strip(),
            "para_ids": [p.get(q("w14:paraId")) for p in ps],
        }
    return out


def comments_ex(doc):
    """{paraId: (paraIdParent, done)} from commentsExtended."""
    name = doc.part_by_type("commentsExtended", doc.main_part())
    root = doc.xml(name) if name else None
    if root is None:
        return {}
    return {e.get(q("w15:paraId")): (e.get(q("w15:paraIdParent")), e.get(q("w15:done")))
            for e in root.findall("w15:commentEx", NS)}


def anchored_text(doc, cid, part=None):
    """Accepted-view text between commentRangeStart and commentRangeEnd of comment `cid` (paragraphs
    joined with a newline); None when the comment has no range."""
    root = doc.xml(part or doc.main_part())
    sid = str(cid)
    inside, buf, found = False, [], False
    for p in root.iter(q("w:p")):
        if p.find(".//w:p", NS) is not None:
            continue  # only innermost paragraphs
        line = []
        for el in p.iter():
            if el.tag == q("w:commentRangeStart") and el.get(q("w:id")) == sid:
                inside, found = True, True
                continue
            if el.tag == q("w:commentRangeEnd") and el.get(q("w:id")) == sid:
                if inside:
                    buf.append("".join(line))
                return "\n".join(buf) if found else None
            if inside and el.tag in (q("w:t"), q("w:tab")) and not _in(el, "w:del", p):
                line.append(el.text or "" if el.tag == q("w:t") else "\t")
        if inside:
            buf.append("".join(line))
    return None


def has_reference(doc, cid):
    return any(el.get(q("w:id")) == str(cid) for el in doc.document.iter(q("w:commentReference")))


# --------------------------------------------------------------------------------------- formatting
def _style_map(doc):
    root = doc.xml("word/styles.xml")
    return {s.get(q("w:styleId")): s for s in root.findall("w:style", NS)} if root is not None else {}


def _rpr_props(rpr):
    props = {}
    if rpr is None:
        return props
    c = rpr.find("w:color", NS)
    if c is not None:
        props["color"] = c.get(q("w:val")).upper()
    f = rpr.find("w:rFonts", NS)
    if f is not None and f.get(q("w:ascii")):
        props["font"] = f.get(q("w:ascii"))
    for tag in ("b", "i"):
        e = rpr.find(f"w:{tag}", NS)
        if e is not None:
            props[tag] = e.get(q("w:val"), "true") not in ("0", "false", "off")
    sz = rpr.find("w:sz", NS)
    if sz is not None:
        props["size"] = sz.get(q("w:val"))
    return props


def effective_run(doc, p, r):
    """Effective color / font / bold / italic / size of run `r` in paragraph `p`: docDefaults, then the
    paragraph style chain, then the character style chain, then direct formatting."""
    styles = _style_map(doc)
    props = {}
    root = doc.xml("word/styles.xml")
    if root is not None:
        props.update(_rpr_props(root.find("w:docDefaults/w:rPrDefault/w:rPr", NS)))

    def chain(sid):
        seen, out = set(), []
        while sid and sid in styles and sid not in seen:
            seen.add(sid)
            out.append(styles[sid])
            b = styles[sid].find("w:basedOn", NS)
            sid = b.get(q("w:val")) if b is not None else None
        return list(reversed(out))

    pstyle = style_of(p)
    if pstyle is None:
        pstyle = next((k for k, s in styles.items() if s.get(q("w:type")) == "paragraph"
                       and s.get(q("w:default")) in ("1", "true")), None)
    for s in chain(pstyle):
        props.update(_rpr_props(s.find("w:rPr", NS)))
    rpr = r.find("w:rPr", NS)
    if rpr is not None:
        rs = rpr.find("w:rStyle", NS)
        if rs is not None:
            for s in chain(rs.get(q("w:val"))):
                props.update(_rpr_props(s.find("w:rPr", NS)))
        props.update(_rpr_props(rpr))
    return props


def numbering_format(doc, p):
    """(numFmt, lvlText) of a numbered paragraph's level, through its own numPr or its style's."""
    num = p.find("w:pPr/w:numPr", NS)
    styles = _style_map(doc)
    sid = style_of(p)
    seen = set()
    while num is None and sid and sid in styles and sid not in seen:
        seen.add(sid)
        num = styles[sid].find("w:pPr/w:numPr", NS)
        b = styles[sid].find("w:basedOn", NS)
        sid = b.get(q("w:val")) if b is not None else None
    if num is None or num.find("w:numId", NS) is None:
        return None
    num_id = num.find("w:numId", NS).get(q("w:val"))
    ilvl_el = num.find("w:ilvl", NS)
    ilvl = ilvl_el.get(q("w:val")) if ilvl_el is not None else "0"
    name = doc.part_by_type("numbering", doc.main_part())
    root = doc.xml(name) if name else None
    if root is None:
        return None
    n = next((n for n in root.findall("w:num", NS) if n.get(q("w:numId")) == num_id), None)
    if n is None:
        return None
    aid = n.find("w:abstractNumId", NS).get(q("w:val"))
    a = next((a for a in root.findall("w:abstractNum", NS) if a.get(q("w:abstractNumId")) == aid), None)
    if a is None:
        return None
    lvl = next((lv for lv in a.findall("w:lvl", NS) if lv.get(q("w:ilvl")) == ilvl), None)
    if lvl is None:
        return None
    return (lvl.find("w:numFmt", NS).get(q("w:val")), lvl.find("w:lvlText", NS).get(q("w:val")))


# ------------------------------------------------------------------------------------- PNG and PDF
def png_pixels(path):
    """(width, height, fraction of non-white pixels, raw RGB bytes sample) of an 8-bit PNG."""
    data = Path(path).read_bytes()
    if data[:8] != b"\x89PNG\r\n\x1a\n":
        raise ValueError(f"{path}: not a PNG")
    pos, idat, width = 8, [], None
    palette = None
    while pos < len(data):
        n = int.from_bytes(data[pos:pos + 4], "big")
        kind = data[pos + 4:pos + 8]
        chunk = data[pos + 8:pos + 8 + n]
        if kind == b"IHDR":
            width, height = int.from_bytes(chunk[0:4], "big"), int.from_bytes(chunk[4:8], "big")
            depth, ctype = chunk[8], chunk[9]
        elif kind == b"PLTE":
            palette = chunk
        elif kind == b"IDAT":
            idat.append(chunk)
        pos += 12 + n
    if depth != 8:
        raise ValueError(f"{path}: bit depth {depth} not handled")
    bpp = {0: 1, 2: 3, 3: 1, 4: 2, 6: 4}[ctype]
    raw = zlib.decompress(b"".join(idat))
    stride = width * bpp
    prev = bytearray(stride)
    nonwhite = 0
    rows = []
    for y in range(height):
        f = raw[y * (stride + 1)]
        line = bytearray(raw[y * (stride + 1) + 1:(y + 1) * (stride + 1)])
        for i in range(stride):
            a = line[i - bpp] if i >= bpp else 0
            b = prev[i]
            c = prev[i - bpp] if i >= bpp else 0
            if f == 1:
                line[i] = (line[i] + a) & 255
            elif f == 2:
                line[i] = (line[i] + b) & 255
            elif f == 3:
                line[i] = (line[i] + ((a + b) >> 1)) & 255
            elif f == 4:
                pa, pb, pc = abs(b - c), abs(a - c), abs(a + b - 2 * c)
                line[i] = (line[i] + (a if pa <= pb and pa <= pc else b if pb <= pc else c)) & 255
        prev = line
        rows.append(bytes(line))
    for line in rows:
        for x in range(0, stride, bpp):
            if ctype in (2, 6):
                r, g, b_ = line[x], line[x + 1], line[x + 2]
                alpha = line[x + 3] if ctype == 6 else 255
            elif ctype == 3:
                k = line[x]
                r, g, b_ = palette[3 * k:3 * k + 3]
                alpha = 255
            else:
                r = g = b_ = line[x]
                alpha = line[x + 1] if ctype == 4 else 255
            if alpha > 0 and min(r, g, b_) < 200:
                nonwhite += 1
    return width, height, nonwhite / (width * height), rows


def png_differs(a, b):
    """Fraction of rows that differ between two PNGs of the same size (1.0 if the sizes differ)."""
    wa, ha, _, ra = png_pixels(a)
    wb, hb, _, rb = png_pixels(b)
    if (wa, ha) != (wb, hb):
        return 1.0
    return sum(x != y for x, y in zip(ra, rb)) / ha


def pdf_info(path):
    """(pages, has_marks, content digest) of a PDF: page objects counted in the file and in its
    object streams; marks are text-showing, path-painting or image operators in content streams."""
    data = Path(path).read_bytes()
    if not data.startswith(b"%PDF"):
        raise ValueError(f"{path}: not a PDF")
    chunks = [data]
    for m in re.finditer(rb"stream\r?\n", data):
        start = m.end()
        end = data.find(b"endstream", start)
        if end < 0:
            continue
        try:
            chunks.append(zlib.decompress(data[start:end]))
        except zlib.error:
            try:
                chunks.append(zlib.decompressobj().decompress(data[start:end]))
            except zlib.error:
                pass
    blob = b"\n".join(chunks)
    pages = len(re.findall(rb"/Type\s*/Page(?![a-zA-Z])", blob))
    marks = bool(re.search(rb"(\bTJ\b|\bTj\b|\bDo\b|\bre\s+f\b)", blob))
    return pages, marks, zlib.crc32(b"".join(sorted(chunks[1:]))) if len(chunks) > 1 else 0


# -------------------------------------------------------------------------------------------- report
def report(checks: dict, details=None):
    """Print the check record as JSON and exit 0 when every check passed, 1 otherwise."""
    ok = all(bool(v) for v in checks.values())
    print(json.dumps({"ok": ok, "checks": {k: bool(v) for k, v in checks.items()}, "details": details or {}},
                     ensure_ascii=False, default=str))
    sys.exit(0 if ok else 1)


def guard(fn):
    """Run a check function; a crash on a malformed output is a failed check, not a traceback."""
    try:
        fn()
    except SystemExit:
        raise
    except Exception as e:  # noqa: BLE001
        report({"readable": False}, {"error": f"{type(e).__name__}: {e}"})
