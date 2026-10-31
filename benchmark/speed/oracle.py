"""Neutral reads for the speed section: zipfile and lxml only, no stack's reader.

Picks, per corpus file, the inputs of the editing operations (the word to replace, the phrase to comment on,
sample strings a reader must return) and checks each operation's output cheaply (counts, anchors, revision
marks, pages). Selection rules keep every stack's documented path able to do the task, so that a time is
compared with a time: the word occurs only in paragraphs that every stack's replacement reaches (plain runs of
body paragraphs and table cells for Word; text runs of slides and of the notes body for PowerPoint), nowhere
else in the package.
"""
from __future__ import annotations

import posixpath
import re
import subprocess
import zipfile
from collections import Counter

from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
MC_FALLBACK = "{http://schemas.openxmlformats.org/markup-compatibility/2006}Fallback"
PARSER = etree.XMLParser(resolve_entities=False, no_network=True, huge_tree=True, load_dtd=False)
w = lambda tag: f"{{{W}}}{tag}"  # noqa: E731
a = lambda tag: f"{{{A}}}{tag}"  # noqa: E731
WORDS = re.compile(r"[^\W\d_]{5,}")
SAFE = re.compile(r"^[A-Za-z0-9 ]+$")
REVISION_TAGS = {w(t) for t in ("ins", "del", "moveFrom", "moveTo", "rPrChange", "pPrChange", "sectPrChange",
                                "tblPrChange", "trPrChange", "tcPrChange", "tblGridChange", "numberingChange",
                                "cellIns", "cellDel", "cellMerge")}
STORY = re.compile(r"word/(document|header\d*|footer\d*|footnotes|endnotes|comments)\.xml$")


def _xml(z: zipfile.ZipFile, name: str):
    return etree.fromstring(z.read(name), PARSER)


def norm(text: str) -> str:
    """Whitespace collapsed and Markdown escapes dropped: how a sample is looked for in a reader's output."""
    return re.sub(r"\s+", " ", text.replace("\\", "")).strip()


# ---------------------------------------------------------------------------------------------- Word
def _in_fallback(el) -> bool:
    return any(x.tag == MC_FALLBACK for x in el.iterancestors())


def _own_ts(p):
    """The w:t of paragraph p itself (not those of a text box anchored in it)."""
    for t in p.iter(w("t")):
        if next(t.iterancestors(w("p"))) is p:
            yield t


def docx_paragraphs(path: str):
    """[(part, p element, full text, plain text, eligible)] for every paragraph of every part holding w:t. Full: every
    w:t of the paragraph (tracked insertions, hyperlinks, fields, content controls included; deletions are
    w:delText and stay out). Plain: the w:t of runs that are direct children of the paragraph. Eligible: a paragraph
    of document.xml whose ancestors are only the body, tables, rows and cells."""
    out = []
    with zipfile.ZipFile(path) as z:
        for name in sorted(n for n in z.namelist() if n.startswith("word/") and n.endswith(".xml")):
            data = z.read(name)
            if b"<w:t" not in data and b":t>" not in data:
                continue
            root = etree.fromstring(data, PARSER)
            for p in root.iter(w("p")):
                if _in_fallback(p):
                    continue
                ts = list(_own_ts(p))
                full = "".join(t.text or "" for t in ts)
                plain = "".join(t.text or "" for t in ts if t.getparent().tag == w("r") and t.getparent().getparent() is p)
                eligible = name == "word/document.xml" and all(
                    x.tag in (w("body"), w("tbl"), w("tr"), w("tc")) for x in p.iterancestors() if x.tag != w("document"))
                out.append((name, p, full, plain, eligible))
    return out


def docx_other_text(path: str) -> str:
    """The text of the charts, diagrams and glossary parts (rdocx replace also edits a chart's cached labels)."""
    with zipfile.ZipFile(path) as z:
        return "\n".join("".join(_xml(z, n).itertext()) for n in z.namelist()
                         if re.match(r"word/(charts|diagrams|glossary)/.*\.xml$", n))


def choose_docx_word(path: str):
    """(word, count): the most frequent word of five letters or more that occurs only in plain runs of eligible
    paragraphs (every stack's replacement reaches those), and not in charts, diagrams or the glossary; None when
    no word occurs at least three times so."""
    paras = docx_paragraphs(path)
    other = docx_other_text(path)
    cand = Counter()
    for _, _, _, plain, ok in paras:
        if ok:
            cand.update(WORDS.findall(plain))
    for word, _ in sorted(cand.items(), key=lambda kv: (-kv[1], kv[0])):
        total, good = 0, word not in other
        for _, _, full, plain, ok in paras:
            n = full.count(word)
            if not n:
                continue
            if not ok or plain.count(word) != n:
                good = False
                break
            total += n
        if good and total >= 3:
            return word, total
    return None, 0


def replacement_for(word: str) -> str:
    return word.upper() if not word.isupper() else word.lower()


def docx_count(path: str, text: str, story_only: bool = True) -> int:
    """Occurrences of text in every paragraph of the story parts (the accepted view: w:t only)."""
    return sum(full.count(text) for name, _, full, _, _ in docx_paragraphs(path) if not story_only or STORY.match(name))


def _simple_body_paragraph(p) -> bool:
    if p.getparent().tag != w("body"):
        return False
    for child in p:
        if child.tag not in (w("pPr"), w("r"), w("proofErr"), w("bookmarkStart"), w("bookmarkEnd")):
            return False
        if child.tag == w("r"):
            kinds = [c.tag for c in child if c.tag != w("rPr")]
            if kinds.count(w("t")) != 1 or any(k not in (w("t"), w("lastRenderedPageBreak")) for k in kinds):
                return False
    return True


def choose_docx_phrase(path: str):
    """A phrase of four words, ASCII letters, digits and spaces, that occurs once in the whole document and lies in
    a simple body paragraph (runs of one w:t each, no field, revision, link or control), preferably across a run
    boundary, in the paragraph nearest the middle of the document. Returns (phrase, runs it spans) or (None, 0)."""
    paras = docx_paragraphs(path)
    alltext = "\n".join(x[2] for x in paras)
    body = [(p, plain) for name, p, full, plain, ok in paras if ok and _simple_body_paragraph(p) and full == plain]
    order = sorted(range(len(body)), key=lambda i: abs(i - len(body) / 2))
    fallback = None
    for i in order:
        p, text = body[i]
        bounds, pos = [], 0
        for t in _own_ts(p):
            pos += len(t.text or "")
            bounds.append(pos)
        for m in re.finditer(r"(?=(\b[A-Za-z0-9]+(?: [A-Za-z0-9]+){3}\b))", text):
            phrase, start = m.group(1), m.start(1)
            if not SAFE.match(phrase) or len(phrase) < 18 or alltext.count(phrase) != 1:
                continue
            runs = 1 + sum(1 for b in bounds if start < b < start + len(phrase))
            if runs >= 2:
                return phrase, runs
            fallback = fallback or (phrase, runs)
        if fallback:
            return fallback
    return None, 0


def docx_samples(path: str, n: int = 3):
    """n phrases of four to six words (ASCII letters) from single runs of eligible paragraphs, spread over the
    document: what a reader of the whole text must return. Title and Subtitle paragraphs are left out: pandoc moves
    them into its metadata, which -t markdown without --standalone does not print (NOTES.md)."""
    runs = []
    for _, p, _, _, ok in docx_paragraphs(path):
        style = p.find(f"{w('pPr')}/{w('pStyle')}")
        if not ok or (style is not None and "title" in (style.get(w("val")) or "").lower()):
            continue  # pandoc moves Title and Subtitle paragraphs into its metadata (NOTES.md)
        for r in p.iterchildren(w("r")):
            text = "".join(t.text or "" for t in r.iterchildren(w("t")))
            m = re.search(r"\b[A-Za-z]+(?: [A-Za-z]+){3,5}\b", text)
            if m:
                runs.append(m.group(0))
    runs = list(dict.fromkeys(runs))
    if not runs:
        return []
    return list(dict.fromkeys(runs[int(k * (len(runs) - 1) / max(1, n - 1))] for k in range(n)))


def comment_check(path: str, comment_text: str, phrase: str) -> dict:
    """Is there a comment with this text, anchored (range start to end, w:t in between) exactly on phrase?"""
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        if "word/comments.xml" not in names:
            return {"ok": False, "why": "no comments part"}
        ids = [c.get(w("id")) for c in _xml(z, "word/comments.xml").iter(w("comment"))
               if comment_text in "".join(t.text or "" for t in c.iter(w("t")))]
        if not ids:
            return {"ok": False, "why": "comment text not found"}
        doc = _xml(z, "word/document.xml")
    cid = ids[-1]
    inside, got, ref = False, [], False
    for el in doc.iter():
        if el.tag == w("commentRangeStart") and el.get(w("id")) == cid:
            inside = True
        elif el.tag == w("commentRangeEnd") and el.get(w("id")) == cid:
            inside = False
        elif el.tag == w("commentReference") and el.get(w("id")) == cid:
            ref = True
        elif inside and el.tag == w("t"):
            got.append(el.text or "")
    anchored = "".join(got)
    return {"ok": anchored == phrase and ref, "anchored": anchored, "reference": ref}


def revision_marks(path: str) -> int:
    with zipfile.ZipFile(path) as z:
        return sum(1 for n in z.namelist() if STORY.match(n) for el in _xml(z, n).iter() if el.tag in REVISION_TAGS)


# ---------------------------------------------------------------------------------------------- PowerPoint
def _rels(z, part):
    name = posixpath.join(posixpath.dirname(part), "_rels", posixpath.basename(part) + ".rels")
    if name not in z.namelist():
        return {}
    return {r.get("Id"): (r.get("Type").rsplit("/", 1)[-1], posixpath.normpath(posixpath.join(posixpath.dirname(part), r.get("Target"))))
            for r in _xml(z, name)}


def slide_parts(path: str):
    """[(slide part, notes part or None)] in presentation order."""
    with zipfile.ZipFile(path) as z:
        pres = _xml(z, "ppt/presentation.xml")
        rels = _rels(z, "ppt/presentation.xml")
        out = []
        for sld in pres.iter(f"{{{P}}}sldId"):
            part = rels[sld.get(f"{{{R}}}id")][1]
            notes = next((t for k, t in _rels(z, part).values() if k == "notesSlide"), None)
            out.append((part, notes))
        return out


def _a_paragraphs(root):
    for p in root.iter(a("p")):
        ts = [t for t in p.iter(a("t"))]
        full = "".join(t.text or "" for t in ts)
        plain = "".join(t.text or "" for t in ts if t.getparent().tag == a("r") and t.getparent().getparent() is p)
        yield p, full, plain


def _notes_body(p) -> bool:
    for sp in p.iterancestors(f"{{{P}}}sp"):
        ph = sp.find(f"{{{P}}}nvSpPr/{{{P}}}nvPr/{{{P}}}ph")
        return ph is not None and ph.get("type") == "body"
    return False


def pptx_paragraphs(path: str):
    """[(part, full, plain, eligible)] for every a:p of every XML part. Eligible: a paragraph of a slide, or of the
    notes body placeholder of a notes slide."""
    with zipfile.ZipFile(path) as z:
        slides = set()
        notes = set()
        for s, n in slide_parts(path):
            slides.add(s)
            if n:
                notes.add(n)
        out = []
        for name in sorted(n for n in z.namelist() if n.startswith("ppt/") and n.endswith(".xml")):
            data = z.read(name)
            if b":t>" not in data:
                continue
            root = etree.fromstring(data, PARSER)
            for p, full, plain in _a_paragraphs(root):
                ok = name in slides or (name in notes and _notes_body(p))
                out.append((name, full, plain, ok))
        return out


def choose_pptx_word(path: str):
    paras = pptx_paragraphs(path)
    cand = Counter()
    for _, _, plain, ok in paras:
        if ok:
            cand.update(WORDS.findall(plain))
    for word, _ in sorted(cand.items(), key=lambda kv: (-kv[1], kv[0])):
        total, good = 0, True
        for _, full, plain, ok in paras:
            n = full.count(word)
            if not n:
                continue
            if not ok or plain.count(word) != n:
                good = False
                break
            total += n
        if good and total >= 3:
            return word, total
    return None, 0


def pptx_count(path: str, text: str) -> int:
    """Occurrences in the slides and the notes slides."""
    names = {s for s, _ in slide_parts(path)} | {n for _, n in slide_parts(path) if n}
    return sum(full.count(text) for name, full, _, _ in pptx_paragraphs(path) if name in names)


def slide_texts(path: str):
    """[(slide text, notes text)] in presentation order, a:t joined by paragraph."""
    out = []
    with zipfile.ZipFile(path) as z:
        for s, n in slide_parts(path):
            st = "\n".join(full for _, full, _ in _a_paragraphs(_xml(z, s)))
            nt = "\n".join(full for p, full, _ in _a_paragraphs(_xml(z, n)) if _notes_body(p)) if n else ""
            out.append((st, nt))
    return out


def pptx_samples(path: str, n: int = 3):
    """Phrases of three or more ASCII words, each from one text run, from slides and from notes, spread over the
    deck (one run: no line break or field inside the phrase)."""
    found_slide, found_notes = [], []
    with zipfile.ZipFile(path) as z:
        for s, nts in slide_parts(path):
            for part, bucket in ((s, found_slide), (nts, found_notes)):
                if not part:
                    continue
                for r in _xml(z, part).iter(a("r")):
                    if part == nts and not _notes_body(r):
                        continue
                    m = re.search(r"\b[A-Za-z]+(?: [A-Za-z]+){2,5}\b", r.findtext(a("t")) or "")
                    if m:
                        bucket.append(m.group(0))
                        break
    pick = lambda xs, k: list(dict.fromkeys(xs[int(i * (len(xs) - 1) / max(1, k - 1))] for i in range(k))) if xs else []  # noqa: E731
    return pick(found_slide, n) + pick(found_notes, 1)


def middle_slide(path: str) -> int:
    """One-based number of the slide in the middle of the deck."""
    return len(slide_parts(path)) // 2 + 1


# ---------------------------------------------------------------------------------------------- PDF
def pdf_pages(pdf: str):
    r = subprocess.run(["pdfinfo", pdf], capture_output=True, text=True)
    m = re.search(r"^Pages:\s+(\d+)", r.stdout, re.M)
    return int(m.group(1)) if m else None
