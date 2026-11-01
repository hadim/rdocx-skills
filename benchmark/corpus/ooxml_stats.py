"""Neutral counts of a .docx or .pptx, read with zipfile and lxml only (no stack's reader).

Self-contained on purpose: build.py runs it as ``python -I ooxml_stats.py FILE`` on downloaded
files (untrusted), so it imports nothing from this folder. The XML parser resolves no entity
and fetches nothing. Prints one JSON object. Author names in docProps/core.xml are not read.
"""

from __future__ import annotations

import json
import re
import sys
import zipfile

from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture"
V = "urn:schemas-microsoft-com:vml"
EP = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
P14 = "http://schemas.microsoft.com/office/powerpoint/2010/main"

PARSER = etree.XMLParser(resolve_entities=False, no_network=True, huge_tree=True, load_dtd=False)


def _xml(z: zipfile.ZipFile, name: str):
    return etree.fromstring(z.read(name), PARSER)


def _count(root, ns: str, local: str) -> int:
    return sum(1 for _ in root.iter(f"{{{ns}}}{local}"))


def _app(z: zipfile.ZipFile, names: set) -> dict:
    out = {}
    if "docProps/app.xml" in names:
        root = _xml(z, "docProps/app.xml")
        for key in ("Application", "AppVersion", "Pages", "Slides", "Words", "Notes", "Template"):
            el = root.find(f"{{{EP}}}{key}")
            if el is not None and el.text:
                out[key] = el.text.strip()
    return out


def docx_stats(z: zipfile.ZipFile, names: set) -> dict:
    doc = _xml(z, "word/document.xml")
    stories = [n for n in names if re.match(r"word/(document|header\d*|footer\d*|footnotes|endnotes)\.xml$", n)]
    rev = {"ins": 0, "del": 0, "move": 0, "format": 0}
    pictures = 0
    for n in stories:
        root = doc if n == "word/document.xml" else _xml(z, n)
        rev["ins"] += _count(root, W, "ins")
        rev["del"] += _count(root, W, "del")
        rev["move"] += _count(root, W, "moveFrom") + _count(root, W, "moveTo")
        rev["format"] += sum(_count(root, W, t) for t in ("rPrChange", "pPrChange", "tblPrChange", "trPrChange", "tcPrChange", "sectPrChange", "numberingChange"))
        pictures += _count(root, PIC, "pic") + _count(root, V, "imagedata")
    comments = _count(_xml(z, "word/comments.xml"), W, "comment") if "word/comments.xml" in names else 0
    footnotes = 0
    if "word/footnotes.xml" in names:
        footnotes = sum(1 for f in _xml(z, "word/footnotes.xml").iter(f"{{{W}}}footnote") if f.get(f"{{{W}}}type") is None)
    instr = " ".join(t.text or "" for t in doc.iter(f"{{{W}}}instrText")) + " " + " ".join(
        f.get(f"{{{W}}}instr", "") for f in doc.iter(f"{{{W}}}fldSimple"))
    body = doc.find(f"{{{W}}}body")
    sects = _count(doc, W, "sectPr")
    landscape = sum(1 for s in doc.iter(f"{{{W}}}pgSz") if s.get(f"{{{W}}}orient") == "landscape")
    return {
        "paragraphs": _count(doc, W, "p"),
        "body_blocks": len(body) if body is not None else 0,
        "tables": _count(doc, W, "tbl"),
        "pictures": pictures,
        "comments": comments,
        "revisions": sum(rev.values()),
        "revisions_by_kind": rev,
        "charts": len([n for n in names if re.match(r"word/charts/chart\d+\.xml$", n)]),
        "footnotes": footnotes,
        "sections": sects,
        "landscape_sections": landscape,
        "headers": len([n for n in names if re.match(r"word/header\d*\.xml$", n)]),
        "footers": len([n for n in names if re.match(r"word/footer\d*\.xml$", n)]),
        "has_toc": bool(re.search(r"\bTOC\b", instr)),
        "fields": sorted(set(re.findall(r"\b(PAGE|NUMPAGES|TOC|REF|PAGEREF|HYPERLINK|SEQ|DATE|STYLEREF|SECTIONPAGES)\b", instr))),
    }


def pptx_stats(z: zipfile.ZipFile, names: set) -> dict:
    pres = _xml(z, "ppt/presentation.xml")
    slides = sorted(n for n in names if re.match(r"ppt/slides/slide\d+\.xml$", n))
    paras = tables = pics = 0
    for n in slides:
        root = _xml(z, n)
        paras += _count(root, A, "p")
        tables += _count(root, A, "tbl")
        pics += _count(root, P, "pic")
    comments = 0
    for n in names:
        if re.match(r"ppt/comments/.*\.xml$", n):
            comments += sum(1 for el in _xml(z, n).iter() if isinstance(el.tag, str) and etree.QName(el).localname == "cm")
    return {
        "slides": _count(pres, P, "sldId"),
        "slide_parts": len(slides),
        "paragraphs": paras,
        "tables": tables,
        "pictures": pics,
        "comments": comments,
        "revisions": 0,
        "charts": len([n for n in names if re.match(r"ppt/charts/chart\d+\.xml$", n)]),
        "notes": len([n for n in names if re.match(r"ppt/notesSlides/notesSlide\d+\.xml$", n)]),
        "sections": _count(pres, P14, "section"),
        "layouts": len([n for n in names if re.match(r"ppt/slideLayouts/slideLayout\d+\.xml$", n)]),
        "masters": len([n for n in names if re.match(r"ppt/slideMasters/slideMaster\d+\.xml$", n)]),
    }


def fonts(z: zipfile.ZipFile, names: set) -> list[str]:
    """Typefaces the file names: the theme's major / minor latin fonts, Word's font table, and the
    latin typefaces of slides, layouts and masters (theme references like +mn-lt left out)."""
    found = set()
    for n in names:
        if re.match(r"(word|ppt)/theme/theme\d+\.xml$", n):
            root = _xml(z, n)
            for tag in ("majorFont", "minorFont"):
                for el in root.iter(f"{{{A}}}{tag}"):
                    lat = el.find(f"{{{A}}}latin")
                    if lat is not None and lat.get("typeface"):
                        found.add(lat.get("typeface"))
    if "word/fontTable.xml" in names:
        for f in _xml(z, "word/fontTable.xml").iter(f"{{{W}}}font"):
            if f.get(f"{{{W}}}name"):
                found.add(f.get(f"{{{W}}}name"))
    for n in names:
        if re.match(r"ppt/(slides/slide|slideLayouts/slideLayout|slideMasters/slideMaster)\d+\.xml$", n):
            for el in _xml(z, n).iter(f"{{{A}}}latin"):
                t = el.get("typeface", "")
                if t and not t.startswith("+"):
                    found.add(t)
    return sorted(found)


def stats(path: str) -> dict:
    with zipfile.ZipFile(path) as z:
        names = set(z.namelist())
        media = [i for i in z.infolist() if re.match(r"(word|ppt)/media/", i.filename)]
        out = {"parts": len(names), "media_files": len(media), "media_bytes": sum(i.file_size for i in media)}
        out["app"] = _app(z, names)
        out["fonts"] = fonts(z, names)
        if "word/document.xml" in names:
            out.update(docx_stats(z, names))
        elif "ppt/presentation.xml" in names:
            out.update(pptx_stats(z, names))
        else:
            out["error"] = "neither word/document.xml nor ppt/presentation.xml"
    return out


if __name__ == "__main__":
    try:
        print(json.dumps(stats(sys.argv[1])))
    except Exception as e:  # report, do not hide
        print(json.dumps({"error": f"{type(e).__name__}: {e}"}))
