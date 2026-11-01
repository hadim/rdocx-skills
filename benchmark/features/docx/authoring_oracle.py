"""The neutral oracle's reader for the authoring tasks D01-D14: zipfile and lxml only, no stack's reader.

Each task's check.py imports it, checks what the task asked for, prints one line per problem and exits 1
when there is any (0 otherwise). Lines starting with "info:" describe the output and never fail it.
"""
from __future__ import annotations

import posixpath
import re
import sys
import zipfile

from lxml import etree

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "v": "urn:schemas-microsoft-com:vml",
    "pr": "http://schemas.openxmlformats.org/package/2006/relationships",
    "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
}
W = "{%s}" % NS["w"]
R_DOC = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
OFF = {"0", "false", "off"}


def q(tag: str) -> str:
    prefix, local = tag.split(":")
    return "{%s}%s" % (NS[prefix], local)


class Check:
    """Collects problems; `done()` prints them and exits."""

    def __init__(self):
        self.problems: list[str] = []
        self.infos: list[str] = []

    def require(self, cond, msg: str):
        if not cond:
            self.problems.append(msg)
        return bool(cond)

    def info(self, msg: str):
        self.infos.append(msg)

    def done(self):
        for m in self.infos:
            print(f"info: {m}")
        for m in self.problems:
            print(f"FAIL: {m}")
        if not self.problems:
            print("ok")
        sys.exit(1 if self.problems else 0)


class Pkg:
    """A .docx package read with zipfile and lxml."""

    def __init__(self, path):
        self.zip = zipfile.ZipFile(path)
        self.names = set(self.zip.namelist())
        self._xml: dict[str, etree._Element] = {}
        self.main = self._main_part()
        self.doc = self.xml(self.main)
        self.body = self.doc.find(q("w:body"))
        st = self.part_by_type("styles")
        self.styles = {s.get(W + "styleId"): s for s in self.xml(st).iter(W + "style")} if st else {}
        nb = self.part_by_type("numbering")
        self.numbering = self.xml(nb) if nb else None

    # --- package -------------------------------------------------------------------------
    def read(self, name: str) -> bytes:
        return self.zip.read(name.lstrip("/"))

    def xml(self, name: str):
        name = name.lstrip("/")
        if name not in self._xml:
            self._xml[name] = etree.fromstring(self.read(name))
        return self._xml[name]

    def _main_part(self) -> str:
        rels = etree.fromstring(self.read("_rels/.rels"))
        for rel in rels.iter("{%s}Relationship" % NS["pr"]):
            if rel.get("Type").endswith("/officeDocument"):
                return rel.get("Target").lstrip("/")
        raise SystemExit("FAIL: no main document relationship")

    def rels(self, part: str) -> dict[str, tuple[str, str, str]]:
        """{rId: (type suffix, target part name or URL, target mode)} of a part."""
        part = part.lstrip("/")
        d, b = posixpath.split(part)
        rels_name = posixpath.join(d, "_rels", b + ".rels")
        out = {}
        if rels_name not in self.names:
            return out
        for rel in etree.fromstring(self.read(rels_name)).iter("{%s}Relationship" % NS["pr"]):
            mode = rel.get("TargetMode") or ""
            target = rel.get("Target")
            if mode != "External":
                target = target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join(d, target))
            out[rel.get("Id")] = (rel.get("Type").rsplit("/", 1)[-1], target, mode)
        return out

    def part_by_type(self, suffix: str, part: str | None = None) -> str | None:
        for t, target, _ in self.rels(part or self.main).values():
            if t == suffix:
                return target
        return None

    def content_type(self, part: str) -> str | None:
        ct = etree.fromstring(self.read("[Content_Types].xml"))
        for o in ct.iter("{%s}Override" % NS["ct"]):
            if o.get("PartName").lstrip("/") == part.lstrip("/"):
                return o.get("ContentType")
        ext = part.rsplit(".", 1)[-1].lower()
        for d in ct.iter("{%s}Default" % NS["ct"]):
            if d.get("Extension").lower() == ext:
                return d.get("ContentType")
        return None

    # --- styles ----------------------------------------------------------------------------
    def style_name(self, style_id: str | None) -> str | None:
        s = self.styles.get(style_id)
        if s is None:
            return None
        n = s.find(W + "name")
        return n.get(W + "val") if n is not None else None

    def para_style(self, p) -> str | None:
        ps = p.find(f"{W}pPr/{W}pStyle")
        return ps.get(W + "val") if ps is not None else None

    def para_style_name(self, p) -> str:
        sid = self.para_style(p)
        if sid is None:
            default = next((k for k, s in self.styles.items() if s.get(W + "type") == "paragraph"
                            and s.get(W + "default") in ("1", "true")), None)
            return (self.style_name(default) or "Normal").lower()
        return (self.style_name(sid) or sid).lower()

    def style_chain(self, style_id):
        seen = []
        while style_id and style_id in self.styles and style_id not in seen:
            seen.append(style_id)
            b = self.styles[style_id].find(W + "basedOn")
            style_id = b.get(W + "val") if b is not None else None
        return [self.styles[s] for s in seen]

    def run_flag(self, r, flag: str) -> bool:
        """Effective toggle (b, i, u...) of a run: direct, character style, paragraph style chain."""
        def val(rpr):
            if rpr is None:
                return None
            el = rpr.find(W + flag)
            if el is None:
                return None
            v = el.get(W + "val")
            if flag == "u":
                return v not in ("none",) if v is not None else True
            return v not in OFF if v is not None else True

        got = val(r.find(W + "rPr"))
        if got is not None:
            return got
        rs = r.find(f"{W}rPr/{W}rStyle")
        if rs is not None:
            for s in self.style_chain(rs.get(W + "val")):
                got = val(s.find(W + "rPr"))
                if got is not None:
                    return got
        p = r.getparent()
        while p is not None and p.tag != W + "p":
            p = p.getparent()
        if p is not None:
            for s in self.style_chain(self.para_style(p)):
                got = val(s.find(W + "rPr"))
                if got is not None:
                    return got
        return False

    def run_fonts(self, r) -> str:
        """Every font name a run asks for, directly or through its styles (joined, lower case)."""
        names = []
        els = [r.find(W + "rPr")]
        rs = r.find(f"{W}rPr/{W}rStyle")
        if rs is not None:
            els += [s.find(W + "rPr") for s in self.style_chain(rs.get(W + "val"))]
        p = r.getparent()
        while p is not None and p.tag != W + "p":
            p = p.getparent()
        if p is not None:
            els += [s.find(W + "rPr") for s in self.style_chain(self.para_style(p))]
        for rpr in els:
            if rpr is None:
                continue
            f = rpr.find(W + "rFonts")
            if f is not None:
                names += [v for k, v in f.attrib.items() if not k.endswith("Theme")]
        return " ".join(names).lower()

    # --- numbering -------------------------------------------------------------------------
    def num_pr(self, p):
        """(numId, ilvl) of a paragraph, direct or through its style chain; None when not numbered."""
        np_ = p.find(f"{W}pPr/{W}numPr")
        if np_ is not None and np_.find(W + "numId") is not None:
            nid = np_.find(W + "numId").get(W + "val")
            il = np_.find(W + "ilvl")
            return (nid, int(il.get(W + "val")) if il is not None else 0) if nid != "0" else None
        for s in self.style_chain(self.para_style(p)):
            np_ = s.find(f"{W}pPr/{W}numPr")
            if np_ is not None and np_.find(W + "numId") is not None:
                nid = np_.find(W + "numId").get(W + "val")
                il = np_.find(W + "ilvl")
                return (nid, int(il.get(W + "val")) if il is not None else 0) if nid != "0" else None
        return None

    def num(self, num_id):
        if self.numbering is None:
            return None
        return next((n for n in self.numbering.iter(W + "num") if n.get(W + "numId") == str(num_id)), None)

    def abstract_of(self, num_id):
        n = self.num(num_id)
        if n is None:
            return None
        aid = n.find(W + "abstractNumId").get(W + "val")
        return next((a for a in self.numbering.iter(W + "abstractNum") if a.get(W + "abstractNumId") == aid), None)

    def num_format(self, p) -> str | None:
        """numFmt of a paragraph's list level ('bullet', 'decimal'...), None when not numbered."""
        np_ = self.num_pr(p)
        if np_ is None:
            return None
        a = self.abstract_of(np_[0])
        if a is None:
            return None
        lvl = next((l for l in a.iter(W + "lvl") if l.get(W + "ilvl") == str(np_[1])), None)
        if lvl is None:
            return None
        f = lvl.find(W + "numFmt")
        return f.get(W + "val") if f is not None else "decimal"

    # --- text ------------------------------------------------------------------------------
    def text(self, el) -> str:
        """Text of a paragraph (or any element): w:t, w:tab, w:br, skipping deleted text, text boxes and
        mc:Fallback copies (a text box is not part of the paragraph that anchors it)."""
        out = []

        def walk(e):
            tag = e.tag
            if tag in (q("mc:Fallback"), W + "txbxContent", W + "delText"):
                return
            if tag == W + "t":
                out.append(e.text or "")
            elif tag == W + "tab" and e.getparent().tag == W + "r":
                out.append("\t")
            elif tag in (W + "br", W + "cr") and e.getparent().tag == W + "r":
                out.append("\n")
            for c in e:
                walk(c)

        walk(el)
        return "".join(out)

    def paragraphs(self, root=None):
        """Every w:p under root (default the body), table cells included, text boxes and fallbacks excluded."""
        root = self.body if root is None else root
        out = []

        def walk(e):
            if e.tag in (q("mc:Fallback"), W + "txbxContent"):
                return
            if e.tag == W + "p":
                out.append(e)
            for c in e:
                walk(c)

        walk(root)
        return out

    def find_para(self, predicate, root=None):
        return next((p for p in self.paragraphs(root) if predicate(self.text(p))), None)

    def runs_with_text(self, p):
        """[(run element, text)] of the runs directly or indirectly under p (hyperlinks, ins...)."""
        out = []
        for r in p.iter(W + "r"):
            anc = r.getparent()
            skip = False
            while anc is not None and anc is not p:
                if anc.tag in (q("mc:Fallback"), W + "txbxContent", W + "del"):
                    skip = True
                anc = anc.getparent()
            if not skip:
                out.append((r, self.text(r)))
        return out

    def fields(self, root) -> list[str]:
        """Field instructions (simple and complex) under root, in order."""
        out, cur, depth = [], [], 0
        for e in root.iter(W + "fldSimple", W + "fldChar", W + "instrText"):
            if e.tag == W + "fldSimple":
                out.append(e.get(W + "instr", "").strip())
            elif e.tag == W + "fldChar":
                t = e.get(W + "fldCharType")
                if t == "begin":
                    depth += 1
                    cur.append([])
                elif t == "end" and cur:
                    depth -= 1
                    out.append("".join(cur.pop()).strip())
            elif cur:
                cur[-1].append(e.text or "")
        return out

    # --- sections ----------------------------------------------------------------------------
    def sect_prs(self):
        """Every section's w:sectPr in order (paragraph-carried ones, then the body's)."""
        out = [p.find(f"{W}pPr/{W}sectPr") for p in self.body.iter(W + "p")]
        out = [s for s in out if s is not None]
        last = self.body.find(W + "sectPr")
        if last is not None:
            out.append(last)
        return out

    def story_for(self, sect_prs, index, kind="footer", variant="default"):
        """Part name of the header/footer a section shows (inherited from earlier sections as Word does)."""
        for k in range(index, -1, -1):
            for ref in sect_prs[k].findall(f"{W}{kind}Reference"):
                if ref.get(W + "type", "default") == variant:
                    return self.rels(self.main)[ref.get(q("r:id"))][1]
        return None


def png_size(data: bytes):
    if data[:8] != b"\x89PNG\r\n\x1a\n":
        return None
    import struct
    return struct.unpack(">II", data[16:24])


def picture_targets(pkg: Pkg, root, part=None):
    """[(drawing element, media part name, extent (cx, cy))] of pictures under root."""
    rels = pkg.rels(part or pkg.main)
    out = []
    for d in root.iter(W + "drawing"):
        blip = d.find(".//a:blip", NS)
        if blip is None:
            continue
        rid = blip.get(q("r:embed"))
        ext = d.find(".//wp:extent", NS)
        size = (int(ext.get("cx")), int(ext.get("cy"))) if ext is not None else None
        out.append((d, rels.get(rid, (None, None, None))[1], size))
    return out


def norm(s: str) -> str:
    return re.sub(r"\s+", " ", s.replace("\u00a0", " ")).strip()
