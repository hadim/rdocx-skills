"""Fault injection on the raw OOXML package: a catalogue of mutations of .docx and .pptx files.

Each mutation takes a valid base package and breaks exactly one rule (``expected="invalid"``), or makes a legal
but unusual change (``expected="valid"``: a negative control). Mutations work on the ZIP entries with
``zipfile`` and ``lxml`` (or on bytes when the fault is below XML), never through a stack's library.

Ground truth is the specification: ECMA-376 Part 1 (the ISO/IEC 29500-4:2016 transitional schemas and the prose
"shall" rules), Part 2 (Open Packaging Conventions), Part 3 (Markup Compatibility), XML 1.0 and the ZIP format.
``basis`` says which: ``zip``, ``opc``, ``xml``, ``schema``, ``prose`` (a "shall" of the Part 1 text) or
``judgement`` (no clause found; the call is explained in ``rule`` and in NOTES.md).

``office`` records what Word or PowerPoint does only when a source says so (quoted, with the source); otherwise
``"unknown"``. ``office_only=True`` marks a file the specification allows but PowerPoint refuses per a cited
source: it counts neither as a detection nor as a false positive.
"""

from __future__ import annotations

import copy
import io
import re
import warnings
import zipfile
from dataclasses import dataclass
from typing import Callable

from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
C = "http://schemas.openxmlformats.org/drawingml/2006/chart"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
PR = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
BENCH = "urn:rdocx-bench:validity"
REL_CUSTOM_XML = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml"

SKILL_PPTX = "built-in pptx SKILL.md"
SKILL_VALIDATOR = "built-in pptx validator, scripts/office/helpers/pptx_slide.py ('the slide-XML schema errors PowerPoint refuses the file over')"


def w(tag):
    return f"{{{W}}}{tag}"


def a(tag):
    return f"{{{A}}}{tag}"


def p(tag):
    return f"{{{P}}}{tag}"


def c(tag):
    return f"{{{C}}}{tag}"


class NotApplicable(Exception):
    """The base has nothing this mutation can act on."""


# --- the package model ---------------------------------------------------------------------------------------
class Pkg:
    """The ZIP entries of a package, in order, with helpers to edit parts and relationships."""

    def __init__(self, data: bytes):
        zf = zipfile.ZipFile(io.BytesIO(data))
        self.entries: list[list] = [[i.filename, zf.read(i.filename)] for i in zf.infolist()]
        self.extra: list[tuple[str, bytes]] = []  # appended as-is (duplicate entries)
        self.compression = zipfile.ZIP_DEFLATED
        self.post: Callable[[bytes], bytes] | None = None

    # entries
    def names(self):
        return [n for n, _ in self.entries]

    def has(self, name):
        return any(n == name for n, _ in self.entries)

    def get(self, name) -> bytes:
        for n, d in self.entries:
            if n == name:
                return d
        raise NotApplicable(f"no part {name}")

    def set(self, name, data: bytes):
        for e in self.entries:
            if e[0] == name:
                e[1] = data
                return
        self.entries.append([name, data])

    def delete(self, name):
        if not self.has(name):
            raise NotApplicable(f"no part {name}")
        self.entries = [e for e in self.entries if e[0] != name]

    def move_last(self, name):
        e = next(e for e in self.entries if e[0] == name)
        self.entries.remove(e)
        self.entries.append(e)

    # XML
    def xml(self, name):
        return etree.fromstring(self.get(name), etree.XMLParser(huge_tree=True))

    def put(self, name, root):
        self.set(name, etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True))

    # relationships
    @staticmethod
    def rels_name(part):
        folder, _, base = part.rpartition("/")
        return f"{folder}/_rels/{base}.rels" if folder else f"_rels/{base}.rels"

    @staticmethod
    def resolve(part, target):
        if target.startswith("/"):
            return target[1:]
        folder = part.rpartition("/")[0]
        segs = folder.split("/") if folder else []
        for s in target.split("/"):
            if s == "..":
                segs.pop()
            elif s and s != ".":
                segs.append(s)
        return "/".join(segs)

    def rels(self, part):
        name = self.rels_name(part) if part else "_rels/.rels"
        if not self.has(name):
            return []
        root = self.xml(name)
        return [(e.get("Id"), e.get("Type"), e.get("Target"), e.get("TargetMode")) for e in root]

    def targets(self, part, type_suffix):
        return [self.resolve(part, t) for _, ty, t, mode in self.rels(part)
                if ty.endswith("/" + type_suffix) and mode != "External"]

    def rel_id(self, part, type_suffix):
        for rid, ty, _, mode in self.rels(part):
            if ty.endswith("/" + type_suffix) and mode != "External":
                return rid
        raise NotApplicable(f"{part} has no {type_suffix} relationship")

    def add_rel(self, part, rid, rtype, target):
        name = self.rels_name(part)
        root = self.xml(name) if self.has(name) else etree.Element(f"{{{PR}}}Relationships", nsmap={None: PR})
        etree.SubElement(root, f"{{{PR}}}Relationship", Id=rid, Type=rtype, Target=target)
        self.put(name, root)

    def main_part(self):
        for _, ty, t, _ in self.rels(""):
            if ty.endswith("/officeDocument"):
                return self.resolve("", t)
        raise NotApplicable("no officeDocument relationship")

    # content types
    def content_types(self):
        return self.xml("[Content_Types].xml")

    def to_bytes(self) -> bytes:
        buf = io.BytesIO()
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")  # duplicate names are intended in one mutation
            with zipfile.ZipFile(buf, "w", self.compression) as zf:
                for name, data in self.entries + [list(e) for e in self.extra]:
                    info = zipfile.ZipInfo(name, date_time=(2026, 10, 1, 0, 0, 0))
                    info.compress_type = zipfile.ZIP_STORED if name.endswith("/") else self.compression
                    info.external_attr = (0o40755 << 16) | 0x10 if name.endswith("/") else 0o644 << 16
                    zf.writestr(info, data)
        out = buf.getvalue()
        return self.post(out) if self.post else out


@dataclass
class Mutation:
    id: str
    fmt: str
    layer: str  # zip | package | xml | schema | reference | semantic
    expected: str  # invalid | valid
    basis: str  # zip | opc | xml | schema | prose | judgement | mce
    rule: str
    description: str
    apply: Callable[[Pkg], None]
    office: str = "unknown"
    office_only: bool = False
    ext: str | None = None  # output extension when not the base's (templates)

    def record(self):
        return {k: v for k, v in self.__dict__.items() if k != "apply"}


CATALOGUE: list[Mutation] = []


def mutation(fmt, layer, expected, basis, rule, description, office="unknown", office_only=False, ext=None):
    def deco(fn):
        CATALOGUE.append(Mutation(f"{fmt}.{fn.__name__}", fmt, layer, expected, basis, rule, description, fn,
                                  office, office_only, ext))
        return fn
    return deco


# --- shared helpers -------------------------------------------------------------------------------------------
def first(root, path, ns=None, what="element"):
    found = root.find(path, ns)
    if found is None:
        raise NotApplicable(f"no {what}")
    return found


def with_ns(root, prefix, uri):
    """A copy of ``root`` whose start tag also declares ``prefix`` (lxml cannot add a declaration in place)."""
    nsmap = dict(root.nsmap)
    nsmap[prefix] = uri
    new = etree.Element(root.tag, nsmap=nsmap)
    for k, v in root.attrib.items():
        new.set(k, v)
    new.text = root.text
    for child in list(root):
        new.append(child)
    return new


def add_ignorable(root, prefix):
    key = f"{{{MC}}}Ignorable"
    if MC not in root.nsmap.values():
        root = with_ns(root, "mc", MC)
    have = root.get(key, "").split()
    root.set(key, " ".join(have + [prefix]))
    return root


def zip_truncate(pkg):
    pkg.post = lambda data: data[: len(data) * 6 // 10]


def zip_duplicate(pkg, part, old: bytes, new: bytes):
    data = pkg.get(part)
    if old not in data:
        raise NotApplicable(f"{old!r} not in {part}")
    pkg.extra.append((part, data.replace(old, new, 1)))


def zip_flip_byte(pkg, part):
    pkg.get(part)

    def flip(data):
        zf = zipfile.ZipFile(io.BytesIO(data))
        info = zf.getinfo(part)
        head = data[info.header_offset: info.header_offset + 30]
        n, m = int.from_bytes(head[26:28], "little"), int.from_bytes(head[28:30], "little")
        start = info.header_offset + 30 + n + m
        k = start + info.compress_size // 2
        b = bytearray(data)
        b[k] ^= 0x5A
        return bytes(b)

    pkg.post = flip


def drop_override(pkg, part):
    ct = pkg.content_types()
    hits = [e for e in ct if e.get("PartName", "").lower() == "/" + part.lower()]
    if not hits:
        raise NotApplicable(f"no Override for {part}")
    for e in hits:
        ct.remove(e)
    pkg.put("[Content_Types].xml", ct)


def drop_default_png(pkg):
    ct = pkg.content_types()
    hits = [e for e in ct if (e.get("Extension") or "").lower() == "png"
            or (e.get("PartName") or "").lower().endswith(".png")]
    if not hits or not any(n.lower().endswith(".png") for n in pkg.names()):
        raise NotApplicable("no png part")
    for e in hits:
        ct.remove(e)
    pkg.put("[Content_Types].xml", ct)


def set_main_content_type(pkg, content_type):
    main = pkg.main_part()
    ct = pkg.content_types()
    e = next((e for e in ct if e.get("PartName", "").lower() == "/" + main.lower()), None)
    if e is None:
        raise NotApplicable("main part has no Override")
    e.set("ContentType", content_type)
    pkg.put("[Content_Types].xml", ct)


def dup_rel_id(pkg, part):
    name = Pkg.rels_name(part)
    root = pkg.xml(name)
    rels = list(root)
    if len(rels) < 2:
        raise NotApplicable("fewer than two relationships")
    rels[1].set("Id", rels[0].get("Id"))
    pkg.put(name, root)


def bad_target_mode(pkg, part):
    name = Pkg.rels_name(part)
    root = pkg.xml(name)
    root[0].set("TargetMode", "Sideways")
    pkg.put(name, root)


def app_company_amp(pkg):
    data = pkg.get("docProps/app.xml")
    raw = b"Harrow & Sons"
    if re.search(rb"<Company>.*?</Company>", data, re.S):
        new = re.sub(rb"<Company>.*?</Company>", b"<Company>" + raw + b"</Company>", data, count=1, flags=re.S)
    elif b"<Company/>" in data:
        new = data.replace(b"<Company/>", b"<Company>" + raw + b"</Company>", 1)
    elif re.search(rb"<((?:\w+:)?Properties)\b[^>]*/>", data):  # an empty, self-closed Properties (docx-js)
        new = re.sub(rb"<((?:\w+:)?Properties)\b([^>]*)/>", rb"<\1\2><Company>" + raw + rb"</Company></\1>",
                     data, count=1)
    else:
        new = re.sub(rb"(</(?:\w+:)?Properties>)", b"<Company>" + raw + rb"</Company>\1", data, count=1)
    if new == data:
        raise NotApplicable("no app.xml Properties")
    pkg.set("docProps/app.xml", new)


def insert_control_char(pkg, part, text_tag: bytes):
    data = pkg.get(part)
    m = re.search(rb"<" + text_tag + rb"(?:\s[^>]*)?>", data)
    if not m:
        raise NotApplicable("no text element")
    pkg.set(part, data[: m.end()] + b"\x01" + data[m.end():])


def truncate_part(pkg, part, cut=120):
    data = pkg.get(part)
    pkg.set(part, data[: len(data) - cut])


def add_custom_xml(pkg):
    main = pkg.main_part()
    item = "customXml/item90.xml"
    pkg.set(item, b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<bench xmlns="urn:rdocx-bench:data"><v>1</v></bench>')
    ct = pkg.content_types()
    if not any((e.get("Extension") or "").lower() == "xml" for e in ct):
        etree.SubElement(ct, f"{{{CT}}}Override", PartName="/" + item, ContentType="application/xml")
        pkg.put("[Content_Types].xml", ct)
    depth = main.count("/")
    pkg.add_rel(main, "rIdBench90", REL_CUSTOM_XML, "../" * depth + item)


def stored(pkg):
    pkg.compression = zipfile.ZIP_STORED


def ct_last(pkg):
    pkg.move_last("[Content_Types].xml")


# --- docx -----------------------------------------------------------------------------------------------------
def doc(pkg):
    main = pkg.main_part()
    return main, pkg.xml(main)


def body(root):
    return first(root, w("body"), what="w:body")


def body_paragraphs(root):
    return [e for e in body(root) if e.tag == w("p")]


def first_run_with_text(root):
    for r in root.iter(w("r")):
        if r.find(w("t")) is not None and r.getparent().tag == w("p"):
            return r
    raise NotApplicable("no run with w:t")


PPR_AFTER_JC = {"textDirection", "textAlignment", "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr",
                "sectPr", "pPrChange"}
PPR_TAIL = {"rPr", "sectPr", "pPrChange"}


def insert_ppr_child(ppr, el, after_set):
    for i, child in enumerate(ppr):
        if etree.QName(child).localname in after_set:
            ppr.insert(i, el)
            return
    ppr.append(el)


def ppr_of(par):
    ppr = par.find(w("pPr"))
    if ppr is None:
        ppr = etree.Element(w("pPr"))
        par.insert(0, ppr)
    return ppr


D = "docx"


@mutation(D, "zip", "invalid", "zip", "The ZIP archive must be complete (end of central directory record): ECMA-376 Part 2, physical package mapping to ZIP.",
          "Archive truncated to 60 % of its bytes (no central directory).")
def zip_truncated(pkg):
    zip_truncate(pkg)


@mutation(D, "zip", "invalid", "opc", "ECMA-376 Part 2, part names: a package shall not contain equivalent part names.",
          "Two ZIP items named word/document.xml, the second with different text.")
def zip_duplicate_entry(pkg):
    main = pkg.main_part()
    data = pkg.get(main)
    m = re.search(rb"<w:t(?:\s[^>]*)?>([^<]{3,})</w:t>", data)
    if not m:
        raise NotApplicable("no text")
    zip_duplicate(pkg, main, m.group(0), m.group(0).replace(m.group(1), b"DUPLICATE"))


@mutation(D, "zip", "invalid", "zip", "ZIP: the stored CRC-32 must match the data (APPNOTE 4.4.7), the integrity check of the physical package.",
          "One byte of the compressed data of word/document.xml flipped (CRC mismatch).")
def zip_bad_crc(pkg):
    zip_flip_byte(pkg, pkg.main_part())


@mutation(D, "package", "invalid", "opc", "ECMA-376 Part 2, content types: the package shall contain a Content Types stream, [Content_Types].xml.",
          "[Content_Types].xml removed.")
def opc_no_content_types(pkg):
    pkg.delete("[Content_Types].xml")


@mutation(D, "package", "invalid", "opc", "ECMA-376 Part 1 11.3.10: the main document part has content type ...wordprocessingml.document.main+xml; without its Override it falls back to the xml Default (application/xml) or has none (Part 2: every part shall have a content type).",
          "Override for /word/document.xml removed.")
def opc_no_override_main(pkg):
    drop_override(pkg, pkg.main_part())


@mutation(D, "package", "invalid", "opc", "ECMA-376 Part 2, content types: every part shall have a content type (an Override or a Default for its extension).",
          "Default for the png extension removed: the picture part has no content type.")
def opc_no_default_png(pkg):
    drop_default_png(pkg)


@mutation(D, "package", "invalid", "opc", "ECMA-376 Part 2 package relationships / Part 1 11.3.10: the package-relationship part names the main document part; without it there is no officeDocument.",
          "_rels/.rels removed.")
def opc_no_root_rels(pkg):
    pkg.delete("_rels/.rels")


@mutation(D, "package", "invalid", "judgement", "The image relationship of the main part targets a part that is not in the package (an internal relationship whose target is missing). ECMA-376 Part 1 15.2.14 image part: the relationship points to an Image part, absent here.",
          "The picture's media part deleted; its relationship and r:embed kept.")
def opc_media_missing(pkg):
    main = pkg.main_part()
    targets = pkg.targets(main, "image")
    if not targets:
        raise NotApplicable("no image relationship")
    pkg.delete(targets[0])


@mutation(D, "package", "invalid", "opc", "ECMA-376 Part 2 relationships schema: Relationship/@Id is xsd:ID, unique within the Relationships part.",
          "Two relationships of word/_rels/document.xml.rels share one Id.")
def opc_dup_rel_id(pkg):
    dup_rel_id(pkg, pkg.main_part())


@mutation(D, "package", "invalid", "opc", "ECMA-376 Part 2 relationships schema: TargetMode is ST_TargetMode, either Internal or External.",
          'TargetMode="Sideways" on a relationship of the main part.')
def opc_bad_target_mode(pkg):
    bad_target_mode(pkg, pkg.main_part())


@mutation(D, "xml", "invalid", "xml", "XML 1.0 well-formedness: every start tag has a matching end tag (document element unclosed).",
          "Last 120 bytes of word/document.xml cut (end tags missing).")
def xml_document_truncated(pkg):
    truncate_part(pkg, pkg.main_part())


@mutation(D, "xml", "invalid", "xml", "XML 1.0 well-formedness: '&' must start an entity or character reference.",
          "Unescaped '&' in docProps/app.xml (<Company>Harrow & Sons</Company>).")
def xml_app_amp(pkg):
    app_company_amp(pkg)


@mutation(D, "xml", "invalid", "xml", "XML 1.0 well-formedness: end tag must match the start tag.",
          "word/comments.xml closes </w:comments> as </w:comment>.")
def xml_comments_mismatched(pkg):
    main = pkg.main_part()
    targets = pkg.targets(main, "comments")
    if not targets:
        raise NotApplicable("no comments part")
    data = pkg.get(targets[0])
    if b"</w:comments>" not in data:
        raise NotApplicable("no </w:comments>")
    pkg.set(targets[0], data.replace(b"</w:comments>", b"</w:comment>"))


@mutation(D, "xml", "invalid", "xml", "XML 1.0 Char production: U+0001 is not an allowed character.",
          "A raw 0x01 byte inside the first w:t.")
def xml_control_char(pkg):
    insert_control_char(pkg, pkg.main_part(), b"w:t")


@mutation(D, "schema", "invalid", "schema", "CT_P is a sequence (pPr?, EG_PContent*): w:pPr must be the first child.",
          "w:pPr moved after the first w:r of the heading paragraph.")
def schema_ppr_after_r(pkg):
    main, root = doc(pkg)
    for par in root.iter(w("p")):
        ppr, run = par.find(w("pPr")), par.find(w("r"))
        if ppr is not None and run is not None:
            par.remove(ppr)
            run.addnext(ppr)
            pkg.put(main, root)
            return
    raise NotApplicable("no paragraph with pPr and a run")


@mutation(D, "schema", "invalid", "schema", "CT_ParaRPr is a sequence (ins?, del?, moveFrom?, moveTo?, EG_RPrBase*, rPrChange?): the paragraph-mark w:del must precede w:b.",
          "Paragraph-mark rPr <w:b/><w:del .../> (deleted paragraph mark after a formatting child).")
def schema_pmark_del_misordered(pkg):
    main, root = doc(pkg)
    par = body_paragraphs(root)[0]
    ppr = ppr_of(par)
    rpr = etree.Element(w("rPr"))
    etree.SubElement(rpr, w("b"))
    etree.SubElement(rpr, w("del"), {w("id"): "9101", w("author"): "Bench", w("date"): "2026-10-01T00:00:00Z"})
    if ppr.find(w("rPr")) is not None:
        ppr.remove(ppr.find(w("rPr")))
    insert_ppr_child(ppr, rpr, {"sectPr", "pPrChange"})
    pkg.put(main, root)


@mutation(D, "schema", "invalid", "schema", "wml.xsd: no element w:bogus exists in the WordprocessingML namespace, and the namespace is not ignorable (Part 3).",
          "<w:bogus/> inserted in a paragraph.")
def schema_unknown_element(pkg):
    main, root = doc(pkg)
    par = body_paragraphs(root)[0]
    run = first(par, w("r"), what="run")
    run.addprevious(etree.Element(w("bogus")))
    pkg.put(main, root)


@mutation(D, "schema", "invalid", "schema", 'ST_Jc enumeration has no value "middle".',
          '<w:jc w:val="middle"/> on the first paragraph.')
def schema_bad_enum(pkg):
    main, root = doc(pkg)
    par = body_paragraphs(root)[0]
    ppr = ppr_of(par)
    jc = ppr.find(w("jc"))
    if jc is None:
        jc = etree.Element(w("jc"))
        insert_ppr_child(ppr, jc, PPR_AFTER_JC)
    jc.set(w("val"), "middle")
    pkg.put(main, root)


@mutation(D, "schema", "invalid", "schema", "CT_R has element-only content (EG_RPr?, EG_RunInnerContent*): character data directly in w:r is not allowed.",
          "Text 'stray' placed directly in a w:r, outside w:t.")
def schema_text_outside_t(pkg):
    main, root = doc(pkg)
    run = first_run_with_text(root)
    run.text = "stray"
    pkg.put(main, root)


@mutation(D, "schema", "invalid", "prose", "ECMA-376 Part 1 17.4.66 tc: a table cell shall contain at least one block-level element (else the document is considered corrupt). The XSD cannot enforce it: EG_BlockLevelElts+ is satisfied by an empty group.",
          "Every paragraph of the first table cell removed (tcPr kept).")
def schema_tc_no_p(pkg):
    main, root = doc(pkg)
    tc = first(root, f".//{w('tc')}", what="table cell")
    for child in list(tc):
        if child.tag != w("tcPr"):
            tc.remove(child)
    pkg.put(main, root)


@mutation(D, "schema", "invalid", "schema", "CT_Body is a sequence (EG_BlockLevelElts*, sectPr?): the body's w:sectPr must be last.",
          "Body w:sectPr moved to be the first child of w:body.")
def schema_sectpr_not_last(pkg):
    main, root = doc(pkg)
    b = body(root)
    sect = b[-1]
    if sect.tag != w("sectPr"):
        raise NotApplicable("body has no final sectPr")
    b.remove(sect)
    b.insert(0, sect)
    pkg.put(main, root)


@mutation(D, "schema", "invalid", "schema", "w:color/@w:val is ST_HexColor (auto or ST_HexColorRGB = hexBinary of length 3): '#FF0000' is neither.",
          '<w:color w:val="#FF0000"/> on a run.')
def schema_color_hash(pkg):
    main, root = doc(pkg)
    run = first_run_with_text(root)
    rpr = run.find(w("rPr"))
    if rpr is None:
        rpr = etree.Element(w("rPr"))
        run.insert(0, rpr)
    rpr.insert(0, etree.Element(w("color"), {w("val"): "#FF0000"}))
    pkg.put(main, root)


@mutation(D, "schema", "invalid", "schema", "wp:extent is CT_PositiveSize2D: cx is ST_PositiveCoordinate (>= 0).",
          "Picture wp:extent cx made negative.")
def schema_negative_extent(pkg):
    main, root = doc(pkg)
    ext = first(root, f".//{{{WP}}}extent", what="wp:extent")
    ext.set("cx", "-" + ext.get("cx"))
    pkg.put(main, root)


@mutation(D, "schema", "invalid", "schema", "CT_TrackChange: w:id is required (use=required).",
          "w:id removed from the tracked insertion (w:ins).")
def schema_ins_no_id(pkg):
    main, root = doc(pkg)
    ins = first(root, f".//{w('ins')}", what="w:ins")
    del ins.attrib[w("id")]
    pkg.put(main, root)


@mutation(D, "reference", "invalid", "prose", "ECMA-376 Part 1 20.1.8.13 blip: r:embed specifies the relationship of the part that holds the picture; rId999 is no relationship of the part.",
          'Picture a:blip r:embed="rId999" (no such relationship).')
def ref_blip_missing_rel(pkg):
    main, root = doc(pkg)
    blip = first(root, f".//{a('blip')}", what="a:blip")
    blip.set(f"{{{R}}}embed", "rId999")
    pkg.put(main, root)


@mutation(D, "reference", "invalid", "prose", "ECMA-376 Part 1 17.10.5 headerReference: r:id identifies the relationship to the header part; rId999 does not exist.",
          '<w:headerReference w:type="default" r:id="rId999"/> added to the body sectPr.')
def ref_header_missing_rel(pkg):
    main, root = doc(pkg)
    b = body(root)
    sect = b[-1]
    if sect.tag != w("sectPr"):
        raise NotApplicable("no body sectPr")
    sect.insert(0, etree.Element(w("headerReference"), {w("type"): "default", f"{{{R}}}id": "rId999"}))
    pkg.put(main, root)


@mutation(D, "reference", "invalid", "prose", "ECMA-376 Part 1 17.13.4.5 commentReference: w:id links the reference to a comment of the Comments part; no comment 999 exists.",
          'w:commentReference w:id="999" (no such comment).')
def ref_comment_ref_missing(pkg):
    main, root = doc(pkg)
    ref = first(root, f".//{w('commentReference')}", what="commentReference")
    ref.set(w("id"), "999")
    pkg.put(main, root)


@mutation(D, "reference", "invalid", "prose", "ECMA-376 Part 1 17.11.14 footnoteReference: the id shall match a footnote of the footnotes part, else the document is non-conformant.",
          '<w:footnoteReference w:id="999"/> run added (no such footnote).')
def ref_footnote_ref_missing(pkg):
    main, root = doc(pkg)
    par = body_paragraphs(root)[1]
    r = etree.SubElement(par, w("r"))
    etree.SubElement(r, w("footnoteReference"), {w("id"): "999"})
    pkg.put(main, root)


@mutation(D, "reference", "invalid", "prose", "ECMA-376 Part 1 17.13.6.2 bookmarkStart: w:id is a unique identifier of the annotation; two bookmarks share id 77.",
          "Two bookmarks (names a and b) with the same w:id 77.")
def ref_dup_bookmark_id(pkg):
    main, root = doc(pkg)
    pars = body_paragraphs(root)
    for par, name in ((pars[0], "bench_a"), (pars[1], "bench_b")):
        ppr = par.find(w("pPr"))
        at = 1 if ppr is not None else 0
        par.insert(at, etree.Element(w("bookmarkStart"), {w("id"): "77", w("name"): name}))
        par.append(etree.Element(w("bookmarkEnd"), {w("id"): "77"}))
    pkg.put(main, root)


@mutation(D, "reference", "invalid", "prose", "ECMA-376 Part 1 17.13.5 revisions: w:id of an annotation is a unique identifier; the deletion and the insertion share one id.",
          "The w:del takes the w:id of the w:ins.")
def ref_dup_revision_id(pkg):
    main, root = doc(pkg)
    ins = first(root, f".//{w('ins')}", what="w:ins")
    dele = first(root, f".//{w('del')}", what="w:del")
    dele.set(w("id"), ins.get(w("id")))
    pkg.put(main, root)


@mutation(D, "semantic", "invalid", "prose", "ECMA-376 Part 1 17.3.3.7 delText: delText shall be used for all text runs inside a deleted region (w:del); w:t there is the inserted-text element.",
          "Inside w:del, w:delText renamed to w:t.")
def sem_del_with_t(pkg):
    main, root = doc(pkg)
    dt = first(root, f".//{w('del')}//{w('delText')}", what="w:delText")
    dt.tag = w("t")
    pkg.put(main, root)


# docx negative controls
@mutation(D, "zip", "valid", "zip", "Identity: entries read and written back (deflate, same order).",
          "Repack without change (checks the harness).")
def ctl_repack_identity(pkg):
    pass


@mutation(D, "zip", "valid", "zip", "ECMA-376 Part 2 allows stored (uncompressed) ZIP items.",
          "Every item stored without compression.")
def ctl_zip_stored(pkg):
    stored(pkg)


@mutation(D, "zip", "valid", "opc", "ECMA-376 Part 2 does not fix the order of ZIP items.",
          "[Content_Types].xml written as the last ZIP item.")
def ctl_content_types_last(pkg):
    ct_last(pkg)


@mutation(D, "schema", "valid", "mce", "ECMA-376 Part 3 (Markup Compatibility): mc:AlternateContent with a Choice and a Fallback is legal anywhere.",
          'First run wrapped in mc:AlternateContent (Choice Requires="w14", Fallback with the same run).')
def ctl_alternate_content(pkg):
    main, root = doc(pkg)
    run = first_run_with_text(root)
    alt = etree.Element(f"{{{MC}}}AlternateContent", nsmap={"mc": MC})
    choice = etree.SubElement(alt, f"{{{MC}}}Choice", Requires="w14",
                              nsmap={"w14": "http://schemas.microsoft.com/office/word/2010/wordml"})
    fallback = etree.SubElement(alt, f"{{{MC}}}Fallback")
    run.addprevious(alt)
    choice.append(copy.deepcopy(run))
    fallback.append(run)
    pkg.put(main, root)


@mutation(D, "schema", "valid", "mce", "ECMA-376 Part 3: elements and attributes of a namespace listed in mc:Ignorable are ignored by consumers that do not understand it.",
          "Unknown element and attribute of an ignorable namespace in a paragraph.")
def ctl_ignorable_unknown(pkg):
    main, root = doc(pkg)
    root = add_ignorable(with_ns(root, "bench", BENCH), "bench")
    par = body_paragraphs(root)[0]
    par.set(f"{{{BENCH}}}mark", "1")
    first(par, w("r"), what="run").addprevious(etree.Element(f"{{{BENCH}}}note"))
    pkg.put(main, root)


@mutation(D, "package", "valid", "opc", "ECMA-376 Part 1 15.2.4: a Custom XML Data Storage part may be related from the main document.",
          "Custom XML part added and related from the main part.")
def ctl_custom_xml_part(pkg):
    add_custom_xml(pkg)


@mutation(D, "reference", "valid", "schema", "A style may be defined and never used.",
          "Paragraph style BenchUnused defined in styles.xml, used nowhere.")
def ctl_unused_style(pkg):
    main = pkg.main_part()
    st = pkg.targets(main, "styles")
    if not st:
        raise NotApplicable("no styles part")
    root = pkg.xml(st[0])
    style = etree.SubElement(root, w("style"), {w("type"): "paragraph", w("customStyle"): "1", w("styleId"): "BenchUnused"})
    etree.SubElement(style, w("name"), {w("val"): "Bench Unused"})
    etree.SubElement(style, w("qFormat"))
    pkg.put(st[0], root)


@mutation(D, "schema", "valid", "schema", "No length limit on a paragraph or a w:t.",
          "A paragraph of 200,000 characters added before the body sectPr.")
def ctl_long_paragraph(pkg):
    main, root = doc(pkg)
    b = body(root)
    par = etree.Element(w("p"))
    t = etree.SubElement(etree.SubElement(par, w("r")), w("t"))
    t.text = ("lorem ipsum dolor " * 11112)[:200000]
    (b[-1].addprevious if b[-1].tag == w("sectPr") else b.append)(par)
    pkg.put(main, root)


@mutation(D, "schema", "valid", "schema", "CT_Tc needs one block-level element: an empty <w:p/> is enough.",
          "Last table cell holds only an empty <w:p/>.")
def ctl_empty_cell_paragraph(pkg):
    main, root = doc(pkg)
    tcs = root.findall(f".//{w('tc')}")
    if not tcs:
        raise NotApplicable("no table")
    tc = tcs[-1]
    for child in list(tc):
        if child.tag != w("tcPr"):
            tc.remove(child)
    etree.SubElement(tc, w("p"))
    pkg.put(main, root)


@mutation(D, "schema", "valid", "schema", "ISO/IEC 29500-4:2016 wml.xsd: EG_RPrBase is an xsd:choice repeated, so run properties may come in any order (the 1st edition's sequence order is not required in transitional).",
          "A run's rPr written as <w:i/><w:b/> (reverse of the canonical order).")
def ctl_rpr_reordered(pkg):
    main, root = doc(pkg)
    run = first_run_with_text(root)
    rpr = run.find(w("rPr"))
    if rpr is not None:
        run.remove(rpr)
    rpr = etree.Element(w("rPr"))
    etree.SubElement(rpr, w("i"))
    etree.SubElement(rpr, w("b"))
    run.insert(0, rpr)
    pkg.put(main, root)


@mutation(D, "package", "valid", "opc", "ECMA-376 Part 1 11.3.10: a template's main part has content type ...wordprocessingml.template.main+xml (.dotx).",
          "Main part content type set to the template type, file named .dotx.", ext=".dotx")
def ctl_dotx_template(pkg):
    set_main_content_type(pkg, "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml")


@mutation(D, "reference", "valid", "judgement", "w:pStyle/@w:val is ST_String; no schema or OPC rule requires the style to be defined, and consumers fall back to the default paragraph style (judgement: valid).",
          'A paragraph with <w:pStyle w:val="NoSuchStyle"/> (style not defined).')
def ctl_style_undefined(pkg):
    main, root = doc(pkg)
    par = body_paragraphs(root)[2]
    ppr = ppr_of(par)
    ps = ppr.find(w("pStyle"))
    if ps is None:
        ps = etree.Element(w("pStyle"))
        ppr.insert(0, ps)
    ps.set(w("val"), "NoSuchStyle")
    pkg.put(main, root)


# --- pptx -----------------------------------------------------------------------------------------------------
X = "pptx"


def presentation(pkg):
    main = pkg.main_part()
    return main, pkg.xml(main)


def slides(pkg):
    main, root = presentation(pkg)
    rels = {rid: pkg.resolve(main, t) for rid, _, t, _ in pkg.rels(main)}
    lst = root.find(p("sldIdLst"))
    if lst is None:
        raise NotApplicable("no sldIdLst")
    return [rels[e.get(f"{{{R}}}id")] for e in lst]


def slide_with(pkg, path, what):
    for name in slides(pkg):
        root = pkg.xml(name)
        el = root.find(path)
        if el is not None:
            return name, root, el
    raise NotApplicable(f"no slide with {what}")


def chart_part(pkg):
    for name in slides(pkg):
        for t in pkg.targets(name, "chart"):
            return t
    raise NotApplicable("no chart")


SPPR_ORDER = ["xfrm", "custGeom", "prstGeom", "noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill",
              "ln", "effectLst", "effectDag", "scene3d", "sp3d", "extLst"]


def sppr_insert(sppr, el):
    rank = SPPR_ORDER.index(etree.QName(el).localname)
    for i, child in enumerate(sppr):
        if SPPR_ORDER.index(etree.QName(child).localname) > rank:
            sppr.insert(i, el)
            return
    sppr.append(el)


FILLS = {"noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill"}


def shape_color(pkg, value):
    """Give the first slide's first p:sp a solid fill of colour ``value`` (whatever fill it had is replaced)."""
    name, root, sp = slide_with(pkg, f".//{p('sp')}", "a shape")
    sppr = first(sp, p("spPr"), what="spPr")
    for child in list(sppr):
        if etree.QName(child).localname in FILLS:
            sppr.remove(child)
    fill = etree.Element(a("solidFill"))
    etree.SubElement(fill, a("srgbClr"), val=value)
    sppr_insert(sppr, fill)
    pkg.put(name, root)


@mutation(X, "zip", "invalid", "zip", "The ZIP archive must be complete (end of central directory record): ECMA-376 Part 2, physical package mapping to ZIP.",
          "Archive truncated to 60 % of its bytes (no central directory).")
def zip_truncated(pkg):  # noqa: F811 (one function per format, registered by the decorator)
    zip_truncate(pkg)


@mutation(X, "zip", "invalid", "opc", "ECMA-376 Part 2, part names: a package shall not contain equivalent part names.",
          "Two ZIP items named after the first slide, the second with different text.")
def zip_duplicate_entry(pkg):  # noqa: F811
    name = slides(pkg)[0]
    data = pkg.get(name)
    m = re.search(rb"<a:t>([^<]{3,})</a:t>", data)
    if not m:
        raise NotApplicable("no text")
    zip_duplicate(pkg, name, m.group(0), m.group(0).replace(m.group(1), b"DUPLICATE"))


@mutation(X, "zip", "invalid", "zip", "ZIP: the stored CRC-32 must match the data (APPNOTE 4.4.7).",
          "One byte of the compressed data of the first slide flipped (CRC mismatch).")
def zip_bad_crc(pkg):  # noqa: F811
    zip_flip_byte(pkg, slides(pkg)[0])


@mutation(X, "package", "invalid", "opc", "ECMA-376 Part 2, content types: the package shall contain a Content Types stream, [Content_Types].xml.",
          "[Content_Types].xml removed.")
def opc_no_content_types(pkg):  # noqa: F811
    pkg.delete("[Content_Types].xml")


@mutation(X, "package", "invalid", "opc", "ECMA-376 Part 1 13.3.8: a slide part has content type ...presentationml.slide+xml; without its Override it falls back to the xml Default (application/xml) or has none (Part 2).",
          "Override for the first slide removed.")
def opc_no_override_slide(pkg):
    drop_override(pkg, slides(pkg)[0])


@mutation(X, "package", "invalid", "opc", "ECMA-376 Part 2, content types: every part shall have a content type (an Override or a Default for its extension).",
          "Default for the png extension removed: the picture part has no content type.")
def opc_no_default_png(pkg):  # noqa: F811
    drop_default_png(pkg)


@mutation(X, "package", "invalid", "opc", "ECMA-376 Part 2 package relationships / Part 1 13.3.6: the package-relationship part names the presentation part.",
          "_rels/.rels removed.")
def opc_no_root_rels(pkg):  # noqa: F811
    pkg.delete("_rels/.rels")


@mutation(X, "package", "invalid", "judgement", "The image relationship of a slide targets a part that is not in the package (an internal relationship whose target is missing). ECMA-376 Part 1 15.2.14 image part.",
          "The picture's media part deleted; its relationship and r:embed kept.")
def opc_media_missing(pkg):  # noqa: F811
    for name in slides(pkg):
        targets = pkg.targets(name, "image")
        if targets:
            pkg.delete(targets[0])
            return
    raise NotApplicable("no image")


@mutation(X, "package", "invalid", "opc", "ECMA-376 Part 2 relationships schema: Relationship/@Id is xsd:ID, unique within the Relationships part.",
          "Two relationships of ppt/_rels/presentation.xml.rels share one Id.")
def opc_dup_rel_id(pkg):  # noqa: F811
    dup_rel_id(pkg, pkg.main_part())


@mutation(X, "package", "invalid", "opc", "ECMA-376 Part 2 relationships schema: TargetMode is ST_TargetMode, either Internal or External.",
          'TargetMode="Sideways" on a relationship of the presentation part.')
def opc_bad_target_mode(pkg):  # noqa: F811
    bad_target_mode(pkg, pkg.main_part())


@mutation(X, "xml", "invalid", "xml", "XML 1.0 well-formedness: every start tag has a matching end tag.",
          "Last 120 bytes of the first slide cut (end tags missing).")
def xml_slide_truncated(pkg):
    truncate_part(pkg, slides(pkg)[0])


@mutation(X, "xml", "invalid", "xml", "XML 1.0 well-formedness: '&' must start an entity or character reference.",
          "Unescaped '&' in docProps/app.xml (<Company>Harrow & Sons</Company>), the pptxgenjs pres.company fault.",
          office=f"{SKILL_PPTX}: \"pres.company ... written into the XML unescaped -- an & ... leaves docProps/app.xml ... unparseable, and validate.py fails\" (no statement on PowerPoint itself)")
def xml_app_amp(pkg):  # noqa: F811
    app_company_amp(pkg)


@mutation(X, "xml", "invalid", "xml", "XML 1.0 well-formedness: '&' must start an entity or character reference (here in an attribute value).",
          "Unescaped '&' in the p:cSld name of the first slide's layout, the pptxgenjs defineSlideMaster title fault.",
          office=f"{SKILL_PPTX}: \"a defineSlideMaster() title [is] written into the XML unescaped -- an & ... leaves ... the layout unparseable, and validate.py fails\" (no statement on PowerPoint itself)")
def xml_layout_name_amp(pkg):
    layout = pkg.targets(slides(pkg)[0], "slideLayout")
    if not layout:
        raise NotApplicable("no layout")
    data = pkg.get(layout[0])
    new, n = re.subn(rb'<p:cSld name="[^"]*"', b'<p:cSld name="Harrow & Sons"', data, count=1)
    if not n:
        new, n = re.subn(rb"<p:cSld\b", b'<p:cSld name="Harrow & Sons"', data, count=1)
    if not n:
        raise NotApplicable("no cSld")
    pkg.set(layout[0], new)


@mutation(X, "xml", "invalid", "xml", "XML 1.0 Char production: U+0001 is not an allowed character.",
          "A raw 0x01 byte inside the first a:t of the first slide.")
def xml_control_char(pkg):  # noqa: F811
    insert_control_char(pkg, slides(pkg)[0], b"a:t")


@mutation(X, "schema", "invalid", "schema", "a:srgbClr/@val is ST_HexColorRGB, hexBinary of length 3 (six hex digits): eight digits break it.",
          'Shape fill <a:srgbClr val="00000020"/> (alpha baked into the hex).',
          office=f"{SKILL_PPTX}: \"alpha baked into the hex (\\\"00000020\\\") corrupt[s] the file\"; {SKILL_VALIDATOR}: 'a colour that is not six hex digits'")
def schema_hex8_colour(pkg):
    shape_color(pkg, "00000020")


@mutation(X, "schema", "invalid", "schema", "a:srgbClr/@val is ST_HexColorRGB (hexBinary): '#' is not a hex digit.",
          'Shape fill <a:srgbClr val="#FF0000"/>.',
          office=f"{SKILL_PPTX}: \"\\\"#FF0000\\\" ... corrupt[s] the file\"")
def schema_hash_colour(pkg):
    shape_color(pkg, "#FF0000")


@mutation(X, "schema", "invalid", "schema", "a:outerShdw/@dist is ST_PositiveCoordinate (>= 0).",
          "Outer shadow with dist=-38100 on a shape (the pptxgenjs negative shadow offset).",
          office=f"{SKILL_PPTX}: \"Shadow offset must be >= 0 -- a negative offset corrupts the file\"")
def schema_negative_shadow(pkg):
    name, root, sp = slide_with(pkg, f".//{p('sp')}", "a shape")
    sppr = first(sp, p("spPr"), what="spPr")
    for child in list(sppr):
        if etree.QName(child).localname in ("effectLst", "effectDag"):
            sppr.remove(child)
    eff = etree.Element(a("effectLst"))
    sh = etree.SubElement(eff, a("outerShdw"), blurRad="38100", dist="-38100", dir="2700000", algn="tl", rotWithShape="0")
    etree.SubElement(etree.SubElement(sh, a("srgbClr"), val="000000"), a("alpha"), val="40000")
    sppr_insert(sppr, eff)
    pkg.put(name, root)


@mutation(X, "schema", "invalid", "schema", "CT_Presentation is a sequence: sldIdLst precedes sldSz.",
          "p:sldSz moved before p:sldIdLst in presentation.xml.",
          office=f"{SKILL_PPTX}: \"Never reorder the children of <p:presentation> ... move the element and the same deck becomes unopenable\"")
def schema_presentation_reorder(pkg):
    main, root = presentation(pkg)
    lst, sz = root.find(p("sldIdLst")), root.find(p("sldSz"))
    if lst is None or sz is None:
        raise NotApplicable("no sldIdLst or sldSz")
    root.remove(sz)
    lst.addprevious(sz)
    pkg.put(main, root)


@mutation(X, "schema", "invalid", "schema", "CT_Presentation is a sequence (sldMasterIdLst?, notesMasterIdLst?, handoutMasterIdLst?, sldIdLst?, ...): notesMasterIdLst precedes sldIdLst.",
          "p:notesMasterIdLst moved after p:sldIdLst (the order pptxgenjs writes; not applicable to the pptxgenjs base, which has it already).",
          office=f"{SKILL_PPTX}: \"pptxgenjs writes <p:notesMasterIdLst> right after <p:sldIdLst> ... PowerPoint reads that happily\"")
def schema_notes_master_after_slides(pkg):
    main, root = presentation(pkg)
    lst, notes = root.find(p("sldIdLst")), root.find(p("notesMasterIdLst"))
    if lst is None or notes is None:
        raise NotApplicable("no sldIdLst or notesMasterIdLst")
    kids = list(root)
    if kids.index(notes) > kids.index(lst):
        raise NotApplicable("already after sldIdLst")
    root.remove(notes)
    lst.addnext(notes)
    pkg.put(main, root)


@mutation(X, "schema", "invalid", "schema", "pml.xsd: no element p:bogus exists in the PresentationML namespace, which is not ignorable.",
          "<p:bogus/> inserted in the first slide's spTree.")
def schema_unknown_element(pkg):  # noqa: F811
    name = slides(pkg)[0]
    root = pkg.xml(name)
    tree = first(root, f".//{p('spTree')}", what="spTree")
    tree.append(etree.Element(p("bogus")))
    pkg.put(name, root)


@mutation(X, "schema", "invalid", "schema", 'a:pPr/@algn is ST_TextAlignType, which has no value "middle".',
          'algn="middle" on the first paragraph of the first slide.')
def schema_bad_enum(pkg):  # noqa: F811
    name = slides(pkg)[0]
    root = pkg.xml(name)
    par = first(root, f".//{a('p')}", what="a:p")
    ppr = par.find(a("pPr"))
    if ppr is None:
        ppr = etree.Element(a("pPr"))
        par.insert(0, ppr)
    ppr.set("algn", "middle")
    pkg.put(name, root)


@mutation(X, "schema", "invalid", "schema", "a:ext in a:xfrm / p:xfrm is CT_PositiveSize2D: cx is ST_PositiveCoordinate (>= 0).",
          "First a:ext of a slide shape given a negative cx.")
def schema_negative_extent(pkg):  # noqa: F811
    for name in slides(pkg):
        root = pkg.xml(name)
        ext = next((e for e in root.iter(a("ext")) if e.get("cx", "0").lstrip("-0")), None)
        if ext is not None:
            ext.set("cx", "-" + ext.get("cx"))
            pkg.put(name, root)
            return
    raise NotApplicable("no non-zero xfrm extent")


@mutation(X, "schema", "invalid", "schema", "CT_TextBody is a sequence (bodyPr, lstStyle?, p+): an empty p:txBody misses bodyPr and p.",
          "Every child of the first p:txBody of the first slide removed (<p:txBody/>).",
          office=f"{SKILL_VALIDATOR}: 'a <p:txBody> with no children'")
def schema_txbody_empty(pkg):
    name = slides(pkg)[0]
    root = pkg.xml(name)
    tb = first(root, f".//{p('txBody')}", what="txBody")
    for child in list(tb):
        tb.remove(child)
    pkg.put(name, root)


@mutation(X, "schema", "invalid", "schema", "p:sldId/@id is ST_SlideId: unsignedInt from 256 to 2147483647.",
          'First p:sldId id="100".')
def schema_sldid_below_256(pkg):
    main, root = presentation(pkg)
    first(root, f"{p('sldIdLst')}/{p('sldId')}", what="sldId").set("id", "100")
    pkg.put(main, root)


@mutation(X, "schema", "invalid", "schema", "a:CT_NonVisualDrawingProps: name is required (use=required).",
          "name attribute removed from the first p:cNvPr of the first slide.")
def schema_cnvpr_no_name(pkg):
    name = slides(pkg)[0]
    root = pkg.xml(name)
    cnv = first(root, f".//{p('cNvPr')}", what="cNvPr")
    del cnv.attrib["name"]
    pkg.put(name, root)


@mutation(X, "reference", "invalid", "prose", "ECMA-376 Part 1 19.2.1.33 sldId: id is unique throughout the presentation.",
          "Second p:sldId takes the id of the first.")
def ref_sldid_duplicate(pkg):
    main, root = presentation(pkg)
    ids = root.findall(f"{p('sldIdLst')}/{p('sldId')}")
    if len(ids) < 2:
        raise NotApplicable("fewer than two slides")
    ids[1].set("id", ids[0].get("id"))
    pkg.put(main, root)


@mutation(X, "reference", "invalid", "prose", "ECMA-376 Part 1 19.2.1.33 sldId: r:id names the relationship to the slide part; rId999 does not exist.",
          'First p:sldId r:id="rId999".')
def ref_sldid_missing_rel(pkg):
    main, root = presentation(pkg)
    first(root, f"{p('sldIdLst')}/{p('sldId')}", what="sldId").set(f"{{{R}}}id", "rId999")
    pkg.put(main, root)


@mutation(X, "reference", "invalid", "prose", "ECMA-376 Part 1 13.3.8 slide part: a slide part shall have an implicit relationship to exactly one slide layout part.",
          "slideLayout relationship removed from the first slide's rels.")
def ref_slide_no_layout(pkg):
    name = slides(pkg)[0]
    rels = Pkg.rels_name(name)
    root = pkg.xml(rels)
    hits = [e for e in root if e.get("Type", "").endswith("/slideLayout")]
    if not hits:
        raise NotApplicable("no layout rel")
    for e in hits:
        root.remove(e)
    pkg.put(rels, root)


@mutation(X, "reference", "invalid", "prose", "ECMA-376 Part 1 20.1.8.13 blip: r:embed names the relationship of the picture part; rId999 does not exist.",
          'Picture a:blip r:embed="rId999".')
def ref_blip_missing_rel(pkg):  # noqa: F811
    name, root, blip = slide_with(pkg, f".//{p('pic')}//{a('blip')}", "a picture")
    blip.set(f"{{{R}}}embed", "rId999")
    pkg.put(name, root)


@mutation(X, "reference", "invalid", "prose", "ECMA-376 Part 1 21.2.2.27 chart (graphic frame reference): r:id names the relationship to the chart part; rId999 does not exist.",
          'Graphic frame c:chart r:id="rId999".')
def ref_chart_missing_rel(pkg):
    name, root, ch = slide_with(pkg, f".//{c('chart')}", "a chart")
    ch.set(f"{{{R}}}id", "rId999")
    pkg.put(name, root)


@mutation(X, "reference", "invalid", "judgement", "ECMA-376 Part 1 21.2.2.9 axId 'specifies the identification number for an axis': the bar chart references an axis id the plot area does not declare. No schema rule enforces it (judgement: invalid, a dangling reference).",
          "The chart group's value-axis c:axId points to 99999, declared by no axis (the pptxgenjs secondary-axis fault).",
          office=f"{SKILL_PPTX}: \"pptxgenjs writes axis ids it never declares, and PowerPoint discards that chart and reports the file as corrupt\"")
def ref_undeclared_axis(pkg):
    part = chart_part(pkg)
    root = pkg.xml(part)
    group = next((g for g in root.iter() if isinstance(g.tag, str) and g.tag.startswith(f"{{{C}}}")
                  and g.tag.endswith("Chart") and g.tag not in (c("chart"), c("plotArea"))), None)
    if group is None:
        raise NotApplicable("no chart group")
    ax = group.findall(c("axId"))
    if len(ax) < 2:
        raise NotApplicable("chart group without two axIds")
    ax[1].set("val", "99999")
    pkg.put(part, root)


@mutation(X, "semantic", "invalid", "judgement", "DrawingML table (Part 1 21.1.3): a:tblGrid defines the columns and every a:tr holds one a:tc per grid column; a row with fewer cells leaves a column undefined. No schema rule counts them (judgement: invalid).",
          "Last a:tc of the second table row removed (fewer cells than a:gridCol).")
def sem_table_fewer_cells(pkg):
    name, root, tbl = slide_with(pkg, f".//{a('tbl')}", "a table")
    rows = tbl.findall(a("tr"))
    if len(rows) < 2:
        raise NotApplicable("table with one row")
    rows[1].remove(rows[1].findall(a("tc"))[-1])
    pkg.put(name, root)


@mutation(X, "semantic", "valid", "schema", "ST_DLblPos allows outEnd on any bar chart; ECMA-376 states no restriction for stacked groupings (spec-valid; PowerPoint refuses per the source in office).",
          'Bar chart regrouped as stacked (overlap 100) with data labels at dLblPos="outEnd".',
          office=f"{SKILL_PPTX}: \"On a stacked bar or column chart, dataLabelPosition must be ctr, inEnd, or inBase. outEnd corrupts the file\"",
          office_only=True)
def office_stacked_outend(pkg):
    part = chart_part(pkg)
    root = pkg.xml(part)
    bar = root.find(f".//{c('barChart')}")
    if bar is None:
        raise NotApplicable("no barChart")
    grouping = bar.find(c("grouping"))
    if grouping is None:
        grouping = etree.Element(c("grouping"))
        bar.find(c("barDir")).addnext(grouping)
    grouping.set("val", "stacked")
    for old in bar.findall(c("dLbls")) + bar.findall(c("overlap")):
        bar.remove(old)
    for ser in bar.findall(c("ser")):
        for old in ser.findall(c("dLbls")):
            ser.remove(old)
    dl = etree.Element(c("dLbls"))
    for tag, val in (("dLblPos", "outEnd"), ("showLegendKey", "0"), ("showVal", "1"), ("showCatName", "0"),
                     ("showSerName", "0"), ("showPercent", "0"), ("showBubbleSize", "0")):
        etree.SubElement(dl, c(tag), val=val)
    sers = bar.findall(c("ser"))
    sers[-1].addnext(dl)
    gap = bar.find(c("gapWidth"))
    ov = etree.Element(c("overlap"), val="100")
    (gap.addnext if gap is not None else dl.addnext)(ov)
    pkg.put(part, root)


# pptx negative controls
@mutation(X, "zip", "valid", "zip", "Identity: entries read and written back (deflate, same order).",
          "Repack without change (checks the harness).")
def ctl_repack_identity(pkg):  # noqa: F811
    pass


@mutation(X, "zip", "valid", "zip", "ECMA-376 Part 2 allows stored (uncompressed) ZIP items.",
          "Every item stored without compression.")
def ctl_zip_stored(pkg):  # noqa: F811
    stored(pkg)


@mutation(X, "zip", "valid", "opc", "ECMA-376 Part 2 does not fix the order of ZIP items.",
          "[Content_Types].xml written as the last ZIP item.")
def ctl_content_types_last(pkg):  # noqa: F811
    ct_last(pkg)


@mutation(X, "schema", "valid", "mce", "ECMA-376 Part 3 (Markup Compatibility): mc:AlternateContent with a Choice and a Fallback is legal in a shape tree.",
          'First shape of the first slide wrapped in mc:AlternateContent (Choice Requires="p14", Fallback with the same shape).')
def ctl_alternate_content(pkg):  # noqa: F811
    name, root, sp = slide_with(pkg, f".//{p('spTree')}/{p('sp')}", "a shape")
    alt = etree.Element(f"{{{MC}}}AlternateContent", nsmap={"mc": MC})
    choice = etree.SubElement(alt, f"{{{MC}}}Choice", Requires="p14",
                              nsmap={"p14": "http://schemas.microsoft.com/office/powerpoint/2010/main"})
    fallback = etree.SubElement(alt, f"{{{MC}}}Fallback")
    sp.addprevious(alt)
    choice.append(copy.deepcopy(sp))
    fallback.append(sp)
    pkg.put(name, root)


@mutation(X, "schema", "valid", "schema", "CT_Slide ends with an optional p:extLst whose p:ext elements carry any content under their own uri.",
          "p:extLst with an unknown p:ext added to the first slide.")
def ctl_extension_list(pkg):
    name = slides(pkg)[0]
    root = pkg.xml(name)
    lst = root.find(p("extLst"))
    if lst is None:
        lst = etree.SubElement(root, p("extLst"))
    ext = etree.SubElement(lst, p("ext"), uri="{6A8B0C3E-1F2D-4E5A-9B7C-RDOCXBENCH01}")
    etree.SubElement(ext, f"{{{BENCH}}}data", nsmap={"bench": BENCH}).text = "1"
    pkg.put(name, root)


@mutation(X, "schema", "valid", "mce", "ECMA-376 Part 3: elements and attributes of a namespace listed in mc:Ignorable are ignored by consumers that do not understand it.",
          "Unknown element and attribute of an ignorable namespace in the first slide's shape tree.")
def ctl_ignorable_unknown(pkg):  # noqa: F811
    name = slides(pkg)[0]
    root = add_ignorable(with_ns(pkg.xml(name), "bench", BENCH), "bench")
    tree = first(root, f".//{p('spTree')}", what="spTree")
    sp = first(tree, p("sp"), what="shape")
    sp.set(f"{{{BENCH}}}mark", "1")
    tree.append(etree.Element(f"{{{BENCH}}}note"))
    pkg.put(name, root)


@mutation(X, "package", "valid", "opc", "ECMA-376 Part 1 15.2.4: a Custom XML Data Storage part may be related from the presentation.",
          "Custom XML part added and related from the presentation part.")
def ctl_custom_xml_part(pkg):  # noqa: F811
    add_custom_xml(pkg)


@mutation(X, "schema", "valid", "schema", "No length limit on a:t.",
          "First a:t of the first slide given 100,000 characters.")
def ctl_long_text(pkg):
    name = slides(pkg)[0]
    root = pkg.xml(name)
    t = first(root, f".//{a('t')}", what="a:t")
    t.text = ("lorem ipsum dolor " * 5556)[:100000]
    pkg.put(name, root)


@mutation(X, "schema", "valid", "schema", "CT_TextBody needs bodyPr and one a:p: <a:bodyPr/><a:lstStyle/><a:p/> is a legal empty cell.",
          "Last table cell's txBody reduced to bodyPr, lstStyle and an empty <a:p/>.")
def ctl_empty_cell_paragraph(pkg):  # noqa: F811
    name, root, tbl = slide_with(pkg, f".//{a('tbl')}", "a table")
    tc = tbl.findall(f".//{a('tc')}")[-1]
    tb = first(tc, a("txBody"), what="cell txBody")
    for child in list(tb):
        tb.remove(child)
    etree.SubElement(tb, a("bodyPr"))
    etree.SubElement(tb, a("lstStyle"))
    etree.SubElement(tb, a("p"))
    pkg.put(name, root)


@mutation(X, "schema", "valid", "schema", "ST_SlideId: 2147483647 is the largest legal slide id.",
          'First p:sldId id="2147483647" (upper bound).')
def ctl_sldid_max(pkg):
    main, root = presentation(pkg)
    first(root, f"{p('sldIdLst')}/{p('sldId')}", what="sldId").set("id", "2147483647")
    pkg.put(main, root)


@mutation(X, "package", "valid", "opc", "ECMA-376 Part 1 13.3.6: a template's presentation part has content type ...presentationml.template.main+xml (.potx).",
          "Presentation part content type set to the template type, file named .potx.", ext=".potx")
def ctl_potx_template(pkg):
    set_main_content_type(pkg, "application/vnd.openxmlformats-officedocument.presentationml.template.main+xml")


def catalogue(fmt=None):
    return [m for m in CATALOGUE if fmt is None or m.fmt == fmt]


def apply(m: Mutation, base: bytes) -> bytes:
    pkg = Pkg(base)
    m.apply(pkg)
    return pkg.to_bytes()
