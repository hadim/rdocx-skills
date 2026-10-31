"""Spec oracle: what ECMA-376 says about a package, checked with zipfile and lxml only.

Used to audit the ground truth (base files, mutated files, prevention outputs), not as a competitor: it reports
rule violations at four levels, each issue prefixed by its basis:

- ``zip``: the archive opens, every item's CRC matches, no two items share a name;
- ``opc`` (ECMA-376 Part 2): [Content_Types].xml present; every part has a content type; the package
  relationships name a main part; every internal relationship targets a part of the package; every .rels and
  [Content_Types].xml is schema-valid (opc-*.xsd);
- ``xml``: every XML part is well-formed XML 1.0;
- ``schema`` (ECMA-376 Part 1 / ISO/IEC 29500-4:2016 transitional XSDs): every part whose content type maps to a
  schema validates after Markup Compatibility (Part 3) preprocessing: content in mc:Ignorable namespaces is
  dropped, mc:AlternateContent is replaced by its Fallback; a non-ignorable namespace the schemas do not define
  stays and fails, as Part 3 requires.

Prose rules (dangling ids, duplicate annotation ids) are not checked here: those labels come from the clause
cited in mutations.py. The XSD files come from their publishers (validity/schemas.py): the transitional schemas of
ECMA-376 Part 4 (5th edition, the text of ISO/IEC 29500-4:2016), the OPC schemas of Part 2 and W3C's xml.xsd,
pinned by SHA-256; the oracle strips every mc attribute before validation.
"""

from __future__ import annotations

import io
import posixpath
import sys
import threading
import zipfile
from functools import lru_cache
from pathlib import Path

from lxml import etree

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import common  # noqa: E402

sys.path.insert(0, str(Path(__file__).resolve().parent))
import schemas as _schemas  # noqa: E402  (validity/schemas.py: ECMA-376 and W3C, pinned by SHA-256)

SCHEMAS = _schemas.ensure()
ISO = "transitional"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
PR_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

OFFICE = "application/vnd.openxmlformats-officedocument."
SCHEMA_BY_TYPE = {
    # WordprocessingML
    **{OFFICE + "wordprocessingml." + k + "+xml": f"{ISO}/wml.xsd" for k in (
        "document.main", "template.main", "document.macroEnabled.main", "styles", "settings", "numbering",
        "footnotes", "endnotes", "comments", "fontTable", "webSettings", "header", "footer")},
    "application/vnd.ms-word.document.macroEnabled.main+xml": f"{ISO}/wml.xsd",
    # PresentationML
    **{OFFICE + "presentationml." + k + "+xml": f"{ISO}/pml.xsd" for k in (
        "presentation.main", "template.main", "slideshow.main", "slide", "slideLayout", "slideMaster",
        "notesSlide", "notesMaster", "handoutMaster", "presProps", "viewProps", "comments", "commentAuthors")},
    OFFICE + "presentationml.tableStyles+xml": f"{ISO}/dml-main.xsd",
    # DrawingML and shared
    OFFICE + "theme+xml": f"{ISO}/dml-main.xsd",
    OFFICE + "drawingml.chart+xml": f"{ISO}/dml-chart.xsd",
    OFFICE + "extended-properties+xml": f"{ISO}/shared-documentPropertiesExtended.xsd",
    OFFICE + "custom-properties+xml": f"{ISO}/shared-documentPropertiesCustom.xsd",
    OFFICE + "customXmlProperties+xml": f"{ISO}/shared-customXmlDataProperties.xsd",
    "application/vnd.openxmlformats-package.relationships+xml": "opc/opc-relationships.xsd",
}
# ISO namespaces the transitional schemas define (anything else must be ignorable, Part 3)
KNOWN_NS = {
    "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "http://schemas.openxmlformats.org/presentationml/2006/main",
    "http://schemas.openxmlformats.org/drawingml/2006/main",
    "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "http://schemas.openxmlformats.org/drawingml/2006/chartDrawing",
    "http://schemas.openxmlformats.org/drawingml/2006/diagram",
    "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "http://schemas.openxmlformats.org/drawingml/2006/lockedCanvas",
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "http://schemas.openxmlformats.org/officeDocument/2006/sharedTypes",
    "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties",
    "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties",
    "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
    "http://schemas.openxmlformats.org/officeDocument/2006/customXml",
    "http://schemas.openxmlformats.org/officeDocument/2006/bibliography",
    "http://schemas.openxmlformats.org/schemaLibrary/2006/main",
    "http://schemas.openxmlformats.org/package/2006/relationships",
    "http://schemas.openxmlformats.org/package/2006/content-types",
    "http://www.w3.org/XML/1998/namespace",
    "urn:schemas-microsoft-com:vml",
    "urn:schemas-microsoft-com:office:office",
    "urn:schemas-microsoft-com:office:word",
    "urn:schemas-microsoft-com:office:powerpoint",
    "urn:schemas-microsoft-com:office:excel",
}
# Part 1 11.3, 13.3, 14.2, 15.2: the content type a part reached by these relationship types has
REL_CONTENT_TYPE = {
    **{k: OFFICE + "wordprocessingml." + k + "+xml" for k in (
        "styles", "settings", "numbering", "footnotes", "endnotes", "comments", "fontTable", "webSettings",
        "header", "footer")},
    **{k: OFFICE + "presentationml." + k + "+xml" for k in (
        "slide", "slideLayout", "slideMaster", "notesSlide", "notesMaster", "handoutMaster", "presProps",
        "viewProps", "tableStyles", "commentAuthors")},
    "theme": OFFICE + "theme+xml",
    "chart": OFFICE + "drawingml.chart+xml",
}
MAIN_TYPES = {
    "docx": {OFFICE + "wordprocessingml.document.main+xml", OFFICE + "wordprocessingml.template.main+xml"},
    "pptx": {OFFICE + "presentationml.presentation.main+xml", OFFICE + "presentationml.template.main+xml"},
}


_LOCK = threading.Lock()


@lru_cache(maxsize=None)
def schema(rel: str):
    path = SCHEMAS / rel
    return etree.XMLSchema(etree.parse(str(path), base_url=str(path)))


def _in_scope(el, prefix):
    return el.nsmap.get(prefix)


def mce_preprocess(root):
    """Apply ECMA-376 Part 3 as a consumer that understands only the ISO namespaces."""
    ignorable = set()
    for el in root.iter():
        if not isinstance(el.tag, str):
            continue
        val = el.get(f"{{{MC}}}Ignorable")
        if val:
            for prefix in val.split():
                uri = _in_scope(el, prefix)
                if uri:
                    ignorable.add(uri)
    # AlternateContent: no Choice is taken (they require extension namespaces); the Fallback content stays
    for alt in list(root.iter(f"{{{MC}}}AlternateContent")):
        parent = alt.getparent()
        if parent is None:
            continue
        fallback = alt.find(f"{{{MC}}}Fallback")
        idx = parent.index(alt)
        if fallback is not None:
            for k, child in enumerate(list(fallback)):
                parent.insert(idx + k, child)
        parent.remove(alt)
    for el in list(root.iter()):
        if not isinstance(el.tag, str):
            continue
        ns = etree.QName(el).namespace
        if ns in ignorable and el.getparent() is not None:
            el.getparent().remove(el)
            continue
        for att in list(el.attrib):
            ans = etree.QName(att).namespace
            if ans == MC or ans in ignorable:
                del el.attrib[att]
    return root


def audit(path) -> list[str]:
    """Every rule the package at ``path`` breaks, as ``basis: message`` strings (empty when none found)."""
    data = Path(path).read_bytes()
    issues: list[str] = []
    try:
        zf = zipfile.ZipFile(io.BytesIO(data))
        infos = zf.infolist()
    except zipfile.BadZipFile as e:
        return [f"zip: not a readable ZIP archive ({e})"]
    names = [i.filename for i in infos]
    seen = set()
    for n in names:
        if n.lower() in seen:
            issues.append(f"opc: equivalent part names in the ZIP ({n})")
        seen.add(n.lower())
    parts = {}
    for i in infos:
        if i.is_dir() or i.filename in parts:
            continue
        try:
            parts[i.filename] = zf.read(i)
        except Exception as e:  # CRC or deflate error
            issues.append(f"zip: item {i.filename} unreadable ({e})")
    if "[Content_Types].xml" not in parts:
        issues.append("opc: no [Content_Types].xml")
        return issues
    try:
        ct = etree.fromstring(parts["[Content_Types].xml"])
    except etree.XMLSyntaxError as e:
        return issues + [f"xml: [Content_Types].xml not well-formed ({e})"]
    if not schema("opc/opc-contentTypes.xsd").validate(ct):
        issues.append(f"opc: [Content_Types].xml schema: {schema('opc/opc-contentTypes.xsd').error_log.last_error.message}")
    defaults = {e.get("Extension", "").lower(): e.get("ContentType") for e in ct if e.tag == f"{{{CT_NS}}}Default"}
    overrides = {e.get("PartName", "").lower(): e.get("ContentType") for e in ct if e.tag == f"{{{CT_NS}}}Override"}

    def ctype(name):
        base = posixpath.basename(name)
        ext = base.rsplit(".", 1)[1].lower() if "." in base else ""
        return overrides.get("/" + name.lower()) or defaults.get(ext)

    for name in parts:
        if name == "[Content_Types].xml":
            continue
        if not ctype(name):
            issues.append(f"opc: part {name} has no content type")
    # relationships
    main = None
    for name in parts:
        if not name.endswith(".rels"):
            continue
        source_dir = posixpath.dirname(posixpath.dirname(name))
        try:
            rels = etree.fromstring(parts[name])
        except etree.XMLSyntaxError as e:
            issues.append(f"xml: {name} not well-formed ({e})")
            continue
        for rel in rels:
            if rel.get("TargetMode") == "External":
                continue
            target = rel.get("Target", "")
            resolved = target[1:] if target.startswith("/") else posixpath.normpath(posixpath.join(source_dir, target))
            if resolved not in parts:
                issues.append(f"opc: {name} {rel.get('Id')} targets missing part {resolved}")
            else:
                want = REL_CONTENT_TYPE.get(rel.get("Type", "").rsplit("/", 1)[-1])
                got = ctype(resolved)
                if want and got and got != want:
                    issues.append(f"opc: {resolved} is the target of a {rel.get('Type').rsplit('/', 1)[-1]} "
                                  f"relationship but has content type {got}")
            if name == "_rels/.rels" and rel.get("Type", "").endswith("/officeDocument"):
                main = resolved
    kind = "docx" if any(n.startswith("word/") for n in parts) else "pptx"
    if main is None:
        issues.append("opc: package relationships name no main (officeDocument) part")
    elif main in parts and ctype(main) not in MAIN_TYPES[kind]:
        issues.append(f"opc: main part {main} has content type {ctype(main)}")
    # well-formedness and schemas
    for name, blob in parts.items():
        t = ctype(name) or ""
        if not (name.endswith((".xml", ".rels")) or t.endswith("+xml") or t == "application/xml"):
            continue
        try:
            root = etree.fromstring(blob, etree.XMLParser(huge_tree=True))
        except etree.XMLSyntaxError as e:
            issues.append(f"xml: {name} not well-formed ({e})")
            continue
        rel = SCHEMA_BY_TYPE.get(t)
        if rel is None or name == "[Content_Types].xml":
            continue
        root = mce_preprocess(root)
        sch = schema(rel)
        with _LOCK:  # one XMLSchema object per schema, shared by threads: its error_log is per object
            ok = sch.validate(root)
            errs = [] if ok else [e for e in sch.error_log
                                  if e.type_name != "SCHEMAV_CVC_COMPLEX_TYPE_4" or "dcterms" not in e.message]
        if not ok:
            for e in errs[:3]:
                issues.append(f"schema: {name}:{e.line}: {e.message}")
            if len(errs) > 3:
                issues.append(f"schema: {name}: {len(errs) - 3} more error(s)")
    return issues


if __name__ == "__main__":
    for f in sys.argv[1:]:
        found = audit(f)
        print(f, "OK" if not found else "")
        for i in found:
            print("  ", i[:300])
