"""Neutral oracle helpers for the pptx feature tasks: zipfile and lxml only (no stack's reader).

Every task's check.py imports this module and calls `main(check)`, where `check(output, inputs)` returns a list
of errors (empty when the task is done) and optionally an info dict. `main` also runs every validator on a .pptx
output (rpptx validate, the built-in skill's validate.py, LibreOffice opening it) and prints one JSON record.
"""

from __future__ import annotations

import argparse
import json
import posixpath
import re
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

from lxml import etree

BENCH = Path(__file__).resolve().parents[3]
sys.path.insert(0, str(BENCH))
import common  # noqa: E402

A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
C = "http://schemas.openxmlformats.org/drawingml/2006/chart"
PR = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
NS = {"a": A, "p": P, "r": R, "c": C}
RT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
WARNINGS = []  # package warnings seen by any Pkg.integrity() call, reported by main()


def q(tag):
    pre, local = tag.split(":")
    return f"{{{NS[pre]}}}{local}"


class Pkg:
    """A read-only view of an OPC package: parts, relationships, content types and PresentationML helpers."""

    def __init__(self, path):
        self.path = Path(path)
        self.z = zipfile.ZipFile(self.path)
        self.names = [n for n in self.z.namelist() if not n.endswith("/")]
        self._xml = {}

    # --- package level ---------------------------------------------------------------------
    def has(self, part):
        return part.lstrip("/") in self.names

    def read(self, part):
        return self.z.read(part.lstrip("/"))

    def xml(self, part):
        part = part.lstrip("/")
        if part not in self._xml:
            self._xml[part] = etree.fromstring(self.read(part))
        return self._xml[part]

    @staticmethod
    def rels_name(part):
        d, f = posixpath.split(part.lstrip("/"))
        return posixpath.join(d, "_rels", f + ".rels")

    def rels(self, part):
        """[{id, type, target (absolute part name, no leading slash), mode}] of a part ('' for the package)."""
        name = "_rels/.rels" if part == "" else self.rels_name(part)
        if name not in self.names:
            return []
        out = []
        base = posixpath.dirname(part.lstrip("/"))
        for rel in self.xml(name).findall(f"{{{PR}}}Relationship"):
            mode = rel.get("TargetMode", "Internal")
            tgt = rel.get("Target")
            if mode == "Internal":
                tgt = tgt.lstrip("/") if tgt.startswith("/") else posixpath.normpath(posixpath.join(base, tgt))
            out.append({"id": rel.get("Id"), "type": rel.get("Type"), "target": tgt, "mode": mode})
        return out

    def rel(self, part, rid):
        return next((r["target"] for r in self.rels(part) if r["id"] == rid), None)

    def rels_of_type(self, part, suffix):
        return [r["target"] for r in self.rels(part) if r["type"].endswith("/" + suffix)]

    def content_types(self):
        root = self.xml("[Content_Types].xml")
        defaults = {d.get("Extension").lower(): d.get("ContentType") for d in root.findall(f"{{{CT}}}Default")}
        overrides = {o.get("PartName").lstrip("/"): o.get("ContentType") for o in root.findall(f"{{{CT}}}Override")}
        return defaults, overrides

    def content_type(self, part):
        defaults, overrides = self.content_types()
        part = part.lstrip("/")
        return overrides.get(part) or defaults.get(part.rsplit(".", 1)[-1].lower())

    def integrity(self):
        """Packaging errors: XML that does not parse, relationship targets missing, parts without a content
        type, relationship ids repeated. (An override naming a missing part, which pptxgenjs writes and Office
        accepts, is a warning: self.warnings.)"""
        errors = []
        self.warnings = []
        defaults, overrides = self.content_types()
        for n in self.names:
            if n == "[Content_Types].xml":
                continue
            if n.endswith(".xml") or n.endswith(".rels"):
                try:
                    self.xml(n)
                except etree.XMLSyntaxError as e:
                    errors.append(f"{n} does not parse: {e}")
                    continue
            if not n.endswith(".rels") and self.content_type(n) is None:
                errors.append(f"{n} has no content type")
        for o in overrides:
            if o not in self.names:
                self.warnings.append(f"override for missing part /{o}")
        WARNINGS.extend(self.warnings)
        sources = [""] + [n for n in self.names if not n.endswith(".rels") and n != "[Content_Types].xml"]
        for src in sources:
            ids = [r["id"] for r in self.rels(src)]
            if len(ids) != len(set(ids)):
                errors.append(f"repeated relationship id in rels of /{src}")
            for r in self.rels(src):
                if r["mode"] == "Internal" and r["target"] not in self.names:
                    errors.append(f"/{src or '(package)'} {r['id']} -> missing /{r['target']}")
        return errors

    def reachable(self):
        seen, todo = set(), [""]
        while todo:
            src = todo.pop()
            for r in self.rels(src):
                if r["mode"] == "Internal" and r["target"] not in seen and r["target"] in self.names:
                    seen.add(r["target"])
                    todo.append(r["target"])
        return seen

    def orphans(self):
        """Parts that no relationship chain from the package reaches (rels parts and content types aside)."""
        reach = self.reachable()
        return sorted(n for n in self.names if not n.endswith(".rels") and n != "[Content_Types].xml"
                      and n not in reach)

    # --- PresentationML ----------------------------------------------------------------------
    @property
    def main(self):
        return next(r["target"] for r in self.rels("") if r["type"].endswith("/officeDocument"))

    def sld_ids(self):
        """[(id, r:id, part, hidden)] in presentation order."""
        out = []
        lst = self.xml(self.main).find("p:sldIdLst", NS)
        for s in (lst if lst is not None else []):
            part = self.rel(self.main, s.get(q("r:id")))
            hidden = part is not None and self.has(part) and self.xml(part).get("show") == "0"
            out.append((s.get("id"), s.get(q("r:id")), part, hidden))
        return out

    def slides(self):
        return [p for _, _, p, _ in self.sld_ids()]

    def layout_of(self, slide):
        return (self.rels_of_type(slide, "slideLayout") or [None])[0]

    def master_of(self, layout):
        return (self.rels_of_type(layout, "slideMaster") or [None])[0]

    def theme_of(self, master):
        return (self.rels_of_type(master, "theme") or [None])[0]

    def notes_of(self, slide):
        return (self.rels_of_type(slide, "notesSlide") or [None])[0]

    def layout_name(self, layout):
        return self.xml(layout).find("p:cSld", NS).get("name")

    @staticmethod
    def para_text(p):
        out = []
        for child in p:
            if child.tag in (q("a:r"), q("a:fld")):
                t = child.find("a:t", NS)
                out.append(t.text or "" if t is not None else "")
            elif child.tag == q("a:br"):
                out.append("\n")
        return "".join(out)

    def paragraphs(self, part, root=None):
        root = self.xml(part) if root is None else root
        tree = root.find(".//p:cSld/p:spTree", NS) if root.tag != q("a:p") else None
        tree = root if tree is None else tree
        return [self.para_text(p) for p in tree.iter(q("a:p"))]

    def text(self, part, root=None):
        return "\n".join(self.paragraphs(part, root))

    def shapes(self, part):
        """Every sp, pic, graphicFrame, grpSp, cxnSp under the slide's spTree (groups walked)."""
        tree = self.xml(part).find("p:cSld/p:spTree", NS)
        kinds = (q("p:sp"), q("p:pic"), q("p:graphicFrame"), q("p:grpSp"), q("p:cxnSp"))
        return [e for e in tree.iter(*kinds) if e is not tree]

    @staticmethod
    def ph(shape):
        ph = shape.find("./*/p:nvPr/p:ph", NS)
        if ph is None:
            return None
        return {"type": ph.get("type", "obj"), "idx": ph.get("idx", "0")}

    @staticmethod
    def shape_id(shape):
        return shape.find("./*/p:cNvPr", NS).get("id")

    @staticmethod
    def shape_name(shape):
        return shape.find("./*/p:cNvPr", NS).get("name")

    def title(self, slide):
        for sh in self.shapes(slide):
            ph = self.ph(sh)
            if ph and ph["type"] in ("title", "ctrTitle"):
                return self.text(slide, sh).strip()
        return None

    def notes_text(self, slide):
        notes = self.notes_of(slide)
        if notes is None:
            return None
        for sh in self.shapes(notes):
            ph = self.ph(sh)
            if ph and ph["type"] == "body":
                return self.text(notes, sh)
        return None

    def slide_marker(self, slide, markers):
        """The first marker string found in the slide's title (else its text)."""
        t = (self.title(slide) or "") + "\n" + self.text(slide)
        return next((m for m in markers if m in t), None)


def canon(el):
    """A namespace-declaration-agnostic value of an element (tag, attributes, text, children), for comparisons."""
    if el is None:
        return None
    return (el.tag, tuple(sorted(el.attrib.items())), (el.text or "").strip(), tuple(canon(c) for c in el))


def find_one(pkg, slide, pred):
    return next((sh for sh in pkg.shapes(slide) if pred(sh)), None)


# --- validators ------------------------------------------------------------------------------------
def libreoffice_opens(path, timeout=180):
    with tempfile.TemporaryDirectory(prefix="lo_check_") as tmp:
        prof = Path(tmp) / "profile"
        env = common.env_builtin()
        try:
            subprocess.run([common.SOFFICE, f"-env:UserInstallation={prof.as_uri()}", "--headless", "--convert-to",
                            "pdf", "--outdir", tmp, str(path)], env=env, capture_output=True, timeout=timeout)
        except subprocess.TimeoutExpired:
            return False, "timeout"
        pdf = Path(tmp) / (Path(path).stem + ".pdf")
        if not pdf.exists():
            return False, "no PDF written"
        data = pdf.read_bytes()
        pages = len(re.findall(rb"/Type\s*/Page[^s]", data))
        return (data.startswith(b"%PDF") and pages > 0), f"{pages} pages"


def validators(path, original=None):
    path = Path(path)
    res, notes = {}, {}
    r = common.run([common.RPPTX, "validate", path], env=common.env_rdocx(), timeout=300)
    res["rpptx_validate"] = r["rc"] == 0
    if r["rc"]:
        notes["rpptx_validate"] = (r["stdout"] + r["stderr"]).strip()[-600:]
    cmd = [common.PY, common.BUILTIN_PPTX / "scripts/office/validate.py", path]
    if original:
        cmd += ["--original", original]
    r = common.run(cmd, env=common.env_builtin(), timeout=300)
    res["builtin_validate"] = r["rc"] == 0
    if r["rc"]:
        notes["builtin_validate"] = (r["stdout"] + r["stderr"]).strip()[-600:]
    ok, why = libreoffice_opens(path)
    res["libreoffice_opens"] = ok
    if not ok:
        notes["libreoffice_opens"] = why
    return res, notes


# --- entry point -----------------------------------------------------------------------------------
def main(check, pptx_output=True):
    ap = argparse.ArgumentParser()
    ap.add_argument("output")
    ap.add_argument("--inputs", default=None, help="the folder make_input.py wrote (default: the task's)")
    ap.add_argument("--original", default=None, help="passed to validate.py --original")
    ap.add_argument("--no-validate", action="store_true")
    a = ap.parse_args()
    out = Path(a.output)
    record = {"ok": False, "errors": [], "info": {}}
    if not out.exists():
        record["errors"].append(f"no output at {out}")
    else:
        try:
            res = check(out, Path(a.inputs) if a.inputs else None)
            errors, info = (res if isinstance(res, tuple) else (res, {}))
            record["errors"], record["info"] = list(errors), info
        except Exception as e:  # an unreadable output is a failed task, not a crashed oracle
            record["errors"].append(f"oracle could not read the output: {type(e).__name__}: {e}")
        if WARNINGS:
            record["info"]["package_warnings"] = WARNINGS
        record["ok"] = not record["errors"]
        if pptx_output and not a.no_validate and out.is_file():
            record["valid"], record["valid_notes"] = validators(out, a.original)
    print(json.dumps(record, indent=1))
    sys.exit(0 if record["ok"] else 1)
