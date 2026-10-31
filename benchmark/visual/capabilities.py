"""Part 2 of the visual section: the feedback loop each stack gives an agent, every row tested on the corpus.

Each row runs the command an agent following that stack's skill would run and checks the output with neutral
tools (Poppler, zipfile + lxml). Status: ``pass`` (the stack does it), ``partial`` (only part of it, or through a
route its skill does not document; the evidence says which), ``unsupported`` (no means; the evidence says what was
probed), ``fail`` (it tried and got it wrong).
"""

from __future__ import annotations

import json
import math
import re
import shutil
import subprocess
import time
import zipfile
from pathlib import Path

from lxml import etree

import common
import metrics as M

WORK = common.WORK / "visual" / "capabilities"
CORPUS = common.CACHE / "corpus"
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
ENV_R = common.env_rdocx()
ENV_B = common.env_builtin()


def src(name: str) -> Path:
    kind = "docx" if name.startswith("docx") else "pptx"
    return CORPUS / kind / f"{name}.{kind}"


def real(fid: str) -> dict:
    man = json.loads((common.BENCH / "corpus" / "manifest.json").read_text())
    return next(f for f in man["files"] if f["id"] == fid)


def fresh(*parts) -> Path:
    d = WORK.joinpath(*parts)
    if d.exists():
        shutil.rmtree(d)
    d.mkdir(parents=True)
    return d


def soffice_pdf(path: Path, outdir: Path, kind: str = "docx", extra_filter: str | None = None) -> dict:
    skill = common.BUILTIN_DOCX if kind == "docx" else common.BUILTIN_PPTX
    to = "pdf" if extra_filter is None else f"pdf:writer_pdf_Export:{extra_filter}"
    r = common.run([common.PY, skill / "scripts" / "office" / "soffice.py", "--headless", "--convert-to", to,
                    "--outdir", outdir, path], env=ENV_B, cwd=outdir, timeout=600)
    pdf = outdir / (path.stem + ".pdf")
    r["pdf"] = pdf if pdf.exists() else None
    return r


def pdftoppm(pdf: Path, prefix: Path, first=None, last=None, fmt="-jpeg", dpi=100) -> dict:
    cmd = [common.PDFTOPPM, fmt, "-r", dpi]
    if first:
        cmd += ["-f", first, "-l", last or first]
    r = common.run(cmd + [pdf, prefix], timeout=600)
    r["files"] = sorted(prefix.parent.glob(prefix.name + "-*"))
    return r


def pdf_text(pdf: Path, first=None, last=None) -> str:
    cmd = ["pdftotext", "-enc", "UTF-8"]
    if first:
        cmd += ["-f", str(first), "-l", str(last or first)]
    return subprocess.run(cmd + [str(pdf), "-"], capture_output=True, text=True).stdout


def squash(s: str) -> str:
    return re.sub(r"\s+", " ", s).strip()


def row(capability, kind, stack, status, how, evidence, seconds=None, **extra):
    r = {"capability": capability, "kind": kind, "stack": stack, "status": status, "how": how, "evidence": evidence}
    if seconds is not None:
        r["seconds"] = round(seconds, 2)
    r.update(extra)
    return r


def pylib_probe() -> dict:
    """Public names of python-docx's Document and python-pptx's Presentation that could render or lay out."""
    code = ("import docx, pptx, json; from docx.document import Document; from pptx.presentation import Presentation; "
            "pat = ('render', 'pdf', 'png', 'image', 'thumb', 'layout', 'page', 'fit', 'overflow'); "
            "f = lambda c: sorted(n for n in dir(c) if not n.startswith('_') and any(p in n.lower() for p in pat)); "
            "print(json.dumps({'docx.Document': f(Document), 'pptx.Presentation': f(Presentation)}))")
    r = common.run([common.PY, "-c", code], env=common.env_pylib())
    return json.loads(r["stdout"])


PYLIB_NO_RENDER = None


def pylib_unsupported(capability, kind, what="python-docx and python-pptx have no renderer"):
    names = PYLIB_NO_RENDER
    return row(capability, kind, "pylib", "unsupported", "-",
               f"{what}; public names matching render/pdf/png/image/thumb/layout/page/fit/overflow: {names}")


# --- 1. one page or slide straight to PNG, and 2. selected pages ------------------------------------------
def render_rows() -> list[dict]:
    out = []
    d = fresh("render")
    # rdocx: one page
    r = common.run([common.RDOCX, "render", src("docx-M"), "-o", d / "r1", "--pages", 1, "--dpi", 100, "--font-dir",
                    common.FONTS_DIR], env=ENV_R)
    files = sorted((d / "r1").glob("*.png")) if (d / "r1").exists() else []
    out.append(row("one page or slide to PNG", "docx", "rdocx", "pass" if r["rc"] == 0 and len(files) == 1 else "fail",
                   "rdocx render docx-M.docx -o DIR --pages 1 --dpi 100 --font-dir FONTS",
                   f"{len(files)} PNG written ({files[0].name if files else '-'}), one command, no PDF", r["seconds"]))
    r = common.run([common.RPPTX, "render", src("pptx-M"), "-o", d / "p1", "--slide", 1, "--dpi", 100, "--font-dir",
                    common.FONTS_DIR], env=ENV_R)
    files = sorted((d / "p1").glob("*.png")) if (d / "p1").exists() else []
    out.append(row("one page or slide to PNG", "pptx", "rdocx", "pass" if r["rc"] == 0 and len(files) == 1 else "fail",
                   "rpptx render pptx-M.pptx -o DIR --slide 1 --dpi 100 --font-dir FONTS",
                   f"{len(files)} PNG written, one command, no PDF", r["seconds"]))
    # builtin: soffice.py to a PDF of the whole file, then pdftoppm
    for kind, name, dpi in (("docx", "docx-M", 100), ("pptx", "pptx-M", 150)):
        dd = fresh("render", f"b_{kind}")
        s = soffice_pdf(src(name), dd, kind)
        p = pdftoppm(s["pdf"], dd / "page", 1, 1, dpi=dpi) if s["pdf"] else {"rc": 1, "files": [], "seconds": 0}
        ok = s["pdf"] is not None and len(p["files"]) == 1
        out.append(row("one page or slide to PNG", kind, "builtin", "pass" if ok else "fail",
                       f"soffice.py --headless --convert-to pdf {name}.{kind}; pdftoppm -jpeg -r {dpi} -f 1 -l 1",
                       f"two steps: LibreOffice converts the whole file ({M.page_sizes(s['pdf']).__len__() if s['pdf'] else 0} "
                       f"pages) to PDF, pdftoppm rasterises page 1 ({len(p['files'])} JPEG)",
                       s["seconds"] + p["seconds"], steps={"soffice_pdf": round(s["seconds"], 2), "pdftoppm": round(p["seconds"], 2)}))
    out.append(pylib_unsupported("one page or slide to PNG", "docx+pptx"))

    # selected pages: 3 and 5 of docx-L, slides 3 and 5 of pptx-L
    r = common.run([common.RDOCX, "render", src("docx-L"), "-o", d / "sel_r", "--pages", "3,5", "--dpi", 100,
                    "--font-dir", common.FONTS_DIR], env=ENV_R)
    files = sorted(p.name for p in (d / "sel_r").glob("*.png")) if (d / "sel_r").exists() else []
    ok = r["rc"] == 0 and len(files) == 2 and any("page3" in f for f in files) and any("page5" in f for f in files)
    out.append(row("only selected pages", "docx", "rdocx", "pass" if ok else "fail",
                   "rdocx render docx-L.docx -o DIR --pages 3,5 --dpi 100 --font-dir FONTS",
                   f"{files} from a {real_pages('docx-L')}-page document, nothing else laid out to images", r["seconds"]))
    r = common.run([common.RPPTX, "convert", src("pptx-L"), "--to", "png", "--slides", "3,5", "--dpi", 100,
                    "--font-dir", common.FONTS_DIR, "-o", d / "sel_p.png"], env=ENV_R)
    files = sorted(p.name for p in d.glob("sel_p*.png"))
    out.append(row("only selected pages", "pptx", "rdocx", "pass" if r["rc"] == 0 and len(files) == 2 else "fail",
                   "rpptx convert pptx-L.pptx --to png --slides 3,5 --dpi 100 --font-dir FONTS -o OUT.png",
                   f"{files} (rpptx render takes one --slide; convert takes a range)", r["seconds"]))
    for kind, name in (("docx", "docx-L"), ("pptx", "pptx-L")):
        dd = fresh("render", f"sel_b_{kind}")
        s = soffice_pdf(src(name), dd, kind)
        files = []
        secs = s["seconds"]
        for pg in (3, 5):
            p = pdftoppm(s["pdf"], dd / f"p{pg}", pg, pg)
            files += [f.name for f in p["files"]]
            secs += p["seconds"]
        out.append(row("only selected pages", kind, "builtin", "pass" if len(files) == 2 else "fail",
                       f"soffice.py --convert-to pdf {name}.{kind}; pdftoppm -f 3 -l 3 and -f 5 -l 5",
                       f"{files}; LibreOffice converts all {len(M.page_sizes(s['pdf'])) if s['pdf'] else 0} pages first "
                       f"({s['seconds']:.1f} s of the {secs:.1f} s)", secs))
    out.append(pylib_unsupported("only selected pages", "docx+pptx"))
    return out


def real_pages(name):
    man = json.loads((common.BENCH / "corpus" / "manifest.json").read_text())
    f = next(f for f in man["files"] if f["id"] == name)
    return f["pages"].get("rdocx") or f["pages"].get("libreoffice")


# --- 3. tracked changes and 4. comments in the render ---------------------------------------------------
def revision_windows(docx_path: Path, ctx=40):
    """For each w:del of the body: the paragraph text around it as the tracked view reads (deleted text kept) and
    as the accepted view reads (deleted text dropped), whitespace squashed."""
    root = etree.fromstring(zipfile.ZipFile(docx_path).read("word/document.xml"))
    out = []
    for p in root.iter(W + "p"):
        segs = []
        for el in p.iter(W + "t", W + "delText"):
            anc = el.getparent()
            kind = "normal"
            while anc is not None and anc is not p:
                if anc.tag == W + "del":
                    kind = "del"
                elif anc.tag == W + "ins" and kind == "normal":
                    kind = "ins"
                anc = anc.getparent()
            segs.append((el.text or "", kind))
        for k, (txt, kind) in enumerate(segs):
            if kind != "del" or not txt.strip():
                continue
            before = "".join(t for t, _ in segs[:k])
            after = "".join(t for t, kd in segs[k + 1:])
            after_acc = "".join(t for t, kd in segs[k + 1:] if kd != "del")
            before_acc = "".join(t for t, kd in segs[:k] if kd != "del")
            b = before[-ctx:].split(" ", 1)[-1] if len(before) > ctx else before
            a = after[:ctx].rsplit(" ", 1)[0] if len(after) > ctx else after
            ba = before_acc[-ctx:].split(" ", 1)[-1] if len(before_acc) > ctx else before_acc
            aa = after_acc[:ctx].rsplit(" ", 1)[0] if len(after_acc) > ctx else after_acc
            nxt = next((kd for t, kd in segs[k + 1:] if t.strip()), None)
            out.append({"deleted": txt, "tracked": squash(b + txt + a), "accepted": squash(ba + aa),
                        "replacement": nxt == "ins"})
    return out


def comment_texts(docx_path: Path):
    root = etree.fromstring(zipfile.ZipFile(docx_path).read("word/comments.xml"))
    return [squash("".join(t.text or "" for t in c.iter(W + "t"))) for c in root.iter(W + "comment")]


def shown(windows, pdf: Path):
    text = squash(pdf_text(pdf))
    tracked = sum(1 for w in windows if w["tracked"] in text)
    accepted = sum(1 for w in windows if w["accepted"] in text)
    return tracked, accepted


def review_rows() -> list[dict]:
    out = []
    d = fresh("review")
    doc = src("docx-M")
    wins = revision_windows(doc)
    comments = comment_texts(doc)
    n = len(wins)
    # rdocx: default view, tracked view, legal style
    pdfs = {}
    for view, extra in (("accepted", []), ("tracked", ["--revision-view", "tracked"]),
                        ("tracked-legal", ["--revision-view", "tracked", "--revision-style", "legal"])):
        pdf = d / f"rdocx_{view}.pdf"
        r = common.run([common.RDOCX, "convert", doc, "--to", "pdf", "--font-dir", common.FONTS_DIR, *extra, "-o", pdf],
                       env=ENV_R)
        pdfs[view] = (pdf, r)
    t_acc = shown(wins, pdfs["accepted"][0])
    t_trk = shown(wins, pdfs["tracked"][0])
    ok = t_trk[0] >= 0.9 * n  # (a few tracked contexts also occur elsewhere: the corpus repeats sentence templates)
    out.append(row("tracked changes shown in the render", "docx", "rdocx", "pass" if ok else "fail",
                   "rdocx render docx-M.docx -o DIR --revision-view tracked [--revision-style legal] (same flag on convert --to pdf)",
                   f"{n} deletions in docx-M: their text with its context is in the tracked-view PDF for {t_trk[0]}, in the "
                   f"default (accepted) view for {t_acc[0]} (accepted context found for {t_acc[1]}): the default render hides "
                   "revisions, the flag shows them (struck deletions, underlined insertions, change bars; legal: colours)",
                   pdfs["tracked"][1]["seconds"], default_view_shows_revisions=t_acc[0] > 0))
    s = soffice_pdf(doc, fresh("review", "lo"))
    t_lo = shown(wins, s["pdf"])
    ok = t_lo[0] >= 0.9 * n
    out.append(row("tracked changes shown in the render", "docx", "builtin", "pass" if ok else "fail",
                   "soffice.py --headless --convert-to pdf docx-M.docx (the skill's render)",
                   f"{n} deletions: deleted text with its context in LibreOffice's PDF for {t_lo[0]}, accepted context for "
                   f"{t_lo[1]}; LibreOffice exports the document as it displays it, changes shown (its default)",
                   s["seconds"], default_view_shows_revisions=t_lo[0] > 0))
    out.append(pylib_unsupported("tracked changes shown in the render", "docx"))

    # comments
    def found(pdf):
        text = squash(pdf_text(pdf))
        return sum(1 for c in comments if c in text)

    c_r = max(found(pdfs[v][0]) for v in pdfs)
    out.append(row("comments shown in the render", "docx", "rdocx", "pass" if c_r >= 0.9 * len(comments) else "unsupported",
                   "rdocx render / convert --to pdf (any revision view)",
                   f"{len(comments)} comments in docx-M; comment text found in rdocx's PDFs: {c_r}. No flag puts comments "
                   "in the render (rdocx render --help, convert --help, the to_pdf / render_pages stubs); an agent reads "
                   "them with `rdocx comment list --json` (anchor_text gives the text each covers)"))
    c_lo = found(s["pdf"])
    s2 = soffice_pdf(doc, fresh("review", "lo_margin"), extra_filter='{"ExportNotesInMargin":{"type":"boolean","value":"true"}}')
    c_lo2 = found(s2["pdf"]) if s2["pdf"] else 0
    out.append(row("comments shown in the render", "docx", "builtin",
                   "pass" if c_lo >= 0.9 * len(comments) else ("partial" if c_lo2 >= 0.9 * len(comments) else "unsupported"),
                   "soffice.py --headless --convert-to pdf docx-M.docx",
                   f"{len(comments)} comments; found in LibreOffice's default PDF: {c_lo}. With an export option the skill "
                   f"does not document (--convert-to 'pdf:writer_pdf_Export:{{\"ExportNotesInMargin\":...true}}'): {c_lo2}, "
                   "in the page margin", s["seconds"]))
    out.append(pylib_unsupported("comments shown in the render", "docx"))
    return out, {"tracked_pdfs": {k: str(v[0]) for k, v in pdfs.items()}, "lo_pdf": str(s["pdf"]), "windows": wins}


# --- 5. thumbnail grid -----------------------------------------------------------------------------------
def grid_rows() -> list[dict]:
    out = []
    d = fresh("grid")
    # builtin pptx: thumbnail.py
    r = common.run([common.PY, common.BUILTIN_PPTX / "scripts" / "thumbnail.py", src("pptx-M"), d / "pptx-M-grid"],
                   env=ENV_B, cwd=d, timeout=600)
    files = sorted(p.name for p in d.glob("pptx-M-grid*.jpg"))
    out.append(row("thumbnail grid", "pptx", "builtin", "pass" if r["rc"] == 0 and files else "fail",
                   "python scripts/thumbnail.py pptx-M.pptx pptx-M-grid",
                   f"{len(files)} labelled grid JPEG(s) for 60 slides ({', '.join(files[:3])}...), slides labelled with "
                   "their XML file name", r["seconds"], files=[str(d / f) for f in files]))
    # rpptx: handouts (documented: render_all_handouts(slides_per_page, dpi))
    code = ("import rpptx, sys, pathlib; prs = rpptx.Presentation(sys.argv[1]); "
            "pages = prs.render_all_handouts(9, 60, font_dir=sys.argv[3]); out = pathlib.Path(sys.argv[2]); "
            "[(out / f'handout-{i + 1}.png').write_bytes(b) for i, b in enumerate(pages)]; print(len(pages))")
    r = common.run([common.R_PY, "-c", code, src("pptx-M"), d, common.FONTS_DIR], env=ENV_R, timeout=600)
    files = sorted(d.glob("handout-*.png"))
    want = math.ceil(60 / 9)
    out.append(row("thumbnail grid", "pptx", "rdocx", "pass" if r["rc"] == 0 and len(files) == want else "fail",
                   "prs.render_all_handouts(9, 60, font_dir=FONTS) (Python; also to_handout_pdf(9)); rpptx thumbnail is slide 1 only",
                   f"{len(files)} handout pages of 9 slides for 60 slides (expected {want}); handout layout, slides not "
                   "labelled with their number", r["seconds"], files=[str(f) for f in files]))
    # docx: builtin thumbnail.py refuses .docx; rdocx has no grid
    r = common.run([common.PY, common.BUILTIN_PPTX / "scripts" / "thumbnail.py", src("docx-M"), d / "docx-M-grid"],
                   env=ENV_B, cwd=d)
    out.append(row("thumbnail grid", "docx", "builtin", "unsupported",
                   "python scripts/thumbnail.py docx-M.docx (pptx skill's script)",
                   f"exit {r['rc']}: {(r['stderr'] or r['stdout']).strip()[-120:]}; the docx skill has no grid script"))
    stub = Path(subprocess.run([str(common.R_PY), "-c", "import rdocx, os; print(os.path.dirname(rdocx.__file__))"],
                               capture_output=True, text=True).stdout.strip()) / "_rdocx.pyi"
    hits = [ln.strip()[:60] for ln in stub.read_text().splitlines() if re.search(r"handout|thumbnail|contact|grid_image", ln, re.I)]
    out.append(row("thumbnail grid", "docx", "rdocx", "unsupported", "-",
                   f"no grid or handout for documents: _rdocx.pyi lines matching handout|thumbnail|contact: {hits}; "
                   "rdocx render then a Pillow montage is the way"))
    out.append(pylib_unsupported("thumbnail grid", "docx+pptx"))
    return out


# --- 6. page of each heading ----------------------------------------------------------------------------
def norm(s: str) -> str:
    """A heading as compared: a leading section label ("2.1", "Annex A") dropped, letters and digits only."""
    s = re.sub(r"^\s*(\d+(\.\d+)*\.?|annex [a-z0-9]+[:.]?)\s+", "", s.strip().lower())
    return re.sub(r"[^a-z0-9]", "", s)


def same(a: str, b: str) -> bool:
    return bool(a) and a == b


def outline(pdf: Path) -> list[tuple[int, str]]:
    xml = subprocess.run(["pdftohtml", "-xml", "-i", "-stdout", "-q", str(pdf)], capture_output=True, text=True).stdout
    return [(int(m.group(1)), re.sub(r"<[^>]+>", "", m.group(2))) for m in re.finditer(r'<item page="(\d+)">(.*?)</item>', xml, re.S)]


def seq_match(truth, items):
    """Match each truth heading, in order, to the next item (key, value) with the same text; value or None."""
    out, k = [], 0
    for key in truth:
        for j in range(k, len(items)):
            if same(key, items[j][0]):
                out.append(items[j][1])
                k = j + 1
                break
        else:
            out.append(None)
    return out


PYLIB_PAGES = r"""
import sys, json, docx
from docx.oxml.ns import qn
# Word's own pages at its last save: a page starts at each w:lastRenderedPageBreak (python-docx:
# Paragraph.rendered_page_breaks) and after each next-page section break that is not followed by one.
d = docx.Document(sys.argv[1])
body = d.element.body
LRPB, T = qn("w:lastRenderedPageBreak"), qn("w:t")
pages, n, pending = [], 1, False
for el in body.iterchildren():
    if el.tag == qn("w:sectPr"):
        continue
    lead = 0
    for node in el.iter(LRPB, T):
        if node.tag == T and (node.text or "").strip():
            break
        lead += node.tag == LRPB
    if pending and lead == 0:
        n += 1
    pending = False
    n += lead
    pages.append(n)
    n += len(el.findall(".//" + LRPB)) - lead
    for sp in el.iter(qn("w:sectPr")):  # a section break in this block (a paragraph, or one inside a control)
        t = sp.find(qn("w:type"))
        pending = t is None or t.get(qn("w:val")) in ("nextPage", "oddPage", "evenPage")
print(json.dumps({"pages": pages, "total": n, "rendered_breaks": len(body.findall(".//" + LRPB))}))
"""


def heading_rows() -> list[dict]:
    """Page of each heading against Word's own pages: the outline (bookmarks) of the PDF Word saved of two
    reports. A heading counts when its text is found exactly once among the document's top-level paragraphs."""
    out = []
    fresh("headings")
    cases = ["gov-uk-dcms-taylor-review-past-schemes", "gov-uk-dcms-taylor-review-interim-evaluation"]
    tallies = {"rdocx": 0, "builtin": 0, "pylib": 0}
    n = 0
    detail, totals = {}, {}
    secs = {"rdocx": 0.0, "builtin": 0.0, "pylib": 0.0}
    for fid in cases:
        f = real(fid)
        path = common.BENCH / f["path"]
        word = outline(common.BENCH / f["pdf_twin"]["path"])
        t0 = time.perf_counter()
        text = json.loads(common.run([common.RDOCX, "text", "--json", path], env=ENV_R)["stdout"])
        lay = json.loads(common.run([common.RDOCX, "layout", "--json", "--font-dir", common.FONTS_DIR, path],
                                    env=ENV_R)["stdout"])
        secs["rdocx"] += time.perf_counter() - t0
        page = {b["body_index"]: b["fragments"][0]["physical_page"] for b in lay["body_items"] if b["fragments"]}
        paras = [(norm(p["text"]), p["body_index"]) for p in text["paragraphs"] if p["path"] == []]
        bodies = seq_match([norm(t) for _, t in word], paras)
        s = soffice_pdf(path, fresh("headings", fid))
        secs["builtin"] += s["seconds"]
        lo = [(norm(t), pg) for pg, t in outline(s["pdf"])]
        r = common.run([common.PY, "-c", PYLIB_PAGES, path], env=common.env_pylib())
        secs["pylib"] += r["seconds"]
        py = json.loads(r["stdout"])
        totals[fid] = {"word": len(M.page_sizes(common.BENCH / f["pdf_twin"]["path"])), "pylib": py["total"]}
        det = []
        # unambiguous headings only: found once among the document's top-level paragraphs
        counts = {}
        for k, _ in paras:
            counts[k] = counts.get(k, 0) + 1
        kept = [(pg, t, b) for (pg, t), b in zip(word, bodies) if b is not None and counts[norm(t)] == 1]
        lo_pages = seq_match([norm(t) for _, t, _ in kept], lo)
        for (pg, t, b), lp in zip(kept, lo_pages):
            got = {"rdocx": page.get(b), "builtin": lp, "pylib": py["pages"][b] if b < len(py["pages"]) else None}
            n += 1
            for k, v in got.items():
                tallies[k] += v == pg
            det.append({"heading": t[:50], "word_page": pg, **got})
        detail[fid] = {"headings": det, "word_outline_items": len(word), "pages_total": totals[fid]}
    out.append(row("page of each heading or block", "docx", "rdocx", "pass",
                   "rdocx layout --json --font-dir FONTS F (page of every body block) + rdocx text --json",
                   f"a layout API for every block (rdocx's own pagination). Against Word's PDF outline of two Word reports: "
                   f"Word's page for {tallies['rdocx']} of {n} headings", secs["rdocx"],
                   accuracy=round(tallies["rdocx"] / n, 3)))
    out.append(row("page of each heading or block", "docx", "builtin", "partial",
                   "not in the skill: soffice.py to PDF, then read the PDF outline (pdftohtml -xml) or pdftotext page by page",
                   f"no layout API; headings only, from LibreOffice's PDF bookmarks: Word's page for {tallies['builtin']} "
                   f"of {n}", secs["builtin"], accuracy=round(tallies["builtin"] / n, 3)))
    r = common.run([common.PY, "-c", PYLIB_PAGES, src("docx-M")], env=common.env_pylib())
    nb = json.loads(r["stdout"])["rendered_breaks"]
    out.append(row("page of each heading or block", "docx", "pylib", "partial",
                   "python-docx: count w:lastRenderedPageBreak (Paragraph.rendered_page_breaks) and next-page section breaks",
                   f"Word's own pages as of its last save: Word's page for {tallies['pylib']} of {n} headings (page totals "
                   f"{ {k: v['pylib'] for k, v in totals.items()} } against Word's { {k: v['word'] for k, v in totals.items()} }); "
                   f"only for files Word saved (docx-M, written by docx-js: {nb} rendered page breaks) and stale after any edit",
                   secs["pylib"], accuracy=round(tallies["pylib"] / n, 3)))
    return out, detail


# --- 7. text overflow report for slides ----------------------------------------------------------------
FIXTURE = r'''
import sys
from pptx import Presentation
from pptx.util import Inches, Pt
prs = Presentation(); prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
s = prs.slides.add_slide(prs.slide_layouts[6])
long = " ".join(["Overflowing frame text that cannot fit in two inches by one inch at twenty-four points."] * 3)
for name, (x, w, h, text) in {"Overflows": (1, 2, 1, long), "Fits": (5, 6, 1, "Short title")}.items():
    tb = s.shapes.add_textbox(Inches(x), Inches(1), Inches(w), Inches(h)); tb.name = name
    tf = tb.text_frame; tf.word_wrap = True; tf.text = text
    for r in tf.paragraphs[0].runs:
        r.font.size = Pt(24); r.font.name = "Calibri"
prs.save(sys.argv[1])
'''

PYLIB_FIT = r'''
import sys, json
from pptx import Presentation
prs = Presentation(sys.argv[1]); font = sys.argv[2]; res = {}
for shape in prs.slides[0].shapes:
    tf = shape.text_frame
    before = tf.paragraphs[0].runs[0].font.size.pt
    tf.fit_text(font_family="Carlito", max_size=int(before), font_file=font)  # sets the size that fits
    after = tf.paragraphs[0].runs[0].font.size.pt
    res[shape.name] = {"size": before, "best_fit": after, "overflow": after < before}
print(json.dumps(res))
'''


def overflow_rows() -> list[dict]:
    out = []
    d = fresh("overflow")
    fx = d / "fit-fixture.pptx"
    common.run([common.PY, "-c", FIXTURE, fx], env=common.env_pylib(), check=True)
    truth = {"Overflows": True, "Fits": False}
    r = common.run([common.RPPTX, "fit", "--json", "--font-dir", common.FONTS_DIR, fx], env=ENV_R)
    rep = json.loads(r["stdout"])
    slide = {fr["name"]: fr["overflow"] for fr in rep["frames"] if fr.get("owner_kind") == "slide"}
    ok = all(slide.get(k) == v for k, v in truth.items())
    deck = {}
    for name in ("pptx-M", "pptx-L", "pptx-M-pythonpptx"):
        rr = common.run([common.RPPTX, "fit", "--json", "--font-dir", common.FONTS_DIR, src(name)], env=ENV_R)
        if rr["rc"] in (0, 1):
            js = json.loads(rr["stdout"])
            deck[name] = sum(1 for fr in js["frames"] if fr["overflow"] and fr.get("owner_kind") == "slide")
        else:
            deck[name] = f"exit {rr['rc']}: {rr['stderr'].strip()[-90:]}"
    out.append(row("text-overflow report for slides", "pptx", "rdocx", "pass" if ok else "fail",
                   "rpptx fit --json --font-dir FONTS deck.pptx (exit 1 when a frame overflows)",
                   f"fixture (one frame overflows, one fits): {slide}, exit {r['rc']}; every frame of slides, layouts, "
                   f"masters, notes and table cells with the font scale it needs; slide frames flagged in the corpus decks: {deck}",
                   r["seconds"], corpus_overflowing_slide_frames=deck))
    hits = []
    for script in sorted((common.BUILTIN_PPTX / "scripts").rglob("*.py")):
        if re.search(r"overflow|fit_text|text_fit", script.read_text(errors="replace"), re.I):
            hits.append(script.name)
    out.append(row("text-overflow report for slides", "pptx", "builtin", "unsupported",
                   "visual QA: render and look (SKILL.md, 'Text overflow ... check this first')",
                   f"no script reports overflow (scripts matching overflow|fit_text: {hits}); the skill asks the agent "
                   "to find it on the LibreOffice images"))
    fit = d / "fit-fixture-pylib.pptx"
    shutil.copy(fx, fit)
    r = common.run([common.PY, "-c", PYLIB_FIT, fit, common.FONTS_DIR / "Carlito-Regular.ttf"], env=common.env_pylib())
    res = json.loads(r["stdout"]) if r["rc"] == 0 else {}
    ok = res and all(res[k]["overflow"] == v for k, v in truth.items())
    out.append(row("text-overflow report for slides", "pptx", "pylib", "partial" if ok else "fail",
                   "python-pptx TextFrame.fit_text(font_family, max_size, font_file=...) on a copy, compare the size it sets",
                   f"fixture: {res}; a workaround: it rewrites the font size, needs the font file, measures one font "
                   "for the whole frame and ignores insets and autofit; no report of the deck", r["seconds"]))
    return out


# --- 8. substituted-font warning -------------------------------------------------------------------------
def font_rows() -> list[dict]:
    out = []
    d = fresh("fonts")
    f = real("gov-uk-pins-local-plan-webinar-qa")
    path = common.BENCH / f["path"]
    r = common.run([common.RDOCX, "render", path, "-o", d / "r", "--pages", 1, "--dpi", 60, "--font-dir",
                    common.FONTS_DIR], env=ENV_R)
    warn = [ln for ln in r["stderr"].splitlines() if "approximate" in ln]
    out.append(row("warns when a font is substituted", "docx", "rdocx", "pass" if warn else "fail",
                   "any rdocx render / convert (stderr); Python: render_*_with_diagnostics",
                   f"Local Plan Q&A (set in Aptos): {warn[:2]}", r["seconds"]))
    s = soffice_pdf(path, fresh("fonts", "lo"))
    said = [ln for ln in (s["stdout"] + s["stderr"]).splitlines() if "aptos" in ln.lower()]
    lo_fonts = sorted(M.pdf_fonts(s["pdf"]))
    out.append(row("warns when a font is substituted", "docx", "builtin", "unsupported",
                   "soffice.py --headless --convert-to pdf (stdout, stderr)",
                   f"no message ({len(said)} lines name Aptos); the substitute shows only in the PDF's fonts "
                   f"(pdffonts: {lo_fonts}), which the skill does not tell the agent to check", s["seconds"]))
    out.append(pylib_unsupported("warns when a font is substituted", "docx+pptx"))
    return out


# --- 9. time to the first picture of page 1 -------------------------------------------------------------
def first_picture_rows() -> list[dict]:
    out = []
    for name in ("docx-M", "docx-L", "pptx-M", "pptx-L"):
        kind = name[:4]
        d = fresh("first", name)
        if kind == "docx":
            cmd = [common.RDOCX, "render", src(name), "-o", d / "r", "--pages", 1, "--dpi", 100, "--font-dir", common.FONTS_DIR]
            how = "rdocx render F -o DIR --pages 1 --dpi 100 --font-dir FONTS"
        else:
            cmd = [common.RPPTX, "render", src(name), "-o", d / "r", "--slide", 1, "--dpi", 100, "--font-dir", common.FONTS_DIR]
            how = "rpptx render F -o DIR --slide 1 --dpi 100 --font-dir FONTS"
        r = common.run(cmd, env=ENV_R)
        out.append(row("time to the first picture of page 1 (one cold run)", name, "rdocx",
                       "pass" if r["rc"] == 0 else "fail", how, "one process", r["seconds"]))
        s = soffice_pdf(src(name), d, kind)
        p = pdftoppm(s["pdf"], d / "page", 1, 1, dpi=100 if kind == "docx" else 150)
        out.append(row("time to the first picture of page 1 (one cold run)", name, "builtin",
                       "pass" if p["files"] else "fail",
                       "soffice.py --headless --convert-to pdf F; pdftoppm -jpeg -r 100|150 -f 1 -l 1 F.pdf page",
                       "two processes, the whole file converted first", s["seconds"] + p["seconds"],
                       steps={"soffice_pdf": round(s["seconds"], 2), "pdftoppm": round(p["seconds"], 2)}))
    out.append(pylib_unsupported("time to the first picture of page 1 (one cold run)", "docx+pptx"))
    return out


def run() -> dict:
    global PYLIB_NO_RENDER
    PYLIB_NO_RENDER = pylib_probe()
    rows = []
    rows += render_rows()
    review, review_ctx = review_rows()
    rows += review
    rows += grid_rows()
    heads, head_detail = heading_rows()
    rows += heads
    rows += overflow_rows()
    rows += font_rows()
    rows += first_picture_rows()
    (WORK / "context.json").write_text(json.dumps({"review": review_ctx}, default=str, indent=1))
    return {"rows": rows, "heading_pages": head_detail, "pylib_probe": PYLIB_NO_RENDER}
