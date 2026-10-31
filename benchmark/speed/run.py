"""Speed: how long the operations an agent actually runs take, on realistic files, per stack.

    .cache/venv/bin/python speed/run.py                  # full run: results/speed.json (waits for a quiet machine)
    .cache/venv/bin/python speed/run.py --quick          # 1 run per command, no warm-up: .cache/work/speed/quick.json
    .cache/venv/bin/python speed/run.py --quick --only 'docx.read:docx-S:*'   # op:file:stack:variant patterns, comma-separated

Every measurement is the command an agent runs (process start included), timed with common.timeit: a warm-up run,
then the median of 5 runs (3 for LibreOffice operations and for anything whose warm-up took over 20 s; 1 when it
took over 60 s, flagged). Each output is checked cheaply (oracle.py: counts, anchors, revision marks, pages,
samples of text) and, when it is an Office file, validated by rdocx / rpptx validate and the built-in skill's
validate.py: a fast wrong answer is not a win. The machine load (os.getloadavg) is recorded before and after each
group of commands (one operation on one file, every stack). speed/NOTES.md has the method and the judgement calls.
"""
from __future__ import annotations

import argparse
import fnmatch
import json
import os
import re
import shutil
import statistics
import sys
import time
from pathlib import Path

sys.dont_write_bytecode = True
HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parent))
sys.path.insert(0, str(HERE))
import common  # noqa: E402
import oracle  # noqa: E402

WORK = common.WORK / "speed"
TMP = WORK / "tmp"
CORPUS = common.CACHE / "corpus"
FONTS = str(common.FONTS_DIR)
COMMENT_TEXT = "Speed benchmark comment"
STACKS = ("rdocx", "builtin", "pylib")
CLI_PAGES: dict = {}


# ------------------------------------------------------------------------------------------ environments
def environments() -> dict:
    out = {}
    for name, env in (("rdocx", common.env_rdocx()), ("builtin", common.env_builtin()), ("pylib", common.env_pylib())):
        env["TMPDIR"] = str(TMP)
        env["BUILTIN_DOCX"] = str(common.BUILTIN_DOCX)
        env["BUILTIN_PPTX"] = str(common.BUILTIN_PPTX)
        out[name] = env
    out["builtin"]["TZ"] = "UTC"  # the corpus's node generators (deterministic dates)
    return out


ENV = environments()


class Task:
    """One stack's command for one operation on one file."""

    def __init__(self, op, file, stack, variant, argv=None, *, out=None, out_kind=None, lo=False, check=None,
                 original=None, unsupported=None, primary=True, mkdir=False, stdout_check=False, note=None):
        self.op, self.file, self.stack, self.variant = op, file, stack, variant
        self.argv, self.out, self.out_kind, self.lo, self.check = argv, out, out_kind, lo, check
        self.original, self.unsupported, self.primary, self.mkdir = original, unsupported, primary, mkdir
        self.stdout_check, self.note = stdout_check, note

    @property
    def key(self):
        return f"{self.op}:{self.file['id']}:{self.stack}:{self.variant}"


# ------------------------------------------------------------------------------------------ corpus
def corpus_files() -> list[dict]:
    manifest = json.loads((HERE.parent / "corpus" / "manifest.json").read_text())
    files = []
    for f in manifest["files"]:
        path = common.BENCH / f["path"]
        if not path.exists():
            continue
        files.append({
            "id": f["id"], "kind": f["kind"], "origin": f["origin"], "size_class": f.get("size_class"),
            "path": str(path), "bytes": path.stat().st_size,
            "pages": f.get("pages") if f["kind"] == "docx" else None,
            "slides": f["counts"].get("slides") if f["kind"] == "pptx" else None,
        })
    order = {"synthetic": 0, "real": 1}
    return sorted(files, key=lambda f: (f["kind"], order[f["origin"]], f["bytes"]))


def plan_files() -> list[dict]:
    out = []
    for kind in ("docx", "pptx"):
        for size in ("S", "M", "L"):
            fid = f"{kind}-{size}"
            out.append({"id": fid, "kind": kind, "origin": "plan", "size_class": size,
                        "path": str(CORPUS / "plans" / f"{fid}.json"), "bytes": (CORPUS / "plans" / f"{fid}.json").stat().st_size,
                        "pages": None, "slides": None, "reference": str(CORPUS / kind / f"{fid}.{kind}")})
    return out


# ------------------------------------------------------------------------------------------ checks
def in_order(words: list[str], tokens: list[str], gap: int = 80) -> bool:
    """The words in order in the token stream, at most `gap` tokens apart: a phrase that a Markdown table wraps
    over several lines, its columns interleaved."""
    starts = [i for i, tok in enumerate(tokens) if tok == words[0]]
    for i in starts:
        pos, ok = i, True
        for word in words[1:]:
            nxt = next((j for j in range(pos + 1, min(len(tokens), pos + 1 + gap)) if tokens[j] == word), None)
            if nxt is None:
                ok = False
                break
            pos = nxt
        if ok:
            return True
    return False


def found_samples(text: str, samples: list[str]) -> dict:
    """Each sample phrase in the output, whitespace collapsed; failing that, its words in order (table wrap)."""
    t = oracle.norm(text)
    tokens = re.findall(r"[A-Za-z0-9]+", text)
    exact = [s for s in samples if oracle.norm(s) in t]
    ordered = [s for s in samples if s not in exact and in_order(re.findall(r"[A-Za-z0-9]+", s), tokens)]
    got = exact + ordered
    return {"ok": len(got) == len(samples) and bool(samples), "samples": len(samples), "found": len(got),
            "found_in_order_only": len(ordered), "missing": [s for s in samples if s not in got]}


def pngs(folder: Path) -> int:
    return len(list(Path(folder).glob("*.png"))) if Path(folder).exists() else 0


def stats(path: str) -> dict:
    r = common.run([common.PY, "-I", HERE.parent / "corpus" / "ooxml_stats.py", path], timeout=300)
    try:
        return json.loads(r["stdout"])
    except json.JSONDecodeError:
        return {"error": (r["stderr"] or r["stdout"])[-300:]}


CREATE_KEYS = {
    "docx": ("paragraphs", "tables", "pictures", "comments", "revisions", "footnotes", "sections", "landscape_sections",
             "headers", "footers", "has_toc"),
    "pptx": ("slides", "paragraphs", "tables", "pictures", "charts", "notes", "sections", "layouts"),
}


def validity(task: Task) -> dict | None:
    """rdocx / rpptx validate and the built-in validate.py (--original for an edit) on an Office output."""
    if task.out_kind not in ("docx", "pptx") or not Path(task.out).exists():
        return None
    cli = common.RDOCX if task.out_kind == "docx" else common.RPPTX
    skill = common.BUILTIN_DOCX if task.out_kind == "docx" else common.BUILTIN_PPTX
    r1 = common.run([cli, "validate", task.out], timeout=600)
    cmd = [common.PY, skill / "scripts" / "office" / "validate.py", task.out]
    if task.original:
        cmd += ["--original", task.original]
    r2 = common.run(cmd, env=ENV["builtin"], timeout=600)
    tail = lambda r: (r["stdout"] + r["stderr"]).strip().splitlines()[-3:]  # noqa: E731
    return {"rdocx_validate": r1["rc"] == 0, "builtin_validate": r2["rc"] == 0,
            "rdocx_validate_tail": None if r1["rc"] == 0 else tail(r1),
            "builtin_validate_tail": None if r2["rc"] == 0 else tail(r2)}


# ------------------------------------------------------------------------------------------ operations
def docx_tasks(f: dict, picks: dict) -> list[Task]:
    fid, src = f["id"], f["path"]
    o = WORK / "out" / fid
    R, PY, RPY = str(common.RDOCX), str(common.PY), str(common.R_PY)
    BD = str(common.BUILTIN_DOCX)
    tasks = []
    samples = picks["samples"]

    def text_check(task, stdout):
        return found_samples(stdout, samples)

    tasks += [
        Task("docx.read", f, "rdocx", "text --json", [R, "text", "--json", src], stdout_check=True, check=text_check),
        Task("docx.read", f, "builtin", "pandoc", [str(common.PANDOC), "-t", "markdown", src], stdout_check=True, check=text_check),
        Task("docx.read", f, "builtin", "xml (unzip + lxml)", ["bash", HERE / "docx_builtin_read.sh", src], primary=False,
             stdout_check=True, check=text_check),
        Task("docx.read", f, "pylib", "python-docx", [PY, HERE / "docx_pylib_read.py", src], stdout_check=True, check=text_check),
    ]

    word, n = picks["word"], picks["count"]
    if word:
        new = oracle.replacement_for(word)
        new_before = oracle.docx_count(src, new)

        def replace_check(task, stdout):
            left, got = oracle.docx_count(task.out, word), oracle.docx_count(task.out, new)
            return {"ok": left == 0 and got == new_before + n, "word": word, "expected": n, "left": left,
                    "replaced": got - new_before}

        out = str(o / "replace.docx")
        for op, validate in (("docx.replace", False), ("docx.replace+validate", True)):
            v = "+validate" if validate else ""
            tasks += [
                Task(op, f, "rdocx", "docx_ops replace" + v,
                     ["bash", "-c", 'python -m docx_ops replace "$1" "$2" --edit "$3" "$4" "$5" >/dev/null'
                      + (' && rdocx validate "$2" >/dev/null' if validate else ""), "_", src, out, word, new, str(n)],
                     out=out, out_kind="docx", original=src, check=replace_check),
                Task(op, f, "builtin", "unzip/merge_runs/lxml/zip" + v,
                     ["bash", HERE / "docx_builtin_replace.sh", src, out, word, new, str(n)] + (["--validate"] if validate else []),
                     out=out, out_kind="docx", original=src, check=replace_check),
            ]
            if not validate:
                tasks += [
                    Task(op, f, "rdocx", "rdocx replace --expect", [R, "replace", src, "-p", word, "-v", new, "--expect", str(n), "-o", out],
                         out=out, out_kind="docx", original=src, check=replace_check, primary=False),
                    Task(op, f, "pylib", "python-docx", [PY, HERE / "docx_pylib_replace.py", src, out, word, new, str(n)],
                         out=out, out_kind="docx", original=src, check=replace_check),
                ]
            else:
                tasks.append(Task(op, f, "pylib", "-", unsupported="python-docx has no validator; its workflow has no validation step"))

    phrase = picks["phrase"]
    if phrase:
        def comment_check(task, stdout):
            return oracle.comment_check(task.out, COMMENT_TEXT, phrase) | {"phrase": phrase}

        out = str(o / "comment.docx")
        for op, validate in (("docx.comment", False), ("docx.comment+validate", True)):
            v = "+validate" if validate else ""
            tasks += [
                Task(op, f, "rdocx", "docx_ops comment" + v,
                     ["bash", "-c", 'python -m docx_ops comment "$1" "$2" --anchor "$3" --text "$4" --author "Speed Bench" '
                      '--initials SB --date 2026-10-31T09:00:00Z >/dev/null' + (' && rdocx validate "$2" >/dev/null' if validate else ""),
                      "_", src, out, phrase, COMMENT_TEXT], out=out, out_kind="docx", original=src, check=comment_check),
                Task(op, f, "builtin", "unzip/merge_runs/comment.py/markers/zip" + v,
                     ["bash", HERE / "docx_builtin_comment.sh", src, out, phrase] + (["--validate"] if validate else []),
                     out=out, out_kind="docx", original=src, check=comment_check),
            ]
            if not validate:
                tasks.append(Task(op, f, "rdocx", "rdocx comment add --anchor",
                                  [R, "comment", "add", src, "--anchor", phrase, "--author", "Speed Bench", "--initials", "SB",
                                   "--text", COMMENT_TEXT, "--date", "2026-10-31T09:00:00Z", "-o", out],
                                  out=out, out_kind="docx", original=src, check=comment_check, primary=False))
            tasks.append(Task(op, f, "pylib", "python-docx add_comment", [PY, HERE / "docx_pylib_comment.py", src, out, phrase],
                              out=out, out_kind="docx", original=src, check=comment_check) if not validate else
                         Task(op, f, "pylib", "-", unsupported="python-docx has no validator; its workflow has no validation step"))
    else:
        for stack in STACKS:
            tasks.append(Task("docx.comment", f, stack, "-", unsupported="no phrase meets the selection rules (oracle.choose_docx_phrase)"))

    if picks["revisions"]:
        def accept_check(task, stdout):
            left = oracle.revision_marks(task.out)
            return {"ok": left == 0, "revision_marks_left": left, "revision_marks_before": picks["revisions"]}

        out = str(o / "accepted.docx")
        tasks += [
            Task("docx.accept", f, "rdocx", "rdocx revision accept", [R, "revision", "accept", src, "-o", out],
                 out=out, out_kind="docx", check=accept_check),
            Task("docx.accept", f, "builtin", "accept_changes.py (LibreOffice)", [PY, f"{BD}/scripts/accept_changes.py", src, out],
                 out=out, out_kind="docx", check=accept_check, lo=True),
            Task("docx.accept", f, "pylib", "-", unsupported="python-docx has no API for tracked changes (an lxml rewrite is "
                 "possible: features/docx R06)"),
        ]

    pages = f["pages"] or {}

    def render_check(expected):
        def check(task, stdout):
            n_png = pngs(task.out)
            exp = expected() if callable(expected) else expected
            return {"ok": n_png == exp and bool(exp), "pngs": n_png, "expected": exp}
        return check

    def rdocx_cli_pages():
        """Pages of rdocx's own PDF of the file (convert --font-dir): what rdocx render draws. rdocx's Python
        page_count(font_dir=) (the manifest's pages.rdocx) can differ by a page (NOTES.md, finding 9)."""
        if fid not in CLI_PAGES:
            pdf = WORK / "pages" / f"{fid}.pdf"
            pdf.parent.mkdir(parents=True, exist_ok=True)
            common.run([R, "convert", src, "--to", "pdf", "-o", pdf, "--font-dir", FONTS, "--force"], timeout=600)
            CLI_PAGES[fid] = oracle.pdf_pages(str(pdf)) if pdf.exists() else None
        return CLI_PAGES[fid]

    for op, first in (("docx.render", False), ("docx.render-page1", True)):
        exp_r = 1 if first else rdocx_cli_pages
        exp_lo = 1 if first else (pages.get("libreoffice") or 0)
        out_r, out_b = str(o / f"{op}-rdocx"), str(o / f"{op}-builtin")
        tasks += [
            Task(op, f, "rdocx", "rdocx render", [R, "render", src, "-o", out_r, "--dpi", "100", "--font-dir", FONTS]
                 + (["--pages", "1"] if first else []), out=out_r, out_kind="dir", check=render_check(exp_r)),
            Task(op, f, "builtin", "soffice.py + pdftoppm", ["bash", HERE / "lo_render.sh", BD, src, out_b] + (["first"] if first else []),
                 out=out_b, out_kind="dir", check=render_check(exp_lo), lo=True),
            Task(op, f, "pylib", "-", unsupported="python-docx has no renderer"),
        ]

    def pdf_check(expected):
        def check(task, stdout):
            pdf = task.out if task.out.endswith(".pdf") else str(Path(task.out) / (Path(src).stem + ".pdf"))
            n_pages = oracle.pdf_pages(pdf) if Path(pdf).exists() else None
            return {"ok": bool(n_pages), "pages": n_pages, "manifest_pages": expected,
                    "pages_match_manifest": n_pages == expected}
        return check

    out_r, out_b = str(o / "document.pdf"), str(o / "pdf-builtin")
    tasks += [
        Task("docx.pdf", f, "rdocx", "rdocx convert --to pdf", [R, "convert", src, "--to", "pdf", "-o", out_r, "--font-dir", FONTS],
             out=out_r, out_kind="file", check=pdf_check(pages.get("rdocx"))),
        Task("docx.pdf", f, "builtin", "soffice.py --convert-to pdf",
             [PY, f"{BD}/scripts/office/soffice.py", "--headless", "--convert-to", "pdf", "--outdir", out_b, src],
             out=out_b, out_kind="dir", check=pdf_check(pages.get("libreoffice")), lo=True, mkdir=True),
        Task("docx.pdf", f, "pylib", "-", unsupported="python-docx has no renderer"),
    ]

    def md_check(task, stdout):
        return found_samples(Path(task.out).read_text(errors="replace"), samples) if Path(task.out).exists() else {"ok": False}

    out_r, out_b = str(o / "rdocx.md"), str(o / "pandoc.md")
    tasks += [
        Task("docx.markdown", f, "rdocx", "rdocx convert --to md", [R, "convert", src, "--to", "md", "-o", out_r],
             out=out_r, out_kind="file", check=md_check),
        Task("docx.markdown", f, "builtin", "pandoc -t markdown", [str(common.PANDOC), "-t", "markdown", src, "-o", out_b],
             out=out_b, out_kind="file", check=md_check),
        Task("docx.markdown", f, "pylib", "-", unsupported="no converter in python-docx (features/docx D14 writes one, 47 lines)"),
    ]

    def validate_check(task, stdout):
        rc_file = Path(task.out + ".rc")
        rc = int(rc_file.read_text()) if rc_file.exists() else None
        text = Path(task.out).read_text(errors="replace") if Path(task.out).exists() else ""
        return {"ok": rc is not None, "verdict": "valid" if rc == 0 else "invalid", "tail": text.strip().splitlines()[-2:]}

    out_r, out_b = str(o / "validate-rdocx.txt"), str(o / "validate-builtin.txt")
    tasks += [
        Task("docx.validate", f, "rdocx", "rdocx validate", ["bash", HERE / "rc0.sh", out_r, R, "validate", src],
             out=out_r, out_kind="file", check=validate_check),
        Task("docx.validate", f, "builtin", "validate.py", ["bash", HERE / "rc0.sh", out_b, PY, f"{BD}/scripts/office/validate.py", src],
             out=out_b, out_kind="file", check=validate_check),
        Task("docx.validate", f, "pylib", "-", unsupported="python-docx has no validator"),
    ]
    return tasks


def pptx_tasks(f: dict, picks: dict) -> list[Task]:
    fid, src = f["id"], f["path"]
    o = WORK / "out" / fid
    RP, PY = str(common.RPPTX), str(common.PY)
    BP = str(common.BUILTIN_PPTX)
    samples, slides = picks["samples"], f["slides"]
    tasks = []

    def text_check(task, stdout):
        return found_samples(stdout, samples)

    tasks += [
        Task("pptx.read", f, "rdocx", "rpptx text --notes", [RP, "text", "--notes", src], stdout_check=True, check=text_check),
        Task("pptx.read", f, "builtin", "slide_text.py", [PY, f"{BP}/scripts/slide_text.py", src], stdout_check=True, check=text_check),
        Task("pptx.read", f, "pylib", "python-pptx", [PY, HERE / "pptx_pylib_read.py", src], stdout_check=True, check=text_check),
    ]

    word, n = picks["word"], picks["count"]
    if word:
        new = oracle.replacement_for(word)
        new_before = oracle.pptx_count(src, new)

        def replace_check(task, stdout):
            left, got = oracle.pptx_count(task.out, word), oracle.pptx_count(task.out, new)
            return {"ok": left == 0 and got == new_before + n, "word": word, "expected": n, "left": left,
                    "replaced": got - new_before}

        out = str(o / "replace.pptx")
        for op, validate in (("pptx.replace", False), ("pptx.replace+validate", True)):
            v = "+validate" if validate else ""
            tasks += [
                Task(op, f, "rdocx", "rpptx replace --expect" + v,
                     ["bash", "-c", 'rpptx replace "$1" -p "$3" -v "$4" --expect "$5" -o "$2" >/dev/null'
                      + (' && rpptx validate "$2" >/dev/null' if validate else ""), "_", src, out, word, new, str(n)],
                     out=out, out_kind="pptx", original=src, check=replace_check),
                Task(op, f, "builtin", "zipfile/lxml/zip" + v,
                     ["bash", HERE / "pptx_builtin_replace.sh", src, out, word, new, str(n)] + (["--validate"] if validate else []),
                     out=out, out_kind="pptx", original=src, check=replace_check),
                Task(op, f, "pylib", "python-pptx", [PY, HERE / "pptx_pylib_replace.py", src, out, word, new, str(n)],
                     out=out, out_kind="pptx", original=src, check=replace_check) if not validate else
                Task(op, f, "pylib", "-", unsupported="python-pptx has no validator; its workflow has no validation step"),
            ]

    k = oracle.middle_slide(src)
    parts = oracle.slide_parts(src)
    slide_file = Path(parts[k - 1][0]).name
    before = oracle.slide_texts(src)

    def duplicate_check(task, stdout):
        after = oracle.slide_texts(task.out)
        ok = len(after) == len(before) + 1 and after[k][0] == before[k - 1][0] and after[k - 1][0] == before[k - 1][0]
        return {"ok": ok, "slide": k, "slides_after": len(after), "notes_copied": len(after) > k and after[k][1] == before[k - 1][1]}

    out = str(o / "duplicate.pptx")
    tasks += [
        Task("pptx.duplicate", f, "rdocx", "rpptx slide duplicate", [RP, "slide", "duplicate", src, str(k), "-o", out],
             out=out, out_kind="pptx", original=src, check=duplicate_check),
        Task("pptx.duplicate", f, "builtin", "add_slide.py -o", [PY, f"{BP}/scripts/add_slide.py", src, slide_file, "--after", slide_file, "-o", out],
             out=out, out_kind="pptx", original=src, check=duplicate_check),
        Task("pptx.duplicate", f, "pylib", "python-pptx + lxml copy", [PY, HERE / "pptx_pylib_duplicate.py", src, out, str(k)],
             out=out, out_kind="pptx", original=src, check=duplicate_check),
    ]

    def fit_check(task, stdout):
        rc_file = Path(task.out + ".rc")
        try:
            rec = json.loads(Path(task.out).read_text())
            frames = rec.get("frames", [])
            return {"ok": True, "rc": int(rc_file.read_text()), "frames": len(frames),
                    "overflowing": sum(1 for fr in frames if fr.get("overflow"))}
        except (OSError, json.JSONDecodeError, ValueError) as e:
            return {"ok": False, "error": str(e)}

    out = str(o / "fit.json")
    tasks += [
        Task("pptx.fit", f, "rdocx", "rpptx fit --json", ["bash", HERE / "rc0.sh", out, RP, "fit", "--json", "--font-dir", FONTS, src],
             out=out, out_kind="file", check=fit_check),
        Task("pptx.fit", f, "rdocx", "pptx_ops.py overflow (text_layout)", ["python", "-m", "pptx_ops", "overflow", src],
             stdout_check=True, check=lambda task, stdout: {"ok": True, "overflowing": len(stdout.splitlines())}, primary=False),
        Task("pptx.fit", f, "builtin", "-", unsupported="no text-fit check: the skill's way is visual QA of rendered slides"),
        Task("pptx.fit", f, "pylib", "-", unsupported="python-pptx has no layout engine"),
    ]

    def render_check(expected):
        def check(task, stdout):
            n_png = pngs(task.out)
            return {"ok": n_png == expected and expected > 0, "pngs": n_png, "expected": expected}
        return check

    for op, first in (("pptx.render", False), ("pptx.render-slide1", True)):
        exp = 1 if first else slides
        out_r, out_b = str(o / f"{op}-rdocx"), str(o / f"{op}-builtin")
        tasks += [
            Task(op, f, "rdocx", "rpptx render", [RP, "render", src, "-o", out_r, "--dpi", "100", "--font-dir", FONTS]
                 + (["--slide", "1"] if first else []), out=out_r, out_kind="dir", check=render_check(exp)),
            Task(op, f, "builtin", "soffice.py + pdftoppm", ["bash", HERE / "lo_render.sh", BP, src, out_b] + (["first"] if first else []),
                 out=out_b, out_kind="dir", check=render_check(exp), lo=True),
            Task(op, f, "pylib", "-", unsupported="python-pptx has no renderer"),
        ]

    def thumb_check(task, stdout):
        folder = Path(task.out) if task.out_kind == "dir" else Path(task.out).parent
        imgs = sorted(p.name for p in folder.glob("*") if p.suffix in (".jpg", ".png"))
        return {"ok": bool(imgs), "images": len(imgs)}

    out_r, out_b = str(o / "thumb-rdocx" / "thumb.png"), str(o / "thumb-builtin")
    tasks += [
        Task("pptx.thumbnail", f, "rdocx", "rpptx thumbnail (slide 1 only, not a grid)",
             [RP, "thumbnail", src, "-o", out_r, "--font-dir", FONTS], out=out_r, out_kind="file", check=thumb_check,
             note="not equivalent: rpptx thumbnail renders slide 1 at 320 px; no grid of every slide"),
        Task("pptx.thumbnail", f, "builtin", "thumbnail.py (grid)", [PY, f"{BP}/scripts/thumbnail.py", src, str(Path(out_b) / "thumbs")],
             out=out_b, out_kind="dir", check=thumb_check, lo=True, mkdir=True),
        Task("pptx.thumbnail", f, "pylib", "-", unsupported="python-pptx has no renderer"),
    ]

    def pdf_check(task, stdout):
        pdf = task.out if task.out.endswith(".pdf") else str(Path(task.out) / (Path(src).stem + ".pdf"))
        n_pages = oracle.pdf_pages(pdf) if Path(pdf).exists() else None
        return {"ok": n_pages == slides, "pages": n_pages, "expected": slides}

    out_r, out_b = str(o / "deck.pdf"), str(o / "pdf-builtin")
    tasks += [
        Task("pptx.pdf", f, "rdocx", "rpptx convert --to pdf", [RP, "convert", src, "--to", "pdf", "-o", out_r, "--font-dir", FONTS],
             out=out_r, out_kind="file", check=pdf_check),
        Task("pptx.pdf", f, "builtin", "soffice.py --convert-to pdf",
             [PY, f"{BP}/scripts/office/soffice.py", "--headless", "--convert-to", "pdf", "--outdir", out_b, src],
             out=out_b, out_kind="dir", check=pdf_check, lo=True, mkdir=True),
        Task("pptx.pdf", f, "pylib", "-", unsupported="python-pptx has no renderer"),
    ]

    def validate_check(task, stdout):
        rc_file = Path(task.out + ".rc")
        rc = int(rc_file.read_text()) if rc_file.exists() else None
        text = Path(task.out).read_text(errors="replace") if Path(task.out).exists() else ""
        return {"ok": rc is not None, "verdict": "valid" if rc == 0 else "invalid", "tail": text.strip().splitlines()[-2:]}

    out_r, out_b = str(o / "validate-rpptx.txt"), str(o / "validate-builtin.txt")
    tasks += [
        Task("pptx.validate", f, "rdocx", "rpptx validate", ["bash", HERE / "rc0.sh", out_r, RP, "validate", src],
             out=out_r, out_kind="file", check=validate_check),
        Task("pptx.validate", f, "builtin", "validate.py", ["bash", HERE / "rc0.sh", out_b, PY, f"{BP}/scripts/office/validate.py", src],
             out=out_b, out_kind="file", check=validate_check),
        Task("pptx.validate", f, "pylib", "-", unsupported="python-pptx has no validator"),
    ]
    return tasks


def create_tasks(f: dict) -> list[Task]:
    fid, plan, kind = f["id"], f["path"], f["kind"]
    o = WORK / "out" / fid
    ref = stats(f["reference"])  # the corpus file: docx-js / pptxgenjs from the same plan

    def create_check(task, stdout):
        got = stats(task.out)
        keys = CREATE_KEYS[kind]
        diff = {k: {"got": got.get(k), "corpus": ref.get(k)} for k in keys if got.get(k) != ref.get(k)}
        return {"ok": "error" not in got, "counts": {k: got.get(k) for k in keys}, "differs_from_corpus": diff}

    out = str(o / f"created.{kind}")
    if kind == "docx":
        tasks = [
            Task("docx.create", f, "rdocx", "rdocx Python (chapters + insert_document)", [str(common.R_PY), HERE / "docx_create_rdocx.py", plan, out],
                 out=out, out_kind="docx", check=create_check),
            Task("docx.create", f, "builtin", "docx-js (corpus/gen_docx.js)",
                 [common.NODE, HERE.parent / "corpus" / "gen_docx.js", plan, out], out=out, out_kind="docx", check=create_check),
            Task("docx.create", f, "pylib", "python-docx (+ lxml)", [str(common.PY), HERE / "docx_create_pylib.py", plan, out],
                 out=out, out_kind="docx", check=create_check),
        ]
        if f["size_class"] == "S":
            tasks.append(Task("docx.create", f, "rdocx", "rdocx Python naive (one body, cell handles)",
                              [str(common.R_PY), HERE / "docx_create_rdocx.py", plan, out, "--naive"],
                              out=out, out_kind="docx", check=create_check, primary=False))
        return tasks
    apply_theme = common.BUILTIN_PPTX / "scripts" / "apply_theme.js"
    return [
        Task("pptx.create", f, "rdocx", "rpptx Python", [str(common.R_PY), HERE / "pptx_create_rdocx.py", plan, out],
             out=out, out_kind="pptx", check=create_check),
        Task("pptx.create", f, "builtin", "pptxgenjs + apply_theme.js (corpus/gen_pptx.js)",
             [common.NODE, HERE.parent / "corpus" / "gen_pptx.js", plan, out, apply_theme], out=out, out_kind="pptx", check=create_check),
        Task("pptx.create", f, "pylib", "python-pptx (corpus/gen_pptx_pythonpptx.py)",
             [str(common.PY), HERE.parent / "corpus" / "gen_pptx_pythonpptx.py", plan, out], out=out, out_kind="pptx", check=create_check),
    ]


# ------------------------------------------------------------------------------------------ timing
def clean(task: Task):
    if TMP.exists():
        shutil.rmtree(TMP, ignore_errors=True)
    TMP.mkdir(parents=True, exist_ok=True)
    if task.out:
        p = Path(task.out)
        for q in (p, Path(str(p) + ".rc")):
            if q.is_dir():
                shutil.rmtree(q)
            elif q.exists():
                q.unlink()
        if task.out_kind == "file" or task.out_kind in ("docx", "pptx"):
            p.parent.mkdir(parents=True, exist_ok=True)
        if task.mkdir:
            p.mkdir(parents=True, exist_ok=True)
        elif task.out_kind == "dir":
            p.parent.mkdir(parents=True, exist_ok=True)


def measure(task: Task, quick: bool) -> dict:
    env = ENV[task.stack]
    make = lambda i: task.argv  # noqa: E731
    setup = lambda i: clean(task)  # noqa: E731
    if quick:
        r = common.timeit(make, repeat=1, warmup=0, env=env, setup=setup, timeout=600)
        r["runs"], r["flags"] = 1, ["quick"]
        return r
    warm = common.timeit(make, repeat=1, warmup=0, env=env, setup=setup, timeout=600)
    if not warm["ok"]:
        return warm
    t = warm["median"]
    n = 1 if t > 60 else 3 if (task.lo or t > 20) else 5
    r = common.timeit(make, repeat=n, warmup=0, env=env, setup=setup, timeout=600)
    if r["ok"]:
        r["warmup"], r["runs"] = round(t, 4), n
        r["flags"] = (["single-run (warm-up over 60 s)"] if n == 1 else []) + (["3 runs (LibreOffice)"] if task.lo and n == 3 else []) \
            + (["3 runs (warm-up over 20 s)"] if n == 3 and not task.lo else [])
    return r


def loadavg():
    return [round(x, 2) for x in os.getloadavg()]


def run_group(tasks: list[Task], quick: bool, log) -> list[dict]:
    out = []
    before = loadavg()
    for task in tasks:
        f = task.file
        rec = {"op": task.op, "file": f["id"], "kind": f["kind"], "origin": f["origin"], "size_class": f["size_class"],
               "pages": f.get("pages"), "slides": f.get("slides"), "bytes": f["bytes"], "stack": task.stack,
               "variant": task.variant, "primary": task.primary}
        if task.note:
            rec["note"] = task.note
        if task.unsupported:
            rec.update(status="unsupported", reason=task.unsupported)
            out.append(rec)
            continue
        rec["command"] = " ".join(str(a).replace(str(common.BENCH) + "/", "").replace(str(common.BUILTIN_SKILLS_DIR), "$BUILTIN")
                                  .replace(str(common.R), "$R").replace(str(Path.home()), "~") for a in task.argv)
        t0 = time.perf_counter()
        r = measure(task, quick)
        if not r["ok"]:
            rec.update(status="error", rc=r.get("rc"), error=(r.get("error") or "").replace(str(common.BENCH) + "/", "")[-800:])
            log(f"  {task.key}: ERROR rc={r.get('rc')} {rec['error'][-200:]!r}")
            out.append(rec)
            continue
        stdout = common.run(task.argv, env=ENV[task.stack], timeout=600)["stdout"] if task.stdout_check else ""
        try:
            check = task.check(task, stdout) if task.check else {"ok": True}
        except Exception as e:  # a check that cannot read the output fails the output, it does not stop the run
            check = {"ok": False, "error": f"{type(e).__name__}: {e}"}
        rec.update(status="ok" if check.get("ok") else "fail", median=round(r["median"], 4), min=round(r["min"], 4),
                   max=round(r["max"], 4), samples=[round(s, 4) for s in r["samples"]], runs=r["runs"],
                   warmup=r.get("warmup"), flags=r.get("flags", []), check=check, valid=validity(task))
        if task.op.endswith(("render", "render-page1", "render-slide1")) and check.get("pngs"):
            rec["seconds_per_page"] = round(r["median"] / check["pngs"], 4)
        log(f"  {task.key}: {rec['status']} median {rec['median']:.3f}s ({r['runs']} runs) "
            f"{'' if check.get('ok') else json.dumps(check)[:200]} {time.perf_counter() - t0:.0f}s")
        out.append(rec)
    after = loadavg()
    for rec in out:
        rec["load_before"], rec["load_after"] = before, after
    return out


# ------------------------------------------------------------------------------------------ summary
def summarize(measurements: list[dict]) -> list[dict]:
    groups = {}
    for m in measurements:
        groups.setdefault((m["op"], m["file"]), []).append(m)
    rows = []
    for (op, fid), ms in groups.items():
        first = ms[0]
        row = {"op": op, "file": fid, "origin": first["origin"], "size_class": first["size_class"],
               "pages": first["pages"], "slides": first["slides"], "bytes": first["bytes"], "variants": {}}
        for m in ms:
            if m["status"] in ("ok", "fail"):
                row["variants"][f"{m['stack']}: {m['variant']}"] = {"median": m["median"], "status": m["status"]}
            if not m["primary"]:
                continue
            if m["status"] == "unsupported":
                row[m["stack"]] = None
                row[m["stack"] + "_status"] = "unsupported"
            elif m["status"] == "error":
                row[m["stack"]] = None
                row[m["stack"] + "_status"] = "error"
            else:
                row[m["stack"]] = m["median"]
                row[m["stack"] + "_status"] = m["status"]
        for stack in ("builtin", "pylib"):
            a, b = row.get(stack), row.get("rdocx")
            row[f"{stack}/rdocx"] = round(a / b, 3) if a and b else None
        rows.append(row)
    return rows


# ------------------------------------------------------------------------------------------ main
def wait_for_quiet(log, limit=4.0, poll=120, max_wait=1800) -> dict:
    start = loadavg()
    waited, polls = 0, [start]
    while os.getloadavg()[0] > limit and waited < max_wait:
        log(f"[speed] 1-minute load {os.getloadavg()[0]:.2f} > {limit}: waiting {poll}s ({waited}s so far)")
        time.sleep(poll)
        waited += poll
        polls.append(loadavg())
    return {"at_start": start, "waited_s": waited, "polls": polls, "started_at_load": loadavg(),
            "note": ("started at once" if waited == 0 else
                     f"waited {waited} s for the 1-minute load to drop under {limit}"
                     + (" (gave up after 30 minutes)" if os.getloadavg()[0] > limit else ""))}


def main():
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("--quick", action="store_true", help="1 run per command, no warm-up (development)")
    ap.add_argument("--only", help="comma-separated fnmatch patterns on op:file:stack:variant")
    ap.add_argument("--no-wait", action="store_true", help="do not wait for a quiet machine")
    args = ap.parse_args()
    patterns = args.only.split(",") if args.only else None
    log = lambda s: print(s, flush=True)  # noqa: E731
    WORK.mkdir(parents=True, exist_ok=True)
    full = not args.quick and not patterns
    load = {"limit_1min": 4.0}
    if not args.quick and not args.no_wait:
        load.update(wait_for_quiet(log))

    files = corpus_files()
    tasks: list[Task] = []
    picks_all = {}
    for f in files:
        if f["kind"] == "docx":
            word, count = oracle.choose_docx_word(f["path"])
            phrase, runs = oracle.choose_docx_phrase(f["path"])
            picks = {"word": word, "count": count, "phrase": phrase, "phrase_runs": runs,
                     "samples": oracle.docx_samples(f["path"]), "revisions": oracle.revision_marks(f["path"])}
            tasks += docx_tasks(f, picks)
        else:
            word, count = oracle.choose_pptx_word(f["path"])
            picks = {"word": word, "count": count, "samples": oracle.pptx_samples(f["path"]),
                     "duplicate_slide": oracle.middle_slide(f["path"])}
            tasks += pptx_tasks(f, picks)
        picks_all[f["id"]] = picks
    for f in plan_files():
        tasks += create_tasks(f)
    if patterns:
        tasks = [t for t in tasks if any(fnmatch.fnmatch(t.key, p) for p in patterns)]

    # group by (op, file), groups in a stable order: docx before pptx, operation, then file size
    op_order = {}
    for t in tasks:
        op_order.setdefault(t.op, len(op_order))
    groups = {}
    for t in tasks:
        groups.setdefault((t.op, t.file["id"]), []).append(t)
    keys = sorted(groups, key=lambda k: (op_order[k[0]],))
    log(f"[speed] {len(tasks)} commands in {len(keys)} groups ({'quick' if args.quick else 'full'})")
    t0 = time.perf_counter()
    measurements, group_loads = [], []
    for op, fid in keys:
        log(f"[speed] {op} {fid}  load {loadavg()}")
        ms = run_group(groups[(op, fid)], args.quick, log)
        measurements += ms
        if ms:
            group_loads.append({"op": op, "file": fid, "before": ms[0]["load_before"], "after": ms[0]["load_after"]})
    load["groups"] = group_loads
    load["at_end"] = loadavg()
    data = {
        "section": "speed",
        "versions": common.versions(),
        "mode": "quick" if args.quick else "full" if full else "partial",
        "elapsed_s": round(time.perf_counter() - t0, 1),
        "protocol": "common.timeit: a warm-up run, then the median of 5 runs; 3 for LibreOffice operations and when the "
                    "warm-up took over 20 s; 1 when it took over 60 s (flagged). Process start included.",
        "load": load,
        "inputs": picks_all,
        "measurements": measurements,
        "summary": summarize(measurements),
    }
    if full:
        path = common.save_results("speed", data)
    else:
        path = WORK / ("quick.json" if args.quick else "partial.json")
        path.write_text(json.dumps(data, indent=2, ensure_ascii=False) + "\n")
    log(f"[speed] wrote {path} in {data['elapsed_s']} s")


if __name__ == "__main__":
    main()
