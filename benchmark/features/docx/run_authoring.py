"""Run the authoring and editing tasks D01-D14 on the three stacks and write results/features_docx_authoring.json.

For each task folder `Dxx-slug/`: build the inputs once (`make_input.py WORKDIR`), run each stack's
implementation (`rdocx.py` or `rdocx.sh`, `builtin.sh`, `pylib.py`) as an agent would run it, time it (median
of 5 runs after a warm-up), check its output with the task's neutral oracle (`check.py OUTPUT`), and run the
three validity checks on every .docx output: `rdocx validate`, the built-in skill's `office/validate.py`
(`--original INPUT` for edits) and a LibreOffice PDF conversion through the skill's `office/soffice.py`.

Usage: .cache/venv/bin/python features/docx/run_authoring.py [--task D03 ...] [--quick]
(--quick: one run per stack and no validators, for development.)
"""
from __future__ import annotations

import argparse
import json
import re
import shutil
import sys
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parents[1]))
import common  # noqa: E402

STACKS = ("rdocx", "builtin", "pylib")

# id, slug, title, category, spec, inputs (made by make_input.py, in argument order), output name,
# extra arguments after the output, the input an edit is validated against (--original).
TASKS = [
    dict(id="D01", slug="create-report", title="Create a report", category="create",
         spec="New A4 document: Title, a table of contents, headings on 3 levels, a paragraph with bold and italic "
              "runs, a bulleted and a numbered list, a 5x4 table with a shaded header row, a PNG picture, a footer "
              "'Page X of Y' with PAGE and NUMPAGES fields.",
         inputs=["picture.png"], output="out.docx"),
    dict(id="D02", slug="read-all-text", title="Read every story", category="read",
         spec="Extract to JSON the text of body paragraphs and table cells in order, headers, footers, footnotes, "
              "text boxes and comments, each in its own key; every marker must land in its story and only there.",
         inputs=["input.docx"], output="out.json"),
    dict(id="D03", slug="counted-replace", title="Counted replacement", category="edit",
         spec="Replace 'Service Level Agreement' (5 occurrences: plain, split across runs with identical and with "
              "different formatting, in a table cell, in the header) by 'Support Contract', keeping the formatting "
              "around; with an expected count of 4 the implementation must refuse and write nothing.",
         inputs=["input.docx"], output="out.docx", extra=["5"], original="input.docx",
         refusal=["4"]),
    dict(id="D04", slug="fill-template", title="Fill a template", category="edit",
         spec="Fill {{client}} and {{date}} (placeholders split across runs) and repeat the table's item row "
              "({{name}}, {{qty}}, {{price}}) once per item of a 6-item list, keeping the row's formatting and the "
              "total row after the items.",
         inputs=["template.docx", "data.json"], output="out.docx", original="template.docx"),
    dict(id="D05", slug="table-edit", title="Edit a table", category="edit",
         spec="In an existing 4x5 table: insert a 'Supplier' column after the first one (bold header like its "
              "neighbours), merge the 'Total' cell with the empty cell to its right (one paragraph left, as Word "
              "merges), delete the 'Nuts' row, shade the 'Amount' header cell FFC000.",
         inputs=["input.docx"], output="out.docx", original="input.docx"),
    dict(id="D06", slug="page-setup", title="Page setup of two sections", category="edit",
         spec="In a 2-section document: section 2 landscape with 2 cm margins (section 1 unchanged), a header "
              "'Confidential' and a 'Page X of Y' footer showing in both sections.",
         inputs=["input.docx"], output="out.docx", original="input.docx"),
    dict(id="D07", slug="pictures", title="Insert and replace pictures", category="edit",
         spec="Insert new.png after the paragraph 'Figure placement anchor.' at 10 cm wide (height in proportion), "
              "and replace the existing picture's image by replacement.png keeping its displayed size.",
         inputs=["input.docx", "new.png", "replacement.png"], output="out.docx", original="input.docx"),
    dict(id="D08", slug="chart", title="Insert a native chart", category="edit",
         spec="Insert after the paragraph 'Chart anchor: sales by quarter.' an editable clustered column chart "
              "(embedded workbook) with series 2025 and 2026 over Q1-Q4.",
         inputs=["input.docx"], output="out.docx", original="input.docx"),
    dict(id="D09", slug="equation", title="Insert an equation", category="edit",
         spec="Insert after the paragraph 'The roots are given by:' a display equation (OMML) of the quadratic "
              "formula x = (-b +- sqrt(b^2 - 4ac)) / 2a, structured as the OMML schema requires (m:rad holds "
              "m:deg, hidden, then m:e).",
         inputs=["input.docx"], output="out.docx", original="input.docx"),
    dict(id="D10", slug="hyperlinks", title="Add hyperlinks", category="edit",
         spec="Append to the paragraph 'Further reading: ' an external link 'Project site' to "
              "https://example.com/project, then ' and ', then an internal link 'Appendix A' to a bookmark on the "
              "heading 'Appendix A', then '.'; both links look like links (Hyperlink style or underline).",
         inputs=["input.docx"], output="out.docx", original="input.docx"),
    dict(id="D11", slug="footnote", title="Add a footnote", category="edit",
         spec="Add the footnote 'Source: field survey 2026.' with its reference mark right after the sentence "
              "'The survey covered 120 sites.' in the middle of its paragraph (the document has no footnotes yet).",
         inputs=["input.docx"], output="out.docx", original="input.docx"),
    dict(id="D12", slug="styles-lists", title="Custom style and list restart", category="edit",
         spec="Create the paragraph style 'Body Note' based on Normal (Arial 10 pt, 6 pt before, 12 pt after) and "
              "apply it to the two paragraphs starting 'Note:'; make the second numbered list restart at 1.",
         inputs=["input.docx"], output="out.docx", original="input.docx"),
    dict(id="D13", slug="markdown-to-docx", title="Markdown to Word", category="convert",
         spec="Convert a Markdown file (3 heading levels, bulleted and numbered lists, a table, a fenced code "
              "block, bold and italic, a link) to .docx with Word's own structures (the link looks like one).",
         inputs=["notes.md"], output="out.docx"),
    dict(id="D14", slug="docx-to-markdown", title="Word to Markdown", category="convert",
         spec="Convert a Word document (headings, bulleted and numbered lists made by Word's numbering, a table, "
              "bold and italic) to Markdown.",
         inputs=["input.docx"], output="out.md"),
]

SOFFICE = common.BUILTIN_DOCX / "scripts" / "office" / "soffice.py"
VALIDATE = common.BUILTIN_DOCX / "scripts" / "office" / "validate.py"


def stack_env(stack: str) -> dict:
    env = {"rdocx": common.env_rdocx, "builtin": common.env_builtin, "pylib": common.env_pylib}[stack]()
    env["BUILTIN_DOCX"] = str(common.BUILTIN_DOCX)
    return env


def impl_files(folder: Path, stack: str) -> list[Path]:
    return sorted(p for p in folder.iterdir()
                  if p.is_file() and p.name.startswith((stack + ".", stack + "_")) and p.suffix != ".unsupported")


def entry_cmd(folder: Path, stack: str) -> list:
    if stack == "rdocx":
        if (folder / "rdocx.py").exists():
            # -P: do not put the script's folder first on sys.path, where rdocx.py would shadow the rdocx module.
            return [common.R_PY, "-P", folder / "rdocx.py"]
        return ["bash", folder / "rdocx.sh"]
    if stack == "builtin":
        return ["bash", folder / "builtin.sh"]
    return [common.PY, folder / "pylib.py"]


def loc(path: Path) -> int:
    n = 0
    for line in path.read_text().splitlines():
        s = line.strip()
        if not s or s.startswith("#") or s.startswith("//"):
            continue
        n += 1
    return n


def processes(files: list[Path]) -> int:
    """Separate commands an agent runs: 1 for a script it runs; for a shell script, each command line."""
    sh = [f for f in files if f.suffix == ".sh"]
    if not sh:
        return 1
    n = 0
    for line in sh[0].read_text().splitlines():
        s = line.strip()
        if not s or s.startswith("#") or s.startswith("set ") or s in ("fi", "done", "else", "}", "{"):
            continue
        if re.match(r"^[A-Za-z_][A-Za-z0-9_]*=", s) and "$(" not in s.split("=", 1)[1][:3]:
            continue  # a variable the agent would just type into the next command
        if re.match(r"^[A-Za-z_][A-Za-z0-9_]*=\$\(mktemp", s):
            continue  # a scratch folder
        n += 1
    return n


TOOL_PATTERNS = [
    (r"\$R/rdocx |rdocx (convert|text|replace|validate)", "rdocx CLI"),
    (r"^import rdocx|^\s*import rdocx", "rdocx (Python)"),
    (r"import docx_ops", "rdocx (Python)"),
    (r"import docx_ops", "docx_ops"),
    (r"docx_ops\.cli\(", "rdocx CLI"),
    (r"require\(\"docx\"\)", "docx-js (node)"),
    (r"\bpandoc\b", "pandoc"),
    (r"\bunzip\b", "unzip"),
    (r"\bzip -", "zip"),
    (r"merge_runs\.py", "merge_runs.py"),
    (r"comment\.py", "comment.py"),
    (r"soffice\.py", "soffice.py"),
    (r"^from docx|^import docx\b", "python-docx"),
    (r"^from pptx|^import pptx\b", "python-pptx"),
    (r"lxml|OxmlElement|parse_xml", "lxml"),
    (r"defusedxml", "defusedxml"),
    (r"openpyxl", "openpyxl"),
    (r"^import re$|\bre\.(sub|search|findall|finditer)", "regex"),
]

RAW_XML = re.compile(r"<\/?(w|m|c|a|wp|pic|r|mc|v|o|w14|wps|cx):[A-Za-z]|\bqn\(|OxmlElement|\{http://schemas\.openxml"
                     r"|replace_xml\(|['\"](w|m|c):[a-zA-Z]+['\"]|\._(tr|tc|tbl|p|r|element)\b")


def tools(files: list[Path]) -> list[str]:
    out = []
    for f in files:
        text = f.read_text()
        if f.suffix == ".py" and f.name.startswith("builtin"):
            out.append("python")
        for pat, name in TOOL_PATTERNS:
            if re.search(pat, text, re.M) and name not in out:
                out.append(name)
    return out


def raw_xml(files: list[Path]) -> bool:
    return any(RAW_XML.search(f.read_text()) for f in files)


def validators(out: Path, original: Path | None, workdir: Path) -> dict:
    env = common.env_builtin()
    r = common.run([common.RDOCX, "validate", out], env=common.env_rdocx(), timeout=300)
    cmd = [common.PY, VALIDATE, out] + (["--original", original] if original else [])
    b = common.run(cmd, env=env, timeout=300)
    pdf_dir = workdir / "pdf"
    pdf_dir.mkdir(exist_ok=True)
    s = common.run([common.PY, SOFFICE, "--headless", "--convert-to", "pdf", "--outdir", pdf_dir, out], env=env,
                   timeout=300)
    pdf = pdf_dir / (out.stem + ".pdf")
    lo_ok = s["rc"] == 0 and pdf.exists() and pdf.stat().st_size > 1000
    detail = {}
    if r["rc"] != 0:
        detail["rdocx_validate"] = (r["stdout"] + r["stderr"]).strip()[-600:]
    if b["rc"] != 0:
        detail["builtin_validate"] = (b["stdout"] + b["stderr"]).strip()[-600:]
    if not lo_ok:
        detail["libreoffice_opens"] = (s["stdout"] + s["stderr"]).strip()[-300:]
    return {"rdocx_validate": r["rc"] == 0, "builtin_validate": b["rc"] == 0, "libreoffice_opens": lo_ok}, detail


def run_stack(task: dict, stack: str, work: Path, quick: bool) -> dict:
    folder = HERE / f"{task['id']}-{task['slug']}"
    files = impl_files(folder, stack)
    entry = {"rdocx": ("rdocx.py", "rdocx.sh"), "builtin": ("builtin.sh",), "pylib": ("pylib.py",)}[stack]
    if not any((folder / e).exists() for e in entry):
        unsupported = folder / f"{stack}.unsupported"
        reason = unsupported.read_text().strip() if unsupported.exists() else "no implementation"
        return {"status": "unsupported", "loc": 0, "files": [], "tools": [], "processes": 0, "raw_xml": False,
                "seconds": None, "valid": None, "notes": reason}, None
    sdir = work / stack
    sdir.mkdir()
    inputs = [work / "in" / name for name in task["inputs"]]
    ext = Path(task["output"]).suffix
    env = stack_env(stack)
    base = entry_cmd(folder, stack)

    def cmd(i):
        return base + inputs + [sdir / f"run{i}{ext}"] + task.get("extra", [])

    first = common.run(cmd(0), env=env, cwd=sdir)
    result = {"loc": sum(loc(f) for f in files), "files": [f.name for f in files], "tools": tools(files),
              "processes": processes(files), "raw_xml": raw_xml(files)}
    notes = []
    if ext != ".docx" and result["raw_xml"]:
        # raw_xml counts OOXML written or edited by hand; a reader that parses the parts itself is noted only
        result["raw_xml"] = False
        notes.append("parses the OOXML parts by hand (no write)")
    if first["rc"] != 0:
        result.update(status="fail", seconds=None, valid=None,
                      notes=f"exit {first['rc']}: {(first['stderr'] or first['stdout']).strip()[-800:]}")
        return result, None
    out = sdir / f"run0{ext}"
    chk = common.run([common.PY, folder / "check.py", out], env=common.env_pylib())
    status = "pass" if chk["rc"] == 0 else "fail"
    oracle = [line for line in chk["stdout"].splitlines() if line.startswith(("FAIL", "info"))]
    if chk["rc"] not in (0, 1):
        oracle.append((chk["stderr"] or "").strip()[-600:])
    notes += oracle
    if task.get("refusal"):
        bad = sdir / f"refusal{ext}"
        ref = common.run(base + inputs + [bad] + task["refusal"], env=env, cwd=sdir)
        refused = ref["rc"] != 0 and not bad.exists()
        notes.append(f"refusal with expected count {task['refusal'][0]}: "
                     + ("refused, nothing written" if refused else f"NOT refused (exit {ref['rc']}, "
                        f"output {'written' if bad.exists() else 'absent'})"))
        if not refused:
            status = "fail"
    if quick:
        seconds = first["seconds"]
    else:
        # fresh output names: rdocx's CLI refuses to write over an existing file
        t = common.timeit(lambda i: cmd(f"t{i}"), repeat=5, warmup=1, env=env, cwd=sdir)
        if not t["ok"]:
            notes.append(f"timing run failed: {t['error'][-300:]}")
        seconds = t["median"] if t["ok"] else None
    pending = None
    if ext == ".docx" and not quick:
        original = work / "in" / task["original"] if task.get("original") else None
        pending = (out, original, sdir)
    result.update(status=status, seconds=round(seconds, 3) if seconds else seconds, valid=None,
                  notes="; ".join(notes))
    return result, pending


def run_task(task: dict, quick: bool):
    """Inputs once, then each stack in turn (timings never overlap); validators are returned as pending jobs."""
    folder = HERE / f"{task['id']}-{task['slug']}"
    work = common.fresh_dir("features_docx_authoring", task["id"])
    (work / "in").mkdir()
    if (folder / "make_input.py").exists():
        common.run([common.PY, folder / "make_input.py", work / "in"], env=common.env_pylib(), check=True)
    stacks, jobs = {}, []
    for s in STACKS:
        res, pending = run_stack(task, s, work, quick)
        stacks[s] = res
        if pending:
            jobs.append((res, pending))
    return {"id": task["id"], "title": task["title"], "category": task["category"], "spec": task["spec"],
            "stacks": stacks}, jobs


def apply_validators(jobs):
    def one(job):
        res, (out, original, sdir) = job
        valid, detail = validators(out, original, sdir)
        res["valid"] = valid
        extra = "; ".join(f"{k}: {v}" for k, v in detail.items())
        res["notes"] = "; ".join(x for x in (res["notes"], extra) if x)

    with ThreadPoolExecutor(6) as ex:
        list(ex.map(one, jobs))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--task", action="append", help="run only these task ids")
    ap.add_argument("--quick", action="store_true", help="one run per stack, no validators; results not saved")
    a = ap.parse_args()
    tasks = [t for t in TASKS if not a.task or t["id"] in a.task]
    out, jobs = [], []
    for task in tasks:
        res, task_jobs = run_task(task, a.quick)
        out.append(res)
        jobs += task_jobs
        line = "  ".join(f"{s}={r['status']}({r['loc']} loc{', xml' if r['raw_xml'] else ''})"
                         for s, r in res["stacks"].items())
        print(f"{task['id']} {task['slug']}: {line}", flush=True)
        for s, r in res["stacks"].items():
            if r["notes"] and (r["status"] != "pass" or a.quick):
                print(f"    {s}: {r['notes'][:600]}")
    if a.quick:
        return
    apply_validators(jobs)
    for t in out:
        for s, r in t["stacks"].items():
            if r["valid"] and not all(r["valid"].values()):
                print(f"{t['id']} {s}: valid {r['valid']}")
    if a.task:
        # a partial run updates those tasks in the saved results
        path = common.RESULTS / "features_docx_authoring.json"
        saved = json.loads(path.read_text()) if path.exists() else {"tasks": []}
        keep = {t["id"]: t for t in saved.get("tasks", [])}
        keep.update({t["id"]: t for t in out})
        out = [keep[t["id"]] for t in TASKS if t["id"] in keep]
    data = {"section": "features_docx_authoring", "versions": common.versions(), "tasks": out}
    print(common.save_results("features_docx_authoring", data))


if __name__ == "__main__":
    main()
