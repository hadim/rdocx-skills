"""Run the review, collaboration, layout and rendering tasks R01-R15 for the three stacks and write
results/features_docx_review.json.

For each task folder `Rxx-slug/`: `make_input.py DIR` writes the inputs; each stack's entry file
(`rdocx.sh|py`, `builtin.sh|py`, `pylib.sh|py`) runs with that stack's environment and takes the input
path(s) then the output path; `check.py OUTPUT` is the neutral oracle (exit 0 = pass). A `# bench: {...}`
JSON comment in each entry file declares the tools, the processes one run starts, whether it writes raw
OOXML, the other files it needs, and, for a stack that cannot do the task, `"unsupported"` with the reason.

Usage: .cache/venv/bin/python features/docx/run_review.py [--tasks R01,R04] [--repeat 5] [--jobs 4]
"""
from __future__ import annotations

import argparse
import ast
import json
import re
import shutil
import sys
import threading
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parents[1]))
import common  # noqa: E402

STACKS = ("rdocx", "builtin", "pylib")
TITLES = {
    "R01": ("Comment on a phrase split across runs", "review"),
    "R02": ("Reply to a comment and resolve the thread", "review"),
    "R03": ("List comments with their anchored text", "review"),
    "R04": ("Replace a word as a tracked change", "review"),
    "R05": ("List tracked changes", "review"),
    "R06": ("Accept all tracked changes", "review"),
    "R07": ("Reject the changes of one author", "review"),
    "R08": ("Word-level redline of two versions", "review"),
    "R09": ("Update a stale table of contents", "layout"),
    "R10": ("Update PAGE, NUMPAGES and DATE fields", "layout"),
    "R11": ("Render pages 1-3 to PNG (accepted and tracked views)", "render"),
    "R12": ("Page number of each heading", "layout"),
    "R13": ("Convert to PDF (accepted and tracked views)", "convert"),
    "R14": ("Which paragraphs differ between two versions", "review"),
    "R15": ("Append a document with its own styles, picture and numbering", "edit"),
}
# inputs each implementation takes (before the output), the output name, the author for validate --author
IO = {
    "R08": (["v1.docx", "v2.docx"], "out.docx"),
    "R14": (["v1.docx", "v2.docx"], "out.json"),
    "R15": (["a.docx", "b.docx"], "out.docx"),
    "R03": (["input.docx"], "out.json"),
    "R05": (["input.docx"], "out.json"),
    "R12": (["input.docx"], "out.json"),
    "R11": (["input.docx"], "out"),
    "R13": (["input.docx"], "out"),
}
REDLINE_AUTHOR = {"R04": "Dana Reviewer", "R08": "Dana Reviewer"}
LO_LOCK = threading.Lock()  # accept_changes.py uses one fixed LibreOffice profile: never two at once


def meta_of(path: Path) -> dict:
    for line in path.read_text().splitlines()[:5]:
        m = re.match(r"#\s*bench:\s*(\{.*\})\s*$", line)
        if m:
            return json.loads(m.group(1))
    return {}


def loc(path: Path) -> int:
    """Non-blank, non-comment lines; docstrings count as comments in Python."""
    lines = path.read_text().splitlines()
    skip = set()
    if path.suffix == ".py":
        for node in ast.walk(ast.parse("\n".join(lines))):
            if isinstance(node, ast.Expr) and isinstance(getattr(node, "value", None), ast.Constant) \
                    and isinstance(node.value.value, str):
                skip.update(range(node.lineno, node.end_lineno + 1))
    return sum(1 for i, s in enumerate(lines, 1) if s.strip() and not s.strip().startswith("#") and i not in skip)


def entry(task_dir: Path, stack: str) -> Path | None:
    for ext in (".sh", ".py"):
        p = task_dir / f"{stack}{ext}"
        if p.exists():
            return p
    return None


def env_for(stack: str) -> dict:
    env = {"rdocx": common.env_rdocx, "builtin": common.env_builtin, "pylib": common.env_pylib}[stack]()
    env["BUILTIN_DOCX"] = str(common.BUILTIN_DOCX)
    env["FONT_DIR"] = str(common.FONTS_DIR)
    if stack == "rdocx":
        env["PATH"] = f"{common.R}:{env['PATH']}"
    return env


def command(stack: str, path: Path, args: list) -> list:
    if path.suffix == ".sh":
        return ["bash", path, *args]
    py = common.R_PY if stack == "rdocx" else common.PY
    return [py, "-P", path, *args]  # -P: the entry rdocx.py must not shadow the rdocx module


def validators(out: Path, original: Path, author: str | None, work: Path) -> dict:
    """rdocx validate, the built-in skill's validate.py (--original, --author for redlines) and LibreOffice
    opening the file (soffice.py converting it to PDF)."""
    r = common.run([common.RDOCX, "validate", out], timeout=300)
    cmd = [common.PY, common.BUILTIN_DOCX / "scripts/office/validate.py", out, "--original", original]
    if author:
        cmd += ["--author", author]
    b = common.run(cmd, env=common.env_builtin(), timeout=300)
    pdf_dir = work / "lo-open"
    pdf_dir.mkdir(exist_ok=True)
    lo = common.run([common.PY, common.BUILTIN_DOCX / "scripts/office/soffice.py", "--headless", "--convert-to",
                     "pdf", "--outdir", pdf_dir, out], env=common.env_builtin(), timeout=300)
    pdf = pdf_dir / (out.stem + ".pdf")
    return {
        "rdocx_validate": r["rc"] == 0,
        "builtin_validate": b["rc"] == 0,
        "libreoffice_opens": lo["rc"] == 0 and pdf.exists() and pdf.stat().st_size > 0,
        "_messages": {"rdocx": (r["stdout"] + r["stderr"]).strip()[-600:],
                      "builtin": (b["stdout"] + b["stderr"]).strip()[-900:]},
    }


def run_stack(task_dir: Path, tid: str, stack: str, inputs: list, out_name: str, work: Path) -> dict:
    path = entry(task_dir, stack)
    if path is None:
        return {"status": "unsupported", "loc": 0, "files": [], "tools": [], "processes": 0, "raw_xml": False,
                "seconds": None, "notes": "no implementation"}
    meta = meta_of(path)
    files = meta.get("files", [path.name])
    rec = {
        "loc": sum(loc(task_dir / f) for f in files),
        "files": files,
        "tools": meta.get("tools", []),
        "processes": meta.get("processes", 1),
        "raw_xml": meta.get("raw_xml", False),
    }
    if "unsupported" in meta and not meta.get("run"):
        return {"status": "unsupported", **rec, "loc": 0, "seconds": None, "notes": meta["unsupported"]}
    env = env_for(stack)
    sdir = work / stack
    sdir.mkdir(parents=True, exist_ok=True)
    out = sdir / out_name
    lock = LO_LOCK if meta.get("lo_profile") else None
    if lock:
        lock.acquire()
    try:
        res = common.run(command(stack, path, [*inputs, out]), env=env, cwd=sdir, timeout=600)
    finally:
        if lock:
            lock.release()
    rec["seconds"] = round(res["seconds"], 3)  # replaced by the median of the timed runs (time_stack)
    if res["rc"] != 0:
        rec["run"] = {"rc": res["rc"], "stdout": res["stdout"][-800:], "stderr": res["stderr"][-1500:]}
    chk = common.run([common.PY, task_dir / "check.py", out], env=common.env_pylib(), timeout=600) if res["rc"] == 0 and out.exists() \
        else None
    try:
        rec["check"] = json.loads(chk["stdout"].strip().splitlines()[-1]) if chk else None
    except (json.JSONDecodeError, IndexError):
        rec["check"] = {"ok": False, "error": (chk["stdout"] + chk["stderr"])[-1500:]}
    if meta.get("unsupported"):
        rec["status"] = "unsupported"
    elif res["rc"] != 0 or not out.exists():
        rec["status"] = "fail"
    else:
        rec["status"] = "pass" if rec["check"] and rec["check"].get("ok") else "fail"
    if out_name.endswith(".docx") and out.exists():
        v = validators(out, Path(inputs[0]), REDLINE_AUTHOR.get(tid), sdir)
        rec["validator_messages"] = v.pop("_messages")
        rec["valid"] = v
    notes = [meta.get("notes", meta.get("unsupported", ""))]
    failed = [k for k, v in ((rec.get("check") or {}).get("checks") or {}).items() if not v]
    if failed:
        notes.append("oracle checks failed: " + ", ".join(failed))
    rec["notes"] = "; ".join(n for n in notes if n)
    order = ("status", "loc", "files", "tools", "processes", "raw_xml", "seconds", "valid", "notes")
    return {**{k: rec[k] for k in order if k in rec}, **{k: v for k, v in rec.items() if k not in order}}


def time_stack(task_dir: Path, stack: str, inputs: list, out_name: str, work: Path, repeat: int):
    """Median wall time of `repeat` runs after one warm-up (3 for LibreOffice), each into a fresh output, run
    one task at a time so that runs do not compete for the machine."""
    path = entry(task_dir, stack)
    meta = meta_of(path)
    env, sdir = env_for(stack), work / stack
    n = 3 if meta.get("libreoffice") else repeat

    def make(i):
        o = sdir / f"t{i}" / out_name
        if o.parent.exists():
            shutil.rmtree(o.parent)
        o.parent.mkdir(parents=True)
        return command(stack, path, [*inputs, o])

    timing = common.timeit(make, repeat=n, warmup=1, env=env, cwd=sdir, timeout=600)
    for i in range(n + 1):
        shutil.rmtree(sdir / f"t{i}", ignore_errors=True)
    if not timing.get("ok"):
        return None
    return {k: (round(v, 3) if isinstance(v, float) else [round(x, 3) for x in v])
            for k, v in timing.items() if k in ("median", "min", "max", "samples")}


def run_task(task_dir: Path, stacks=STACKS) -> dict:
    tid = task_dir.name[:3]
    work = common.fresh_dir("features_docx_review", task_dir.name)
    in_dir = work / "input"
    common.run([common.PY, task_dir / "make_input.py", in_dir], env=common.env_pylib(), check=True, timeout=300)
    names, out_name = IO.get(tid, (["input.docx"], "out.docx"))
    inputs = [in_dir / n for n in names]
    title, category = TITLES[tid]
    rec = {"id": tid, "slug": task_dir.name, "title": title, "category": category,
           "spec": (task_dir / "task.md").read_text().strip(), "stacks": {},
           "_run": (inputs, out_name, work)}
    for stack in stacks:
        rec["stacks"][stack] = run_stack(task_dir, tid, stack, inputs, out_name, work)
        s = rec["stacks"][stack]
        print(f"{tid} {stack:8s} {s['status']:12s} loc={s['loc']:4d} {s.get('seconds')}s "
              f"{json.dumps(s.get('valid', {}))} {json.dumps(s.get('check', {}) and s['check'].get('checks', {}))[:400]}",
              flush=True)
    return rec


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--tasks", help="comma-separated ids, default all")
    ap.add_argument("--stacks", default=",".join(STACKS))
    ap.add_argument("--repeat", type=int, default=5, help="timed runs after one warm-up (3 for LibreOffice); 0: none")
    ap.add_argument("--jobs", type=int, default=4)
    ap.add_argument("--no-save", action="store_true")
    a = ap.parse_args()
    dirs = sorted(d for d in HERE.glob("R[0-9][0-9]-*") if d.is_dir())
    if a.tasks:
        want = set(a.tasks.split(","))
        dirs = [d for d in dirs if d.name[:3] in want]
    stacks = tuple(a.stacks.split(","))
    with ThreadPoolExecutor(max_workers=a.jobs) as pool:  # outputs, oracles and validators in parallel
        tasks = list(pool.map(lambda d: run_task(d, stacks), dirs))
    for d, t in zip(dirs, tasks):  # timings one at a time
        inputs, out_name, work = t.pop("_run")
        for stack, s in t["stacks"].items():
            if a.repeat and s.get("seconds") is not None and "run" not in s:
                s["timing"] = time_stack(d, stack, inputs, out_name, work, a.repeat)
                if s["timing"]:
                    s["seconds"] = s["timing"]["median"]
                print(f"{t['id']} {stack:8s} {s['seconds']}s", flush=True)
    data = {"section": "features_docx_review", "versions": common.versions(), "tasks": tasks}
    if not a.no_save:
        print(common.save_results("features_docx_review", data))


if __name__ == "__main__":
    main()
