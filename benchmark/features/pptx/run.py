"""Run the pptx feature tasks for the three stacks and write results/features_pptx.json.

For each task folder Pxx-slug/: make_input.py writes the inputs into a fresh work folder; each stack's
implementation (rdocx.py|rdocx.sh, builtin.py|builtin.sh|builtin.js, pylib.py, or STACK.unsupported with the
reason) runs once on them; check.py (the neutral oracle) decides pass or fail and runs the validators; then the
command is timed as an agent runs it (median of 5 runs after a warm-up, 3 when it runs LibreOffice).

Each implementation's first lines carry its metadata as a comment:
    # bench: tools=a, b; processes=N; raw_xml=yes|no; files=extra1,extra2
LOC counts the non-blank, non-comment lines of the entry file and of the extra files it names.

    .cache/venv/bin/python features/pptx/run.py [--tasks P01,P02] [--stacks rdocx,pylib] [--no-timing] [--merge]
"""

from __future__ import annotations

import argparse
import json
import os
import re
import shutil
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
BENCH = HERE.parents[1]
sys.path.insert(0, str(BENCH))
import common  # noqa: E402

STACKS = ("rdocx", "builtin", "pylib")
ENTRY = {"rdocx": ("rdocx.py", "rdocx.sh"), "builtin": ("builtin.js", "builtin.py", "builtin.sh"), "pylib": ("pylib.py",)}

D = "{in}/deck.pptx"
TASKS = [
    dict(id="P01", category="create", title="Create a 10-slide structured deck",
         spec="From content.json: a 16:9 deck on a theme named Northwind (its colours and fonts: Cambria headings, "
              "Calibri body) and its layouts: title slide, agenda bullets, section divider, a table slide, a native "
              "clustered-column chart slide, a picture slide, a two-column slide, two bullet slides, a closing slide; "
              "speaker notes on every slide; slide numbers on every slide but the title slide.",
         args=["{in}/content.json", "{out}"], out="deck.pptx"),
    dict(id="P02", category="read", title="Read a deck into JSON",
         spec="Titles, every text (groups and table cells included) and speaker notes of an 8-slide deck as JSON; "
              "the oracle looks for the marker strings of each slide in that slide only.",
         args=[D, "{out}"], out="deck.json", pptx=False),
    dict(id="P03", category="edit", title="Counted replace across slides and notes",
         spec="Replace 'Project Falcon' with 'Project Osprey' exactly 7 times across titles, bullets, a table cell, a "
              "group and notes; three occurrences are split across runs with different formatting, two decoys "
              "differ in case. Formatting outside the match is kept, the replacement takes a format of the "
              "matched runs. The same command with an expected count of 8 must refuse and write nothing.",
         args=[D, "{out}", "Project Falcon", "Project Osprey", "7"], out="out.pptx", original=D,
         refuse=[D, "{out}", "Project Falcon", "Project Osprey", "8"]),
    dict(id="P04", category="edit", title="Duplicate, move, delete and hide slides",
         spec="On an 8-slide deck (slides referred to by their original numbers): duplicate slide 3 (content and "
              "notes) to just after slide 5, move the last slide to position 2, delete slide 4 (it holds the only "
              "use of a picture and has notes), hide slide 6. The oracle checks the final order by marker text, the "
              "hidden flag, the copy's notes and picture, relationships, content types and that no part is left "
              "orphaned (the deleted slide, its notes and its picture gone).",
         args=[D, "{out}"], out="out.pptx", original=D),
    dict(id="P05", category="edit", title="Merge slides from another deck keeping their look",
         spec="Import slides 2 and 3 of deck B (theme Harbor, its own master and layouts; a native chart with its "
              "workbook, a picture) into deck A (theme Northwind) at position 2, keeping B's look: the copies use "
              "B's layouts, master and theme, A's slides keep A's.",
         args=["{in}/a.pptx", "{in}/b.pptx", "{out}"], out="out.pptx", original="{in}/a.pptx"),
    dict(id="P06", category="review", title="Report overflowing text frames",
         spec="On a 3-slide deck with 14 text frames of known fit (5 overflow: a wrapped text box, an unwrapped "
              "line, a title and a body placeholder whose size and font size are inherited, a table whose long "
              "cell pushes it off the slide; 9 fit, among them a near miss and a table whose cell is taller than "
              "its stored row, which PowerPoint grows), report the overflowing frames as JSON; scored by "
              "precision and recall (pass: both 1).",
         args=[D, "{out}"], out="report.json", pptx=False, lo={"builtin"}),
    dict(id="P07", category="review", title="Comment anchored to a shape, then a reply",
         spec="Add a comment by 'Dana Reviewer' on slide 3 anchored to the shape named 'Key figure', then a "
              "reply by 'Sam Author'; the oracle accepts modern comments (anchor = the shape's id) and reports "
              "legacy ones as not anchored.",
         args=[D, "{out}"], out="out.pptx", original=D),
    dict(id="P08", category="edit", title="Change a chart's data",
         spec="A formatted clustered-column chart (2 series, 4 quarters, series fills, data labels, title): new "
              "values and a fifth category; formatting kept; caches, formulas and the embedded workbook match.",
         args=[D, "{out}"], out="out.pptx", original=D),
    dict(id="P09", category="edit", title="Replace a picture keeping its frame",
         spec="Replace the image of a cropped picture on slide 3 with new.png, keeping position, size, crop, name "
              "and alt text; the same old image used on slide 2 must stay; no orphan image part.",
         args=[D, "{in}/new.png", "{out}"], out="out.pptx", original=D),
    dict(id="P10", category="edit", title="Table: add a row, merge cells, fill a cell",
         spec="On a styled 4x4 table: add a row at the end formatted like the row above (texts given), merge its "
              "first two cells, fill cell (2, 3) with FFC000.",
         args=[D, "{out}"], out="out.pptx", original=D),
    dict(id="P11", category="edit", title="Set and change speaker notes",
         spec="Add notes to slides 2, 4 and 6 (none before) and replace those of slide 3; the others stay.",
         args=[D, "{out}"], out="out.pptx", original=D),
    dict(id="P12", category="render", title="Render every slide to PNG at 100 dpi",
         spec="One PNG per slide, 1333x750 pixels, not blank: the agent's visual feedback loop.",
         args=[D, "{out}"], out="png", pptx=False, lo={"builtin"}),
    dict(id="P13", category="convert", title="Export the deck to Markdown",
         spec="Titles as headings in slide order, bullets as list items with their nesting, the table as a GFM "
              "table, speaker notes.",
         args=[D, "{out}"], out="deck.md", pptx=False),
    dict(id="P14", category="layout", title="Slide numbers and footer as the Header and Footer dialog does",
         spec="Slide numbers and the footer 'Northwind Confidential' on every slide but the title slide, as "
              "slide-level sldNum and ftr placeholders matching the layouts' (Apply to All, 'Don't show on title "
              "slide').",
         args=[D, "{out}"], out="out.pptx", original=D),
    dict(id="P15", category="create", title="Fill a .potx template",
         spec="From a .potx (3 layouts used, one instruction slide to drop): a .pptx with a title slide, a bullet "
              "slide and a picture slide filled through their placeholders (title, subtitle, body, picture).",
         args=["{in}/template.potx", "{in}/photo.png", "{out}"], out="out.pptx", original="{in}/template.potx"),
]

# Judgement calls and observations that the oracle does not print, per (task, stack).
NOTES = {
    ("P01", "rdocx"): "rpptx.Presentation() is 16:9 with the default layouts; theme name, colours and fonts through "
                      "the theme API; set_header_footer() numbers every slide but title-layout slides.",
    ("P01", "builtin"): "Structured deck as SKILL.md describes: 6 layouts (defineSlideMaster) with named placeholders "
                        "and slideNumber, scheme colours, sections, addNotes, then applyTheme. pptxgenjs writes "
                        "[Content_Types] overrides for slideMaster2..10, which do not exist (package warning).",
    ("P01", "pylib"): "python-pptx's default template is 4:3: slide size set to 16:9, but its layouts keep their "
                      "10-inch geometry (content sits on the left); theme via lxml on the theme part; slide numbers "
                      "by cloning the layout's sldNum placeholder and writing the field.",
    ("P02", "builtin"): "slide_text.py's Markdown parsed back into the JSON.",
    ("P03", "builtin"): "Cross-run replacement written by hand on the slide and notes XML; count checked before "
                        "anything is written.",
    ("P03", "pylib"): "Cross-run replacement written by hand over runs (python-pptx has no replace); emptied runs "
                      "are left empty.",
    ("P04", "builtin"): "add_slide.py duplicates without the notes (documented): notes slide copied and registered by "
                        "hand; sldIdLst edited; clean.py drops the deleted slide, its notes and its picture.",
    ("P04", "pylib"): "No duplicate, move or delete API: shapes deep-copied with their image relationships, "
                      "_sldIdLst edited, drop_rel.",
    ("P05", "builtin"): "No cross-deck helper: the slides' part closure (used layouts, master, theme, chart, workbook, "
                        "image) copied under free names, master layout list pruned, ids renumbered, rels, "
                        "sldMasterIdLst, sldIdLst and content types written. Notes left behind (they would need a "
                        "second notes master).",
    ("P05", "pylib"): "OPC level: A related to B's part objects, parts renamed to free names, master pruned and "
                      "registered with lxml. Notes left behind.",
    ("P05", "rdocx"): "One command; notes travel with the slides.",
    ("P06", "rdocx"): "prs.fit_report(): flags 'Growing table' because table cells are measured against the stored "
                      "row height, while PowerPoint (and rpptx's own renderer) grow the row and the table stays on "
                      "the slide.",
    ("P06", "builtin"): "Automatable stand-in for the skill's visual QA: LibreOffice render at 50 dpi, ink in a "
                        "0.4-inch band outside each text frame that no other shape covers. Blind to overflow into a "
                        "neighbouring shape; any row growth reads as overflow.",
    ("P06", "pylib"): "Own estimator: word wrap with Carlito/Caladea metrics (Pillow), sizes inherited through layout "
                      "and master (lxml), table cells against the stored row height.",
    ("P07", "builtin"): "Modern comments (p188) written by hand: authors part, comment part, the slide's commentRel "
                        "extension, rels, content types.",
    ("P07", "pylib"): "Modern comments (p188) as OPC parts; the slide's commentRel extension with lxml.",
    ("P08", "builtin"): "Caches and formulas with lxml; the embedded workbook rewritten with openpyxl (the skill names "
                        "no tool for it).",
    ("P09", "builtin"): "A new media part: the old image is shared with slide 2, so it is not overwritten.",
    ("P09", "pylib"): "get_or_add_image_part, then the blip's r:embed repointed and the old relationship dropped.",
    ("P10", "rdocx"): "rows.add_row() writes cells without runs, so the run-level edit (paragraphs[0].runs[0].text) "
                      "raises IndexError on them; cell.text = keeps the size through endParaRPr.",
    ("P10", "pylib"): "No add-row API: the last a:tr deep-copied; merge and fill through the API.",
    ("P11", "builtin"): "No notes helper: new notes slides cloned from an existing one, rels and content types by hand.",
    ("P13", "rdocx"): "Valid CommonMark, but every space of titles, runs and notes is written &#32; and punctuation "
                      "is escaped; the chart is left out (diagnostic: unsupported chart projection) where "
                      "slide_text.py prints its data.",
    ("P13", "builtin"): "slide_text.py is the skill's Markdown reader; it prints paragraphs without their level, so the "
                        "nesting cannot be restored. An agent that writes python-pptx code instead gets the pylib "
                        "result.",
    ("P14", "pylib"): "clone_placeholder() for the layout's footer and slide-number placeholders; the slidenum field "
                      "with lxml.",
    ("P14", "builtin"): "The layout's ftr and sldNum placeholders copied into each non-title slide with lxml.",
    ("P15", "builtin"): "add_slide.py from a layout makes an empty slide: placeholders, the picture placeholder with a "
                        "computed srcRect, media, rels written by hand; template content type switched to "
                        "presentation. thumbnail.py (picking layouts by eye) has no role in a script.",
    ("P15", "pylib"): "python-pptx refuses a .potx: content type patched in memory; the instruction slide removed "
                      "through _sldIdLst.",
}


def task_dir(task):
    return next(HERE.glob(f"{task['id']}-*"))


def entry(tdir, stack):
    for name in ENTRY[stack]:
        if (tdir / name).exists():
            return tdir / name
    return None


def header(path):
    meta = {"tools": [], "processes": 1, "raw_xml": False, "files": []}
    for line in path.read_text().splitlines()[:5]:
        m = re.match(r"^\s*(?:#|//)\s*bench:\s*(.*)$", line)
        if not m:
            continue
        for item in m.group(1).split(";"):
            if "=" not in item:
                continue
            k, v = (s.strip() for s in item.split("=", 1))
            if k in ("tools", "files"):
                meta[k] = [s.strip() for s in v.split(",") if s.strip()]
            elif k == "processes":
                meta[k] = int(v)
            elif k == "raw_xml":
                meta[k] = v.lower() in ("yes", "true", "1")
    return meta


def loc(path):
    n, in_block = 0, False
    for line in path.read_text().splitlines():
        s = line.strip()
        if path.suffix == ".js":
            if in_block:
                in_block = "*/" not in s
                continue
            if s.startswith("/*"):
                in_block = "*/" not in s
                continue
            if not s or s.startswith("//"):
                continue
        elif path.suffix in (".py", ".sh"):
            if not s or s.startswith("#"):
                continue
        elif not s:
            continue
        n += 1
    return n


def env_for(stack):
    env = {"rdocx": common.env_rdocx, "builtin": common.env_builtin, "pylib": common.env_pylib}[stack]()
    env["PPTX_SKILL"] = str(common.BUILTIN_PPTX)
    env["FONT_DIR"] = str(common.FONTS_DIR)
    return env


def command(path, stack, args):
    if path.suffix == ".js":
        return [common.NODE, path, *args]
    if path.suffix == ".sh":
        return ["bash", path, *args]
    return [common.R_PY if stack == "rdocx" else common.PY, path, *args]


def fill(args, inputs, out):
    return [a.replace("{in}", str(inputs)).replace("{out}", str(out)) for a in args]


def run_check(tdir, out, inputs, task, validate=True):
    cmd = [common.PY, tdir / "check.py", out, "--inputs", inputs]
    if task.get("original") and validate:
        cmd += ["--original", fill([task["original"]], inputs, out)[0]]
    if not validate:
        cmd.append("--no-validate")
    r = common.run(cmd, env=common.env_pylib(), timeout=900)
    try:
        return json.loads(r["stdout"])
    except json.JSONDecodeError:
        return {"ok": False, "errors": [f"check.py crashed: {(r['stderr'] or r['stdout'])[-800:]}"]}


def run_task(task, stacks, timing):
    tdir = task_dir(task)
    work = common.fresh_dir("features_pptx", task["id"])
    inputs = work / "inputs"
    common.run([common.PY, tdir / "make_input.py", inputs], env=common.env_pylib(), check=True)
    rec = {"id": task["id"], "title": task["title"], "category": task["category"], "spec": task["spec"],
           "task_md": (tdir / "task.md").read_text().strip(), "stacks": {}}
    for stack in stacks:
        unsupported = tdir / f"{stack}.unsupported"
        if unsupported.exists():
            print(f"{task['id']} {stack:8} unsupported", flush=True)
            rec["stacks"][stack] = {"status": "unsupported", "loc": 0, "files": [], "tools": [], "processes": 0,
                                    "raw_xml": False, "seconds": None,
                                    "notes": unsupported.read_text().strip()}
            continue
        path = entry(tdir, stack)
        if path is None:
            continue
        meta = header(path)
        files = [path] + [tdir / f for f in meta["files"]]
        env = env_for(stack)
        sdir = work / stack
        sdir.mkdir()
        out = sdir / task["out"]
        r = common.run(command(path, stack, fill(task["args"], inputs, out)), env=env, cwd=sdir, timeout=600)
        res = {"status": "fail", "loc": sum(loc(f) for f in files), "files": [f.name for f in files],
               "tools": meta["tools"], "processes": meta["processes"], "raw_xml": meta["raw_xml"], "seconds": None}
        notes = []
        if r["rc"] != 0:
            notes.append(f"exit {r['rc']}: {(r['stderr'] or r['stdout']).strip()[-600:]}")
        chk = run_check(tdir, out, inputs, task, validate=task.get("pptx", True))
        ok = r["rc"] == 0 and chk.get("ok")
        notes += chk.get("errors", [])
        if chk.get("valid"):
            res["valid"] = chk["valid"]
            for k, v in (chk.get("valid_notes") or {}).items():
                notes.append(f"{k}: {v}")
        info = chk.get("info") or {}
        for k in ("precision", "recall"):
            if k in info:
                res[k] = info[k]
        if info:
            res["oracle_info"] = info
        if task.get("refuse"):
            rout = sdir / "refused" / task["out"]
            rout.parent.mkdir()
            rr = common.run(command(path, stack, fill(task["refuse"], inputs, rout)), env=env, cwd=sdir, timeout=600)
            refused = rr["rc"] != 0 and not rout.exists()
            res["refuses_wrong_count"] = refused
            if not refused:
                ok = False
                notes.append(f"wrong expected count not refused (exit {rr['rc']}, output written: {rout.exists()})")
        res["status"] = "pass" if ok else "fail"
        if timing and r["rc"] == 0:
            tdirs = sdir / "timing"

            def setup(i):
                d = tdirs / str(i)
                if d.exists():
                    shutil.rmtree(d)
                d.mkdir(parents=True)

            lo = stack in task.get("lo", set())
            t = common.timeit(lambda i: command(path, stack, fill(task["args"], inputs, tdirs / str(i) / task["out"])),
                              repeat=3 if lo else 5, warmup=1, env=env, cwd=sdir, setup=setup, timeout=600)
            res["seconds"] = round(t["median"], 3) if t.get("ok") else None
            if t.get("ok"):
                res["seconds_samples"] = [round(s, 3) for s in t["samples"]]
        extra = NOTES.get((task["id"], stack))
        if extra:
            notes.insert(0, extra)
        res["notes"] = " | ".join(notes)
        rec["stacks"][stack] = res
        print(f"{task['id']} {stack:8} {res['status']:5} loc={res['loc']:4} t={res['seconds']}  {res['notes'][:150]}",
              flush=True)
    return rec


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--tasks", default="")
    ap.add_argument("--stacks", default=",".join(STACKS))
    ap.add_argument("--no-timing", action="store_true")
    ap.add_argument("--merge", action="store_true", help="update the existing results instead of replacing them")
    a = ap.parse_args()
    want = {t for t in a.tasks.split(",") if t}
    stacks = [s for s in a.stacks.split(",") if s]
    old = {}
    path = common.RESULTS / "features_pptx.json"
    if a.merge and path.exists():
        old = {t["id"]: t for t in json.loads(path.read_text())["tasks"]}
    for task in TASKS:
        if want and task["id"] not in want:
            continue
        rec = run_task(task, stacks, not a.no_timing)
        if task["id"] in old:
            old[task["id"]]["stacks"].update(rec["stacks"])
            rec["stacks"] = {s: old[task["id"]]["stacks"][s] for s in STACKS if s in old[task["id"]]["stacks"]}
        old[task["id"]] = rec
    data = {"section": "features_pptx", "versions": common.versions(),
            "tasks": [old[t["id"]] for t in TASKS if t["id"] in old]}
    print(common.save_results("features_pptx", data))


if __name__ == "__main__":
    main()
