"""Turn results/*.json into the Markdown tables of the benchmark.

Each table is a named block. ``report.py`` fills every block it finds between
``<!-- bench:NAME:start -->`` and ``<!-- bench:NAME:end -->`` in the files given with
``--apply`` (the benchmark README, the plugin README), and always writes every block to
``RESULTS.md`` so that a run can be read before it is published.

    .cache/venv/bin/python report.py [--apply README.md ../plugin/README.md]
"""

from __future__ import annotations

import argparse
import json
import re
import statistics
import sys
from pathlib import Path

BENCH = Path(__file__).resolve().parent
RESULTS = BENCH / "results"

STACKS = ("rdocx", "builtin", "pylib")
STACK_NAMES = {
    "rdocx": "rdocx-skills",
    "builtin": "Built-in skill",
    "pylib": "python-docx / python-pptx",
}
STACK_SHORT = {"rdocx": "rdocx-skills", "builtin": "built-in", "pylib": "python-docx/pptx"}
ICON = {"pass": "\u2705", "partial": "\u26a0\ufe0f", "fail": "\u274c", "unsupported": "\u2796"}
XML = "\U0001f9ec"
FEATURE_FILES = ("features_docx_authoring", "features_docx_review", "features_pptx")
FEATURE_LABEL = {
    "features_docx_authoring": "Word: create and edit",
    "features_docx_review": "Word: review, layout and render",
    "features_pptx": "PowerPoint",
}


def load(name):
    p = RESULTS / f"{name}.json"
    return json.loads(p.read_text()) if p.exists() else None


def pct(x):
    return f"{round(100 * x)} %"


def bold_best(values, best="max", fmt=str):
    """Format a row of numbers, the best one(s) in bold. None prints as a dash."""
    nums = [v for v in values if v is not None]
    target = fmt((max if best == "max" else min)(nums)) if nums else None
    return [
        "\u2013" if v is None else (f"**{fmt(v)}**" if fmt(v) == target and len(nums) > 1 else fmt(v))
        for v in values
    ]


def table(header, rows, align=None):
    align = align or ["---"] + [":---:"] * (len(header) - 1)
    out = ["| " + " | ".join(header) + " |", "| " + " | ".join(align) + " |"]
    out += ["| " + " | ".join(str(c) for c in r) + " |" for r in rows]
    return "\n".join(out)


# --- features -----------------------------------------------------------------------------
def task_status(s):
    st = s.get("status")
    if st != "pass" and s.get("precision") is not None and (s.get("precision") or 0) > 0:
        return "partial"
    return st


def feature_tasks():
    out = []
    for f in FEATURE_FILES:
        d = load(f)
        if not d:
            continue
        for t in d["tasks"]:
            out.append((f, t))
    return out


def feature_stats():
    tasks = feature_tasks()
    stats = {s: {"pass": 0, "partial": 0, "fail": 0, "unsupported": 0, "loc": [], "xml": 0, "proc": []} for s in STACKS}
    for _, t in tasks:
        for s in STACKS:
            d = t["stacks"].get(s)
            if not d:
                continue
            st = task_status(d)
            stats[s][st] = stats[s].get(st, 0) + 1
            if st != "unsupported":
                stats[s]["loc"].append(d.get("loc") or 0)
                stats[s]["proc"].append(d.get("processes") or 0)
                stats[s]["xml"] += bool(d.get("raw_xml"))
    # Lines of code on the tasks every stack passes: the like-for-like total.
    common = [t for _, t in tasks if all(task_status(t["stacks"][s]) == "pass" for s in STACKS)]
    for s in STACKS:
        stats[s]["common_loc"] = sum(t["stacks"][s].get("loc") or 0 for t in common)
    return len(tasks), len(common), stats


def block_features_summary():
    n, ncommon, st = feature_stats()
    rows = [
        ["\u2705 Done (oracle passes)"] + bold_best([st[s]["pass"] for s in STACKS], fmt=lambda v: f"{v} / {n}"),
        ["\u26a0\ufe0f Partly done"] + [str(st[s]["partial"]) for s in STACKS],
        ["\u274c Attempted, wrong result"] + bold_best([st[s]["fail"] for s in STACKS], best="min"),
        ["\u2796 Not possible with the stack"] + bold_best([st[s]["unsupported"] for s in STACKS], best="min"),
        [f"\u270d\ufe0f Lines of code, the {ncommon} tasks all three pass"] + bold_best([st[s]["common_loc"] for s in STACKS], best="min"),
        ["\u270d\ufe0f Lines of code per task (median)"] + bold_best(
            [statistics.median(st[s]["loc"]) if st[s]["loc"] else None for s in STACKS], best="min",
            fmt=lambda v: f"{v:g}"),
        [f"{XML} Tasks that need hand-written OOXML"] + bold_best([st[s]["xml"] for s in STACKS], best="min"),
        ["\U0001f527 Commands per task (median)"] + bold_best(
            [statistics.median(st[s]["proc"]) if st[s]["proc"] else None for s in STACKS], best="min",
            fmt=lambda v: f"{v:g}"),
    ]
    return table([f"{n} tasks"] + [STACK_NAMES[s] for s in STACKS], rows)


def cell(s):
    st = task_status(s)
    if st == "unsupported":
        return ICON[st]
    txt = f"{ICON[st]} {s.get('loc', '?')}"
    if s.get("raw_xml"):
        txt += f" {XML}"
    if s.get("precision") is not None:
        txt += f" (P {s['precision']:.2f}, R {s.get('recall', 0):.2f})"
    return txt


def block_features_matrix():
    parts = []
    for f in FEATURE_FILES:
        d = load(f)
        if not d:
            continue
        rows = [[f"{t['id']} {t['title']}"] + [cell(t["stacks"][s]) for s in STACKS] for t in d["tasks"]]
        parts.append(f"**{FEATURE_LABEL[f]}**\n\n" + table(["Task"] + [STACK_NAMES[s] for s in STACKS], rows,
                                                          ["---", ":---", ":---", ":---"]))
    legend = (f"{ICON['pass']} done \u00b7 {ICON['partial']} partly \u00b7 {ICON['fail']} wrong result \u00b7 "
              f"{ICON['unsupported']} not possible \u00b7 the number is the lines of code an agent writes \u00b7 "
              f"{XML} hand-written OOXML")
    return legend + "\n\n" + "\n\n".join(parts)


# --- validity -----------------------------------------------------------------------------
CHECKER_ROWS = {
    "docx": [("rdocx_validate", "`rdocx validate`"), ("builtin_validate", "built-in `validate.py`"),
             ("pylib_open", "python-docx opens it"), ("libreoffice", "LibreOffice converts it"),
             ("pandoc", "pandoc reads it")],
    "pptx": [("rdocx_validate", "`rpptx validate`"), ("builtin_validate", "built-in `validate.py`"),
             ("pylib_open", "python-pptx opens it"), ("libreoffice", "LibreOffice converts it")],
}
LAYERS = ("zip", "package", "xml", "schema", "reference")


def real_false_alarms():
    """Checks that fail on untouched files written by Word or PowerPoint (the real corpus)."""
    m = json.loads((BENCH / "corpus" / "manifest.json").read_text())
    real = [f for f in m["files"] if f.get("origin") == "real"]
    out = {}
    for kind in ("docx", "pptx"):
        files = [f for f in real if f["kind"] == kind]
        for key, names in {
            "rdocx_validate": ("rdocx_validate", "rpptx_validate"),
            "rdocx_open": ("rdocx_open", "rpptx_open"),
            "builtin_validate": ("builtin_validate",),
            "pylib_open": ("python_docx_open", "python_pptx_open"),
            "libreoffice": ("libreoffice_pdf",),
        }.items():
            n = sum(1 for f in files if any(c in (f.get("checks_failed") or []) for c in names))
            out[(kind, key)] = (n, len(files))
    return out


def block_validity_detection():
    v = load("validity")
    if not v:
        return "_validity results missing_"
    fa = real_false_alarms()
    parts = []
    for fmt in ("docx", "pptx"):
        ch = v["summary"][fmt]["checkers"]
        rows = []
        best = max(ch[k]["detection_overall"]["rate"] for k, _ in CHECKER_ROWS[fmt] if k in ch)
        for key, label in CHECKER_ROWS[fmt]:
            c = ch.get(key)
            if not c:
                continue
            rate = pct(c["detection_overall"]["rate"])
            by = c["detection_by_layer"]
            n, of = fa.get((fmt, key), (None, None))
            rows.append(
                [label, f"**{rate}**" if c["detection_overall"]["rate"] == best else rate]
                + [f"{by[l]['detected']}/{by[l]['of']}" if l in by else "\u2013" for l in LAYERS]
                + [f"{c['false_positives_on_controls']['detected']}/{c['false_positives_on_controls']['of']}",
                   f"{n}/{of}" if n is not None else "\u2013"]
            )
        parts.append(f"**{fmt}**\n\n" + table(
            ["Checker", "Invalid files caught", "zip", "package", "XML", "schema", "references",
             "False alarms: valid controls", "False alarms: untouched Office files"], rows))
    return "\n\n".join(parts)


PREVENTION_ORDER = ("rdocx", "rpptx", "python-docx", "python-pptx", "docx-js", "pptxgenjs")


def block_validity_prevention():
    v = load("validity")
    if not v:
        return "_validity results missing_"
    ps = v["prevention_summary"]
    rows = []
    for g in PREVENTION_ORDER:
        s = ps.get(g, {})
        total = sum(s.values())
        stack = "rdocx-skills" if g in ("rdocx", "rpptx") else ("built-in skill" if g in ("docx-js", "pptxgenjs") else "no skill")
        rows.append([
            g, stack, total,
            f"\U0001f6d1 {s.get('refused_at_call', 0)}",
            f"\U0001f4a5 {s.get('written_invalid', 0)}",
            f"\U0001f92b {s.get('written_valid_different', 0)}",
            str(s.get("written_valid_same_as_control", 0)),
        ])
    return table(["Generator", "Stack", "Mistakes tried", "Refused at the call", "Wrote an invalid file",
                  "Wrote a valid file, silently different", "Harmless (same as the right call)"], rows)


def block_validity_atomic():
    v = load("validity")
    if not v:
        return "_validity results missing_"
    rows = []
    for a in v["atomic_save"]:
        rows.append([a["stack"], a["format"], a["save"], "\u2705 intact" if a.get("target_intact") else f"\U0001f4a5 {a.get('target_state', 'damaged')}"])
    return table(["Stack", "Format", "Save that fails half-way (disk full)", "File it was writing over"], rows,
                 ["---", ":---:", "---", "---"])


# --- speed --------------------------------------------------------------------------------
SPEED_OPS = {
    "docx": [
        ("docx.read", "\U0001f4d6 Read all the text"),
        ("docx.replace", "\u270f\ufe0f Counted replacement (the skill's helper)"),
        ("docx.comment", "\U0001f4ac Comment on a phrase"),
        ("docx.accept", "\u2705 Accept all tracked changes"),
        ("docx.render-page1", "\U0001f5bc\ufe0f Page 1 to PNG"),
        ("docx.render", "\U0001f5bc\ufe0f Every page to PNG"),
        ("docx.pdf", "\U0001f4c4 PDF"),
        ("docx.markdown", "\U0001f4dd Markdown"),
        ("docx.validate", "\U0001f50e Validate"),
        ("docx.create", "\U0001f3d7\ufe0f Build the document from scratch"),
    ],
    "pptx": [
        ("pptx.read", "\U0001f4d6 Read text and notes"),
        ("pptx.replace", "\u270f\ufe0f Counted replacement"),
        ("pptx.duplicate", "\U0001f5c2\ufe0f Duplicate a slide"),
        ("pptx.render-slide1", "\U0001f5bc\ufe0f Slide 1 to PNG"),
        ("pptx.render", "\U0001f5bc\ufe0f Every slide to PNG"),
        ("pptx.thumbnail", "\U0001f9e9 Thumbnails"),
        ("pptx.pdf", "\U0001f4c4 PDF"),
        ("pptx.validate", "\U0001f50e Validate"),
        ("pptx.create", "\U0001f3d7\ufe0f Build the deck from scratch"),
    ],
}
SPEED_FILE = {"docx": ("docx-M", "a 52-page Word report"), "pptx": ("pptx-M", "a 60-slide deck")}


def secs(x):
    return f"{x:.2f} s" if x < 10 else f"{x:.0f} s"


def ratio_words(builtin, rdocx):
    if not builtin or not rdocx:
        return "\u2013"
    r = builtin / rdocx
    def x(v):
        return f"{v:.0f}" if v >= 10 else f"{v:.1f}"
    return f"\U0001f680 {x(r)}\u00d7 faster" if r >= 1.2 else (f"\U0001f422 {x(1 / r)}\u00d7 slower" if r <= 1 / 1.2 else "\u2248 same")


def speed_rows(s, op, fid):
    return [r for r in s["summary"] if r["op"] == op and r["file"] == fid]


def block_speed():
    s = load("speed")
    if not s:
        return "_speed results pending_"
    parts = []
    for kind in ("docx", "pptx"):
        fid, label = SPEED_FILE[kind]
        rows = []
        for op, name in SPEED_OPS[kind]:
            r = speed_rows(s, op, fid)
            if not r:
                continue
            r = r[0]
            vals = [r.get(st) if r.get(f"{st}_status", "ok") == "ok" else None for st in STACKS]
            cells = bold_best(vals, best="min", fmt=secs)
            cells = [("\u2796" if c == "\u2013" else c) for c in cells]
            rows.append([name] + cells + [ratio_words(r.get("builtin"), r.get("rdocx"))])
        parts.append(f"**{label[0].upper() + label[1:]}** (`{fid}`), median wall time of the command an agent runs\n\n" + table(
            ["Operation"] + [STACK_NAMES[x] for x in STACKS] + ["rdocx-skills against the built-in skill"], rows))
    return "\n\n".join(parts)


def block_speed_details():
    s = load("speed")
    if not s:
        return "_speed results pending_"
    rows = []
    for r in s["summary"]:
        size = r.get("slides") or (r.get("pages") or {}).get("libreoffice")
        unit = "slides" if r.get("slides") else "pages"
        variants = "; ".join(
            f"{k} {secs(v['median'])}" if v.get("status") == "ok" and v.get("median") is not None else f"{k} {v.get('status')}"
            for k, v in r["variants"].items())
        rows.append([r["op"], f"{r['file']} ({size} {unit})", variants])
    return table(["Operation", "File", "Median per variant"], rows, ["---", "---", "---"])


def chart_speed(out: Path):
    """Every page or slide to PNG, against the size of the file, for both renderers."""
    s = load("speed")
    if not s:
        return None
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt

    colors = {"rdocx": "#2a78d6", "builtin": "#eb6834"}
    names = {"rdocx": "rdocx-skills (rdocx / rpptx render)", "builtin": "built-in skill (LibreOffice + pdftoppm)"}
    fig, axes = plt.subplots(1, 2, figsize=(9, 4), dpi=100)
    for ax, (op, title, unit) in zip(axes, (("docx.render", "Word: every page to PNG", "pages"),
                                             ("pptx.render", "PowerPoint: every slide to PNG", "slides"))):
        for st in ("rdocx", "builtin"):
            for origin in ("synthetic", "real"):
                pts = []
                for r in s["summary"]:
                    if r["op"] != op or r.get("origin") != origin or r.get(f"{st}_status") != "ok" or r.get(st) is None:
                        continue
                    n = r.get("slides") or (r.get("pages") or {}).get("libreoffice")
                    if n:
                        pts.append((n, r[st]))
                pts.sort()
                if not pts:
                    continue
                if origin == "synthetic":
                    ax.plot([p[0] for p in pts], [p[1] for p in pts], "o-", color=colors[st], lw=2, ms=6,
                            label=f"{names[st]}, synthetic files")
                else:
                    ax.plot([p[0] for p in pts], [p[1] for p in pts], "o", color=colors[st], ms=7, mfc="none",
                            mew=1.5, label=f"{names[st]}, public Office files")
        ax.set_title(title, fontsize=10, loc="left")
        ax.set_xlabel(unit, fontsize=9)
        ax.set_ylabel("seconds (median)", fontsize=9)
        ax.grid(True, color="#e5e5e5", lw=0.8)
        for side in ("top", "right"):
            ax.spines[side].set_visible(False)
        ax.tick_params(labelsize=8)
    handles, labels = axes[0].get_legend_handles_labels()
    fig.legend(handles, labels, fontsize=8, frameon=False, loc="lower center", ncol=2)
    fig.tight_layout(rect=(0, 0.14, 1, 1))
    out.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(out, format=out.suffix[1:], **({"metadata": {"Date": None}} if out.suffix == ".svg" else {}))
    plt.close(fig)
    return out


# --- visual -------------------------------------------------------------------------------
def block_visual_fidelity():
    v = load("visual")
    if not v:
        return "_visual results pending_"
    fs = v["fidelity_summary"]
    groups = [("docx", "Word files, PDF exported by Word"),
              ("pptx_office_exported", "Deck, PDF exported by PowerPoint"),
              ("pptx_other_producer", "Decks, PDF from another producer (macOS, Acrobat)")]
    rows = []
    for key, label in groups:
        g = fs.get(key)
        if not g:
            continue
        r, lo = g["rdocx"]["all_rendered"], g["libreoffice"]["all_rendered"]
        n = g["documents"]
        def m(d, k):
            return d[k]["median"] if d.get(k) and d[k].get("n") else None
        for metric, mlabel, fmt in (
            ("page_count_exact", "\U0001f4cf Same page count as Office", None),
            ("ssim_median_per_doc", "\U0001f5bc\ufe0f Page similarity (SSIM, median)", lambda x: f"{x:.2f}"),
            ("words_same_page", "\U0001f524 Words on the same page as Office (median)", pct),
            ("words_within_5pt", "\U0001f3af Words within 5 pt of Office's position (median)", pct),
        ):
            if fmt is None:
                vals = [r["page_count_exact"], lo["page_count_exact"]]
                cells = bold_best(vals, fmt=lambda x: str(x))
                cells = [f"{cells[0]} / {r['rendered']}" + (f" (refuses {r['failed']})" if r.get("failed") else ""),
                         f"{cells[1]} / {lo['rendered']}"]
            else:
                cells = bold_best([m(r, metric), m(lo, metric)], fmt=fmt)
            rows.append([f"{label} ({n})" if metric == "page_count_exact" else "", mlabel] + cells)
    return table(["Files", "Measure", "rdocx-skills (rdocx / rpptx)", "Built-in skill (LibreOffice)"], rows,
                 ["---", "---", ":---:", ":---:"])


CAP_ICON = {"pass": "\u2705", "partial": "\u26a0\ufe0f", "unsupported": "\u2796", "fail": "\u274c"}


def block_visual_capabilities():
    v = load("visual")
    if not v:
        return "_visual results pending_"
    caps = [c for c in v["capabilities"] if not c["capability"].startswith("time to")]
    order = []
    for c in caps:
        key = (c["capability"], c["kind"] if c["stack"] != "pylib" else None)
        if (c["capability"], c["kind"]) not in order and c["stack"] != "pylib":
            order.append((c["capability"], c["kind"]))
    rows = []
    for cap, kind in order:
        cells = []
        for st in STACKS:
            match = [c for c in caps if c["capability"] == cap and c["stack"] == st and (c["kind"] == kind or c["kind"] == "docx+pptx")]
            cell = CAP_ICON.get(match[0]["status"], match[0]["status"]) if match else "\u2013"
            hit = re.search(r"Word's page for (\d+) of (\d+)", match[0].get("evidence", "")) if match else None
            if hit:
                cell += f" {hit.group(1)}/{hit.group(2)} right"
            cells.append(cell)
        rows.append([f"{cap[0].upper() + cap[1:]} ({kind})"] + cells)
    timing = [c for c in v["capabilities"] if c["capability"].startswith("time to") and c["stack"] != "pylib"]
    trows = []
    for kind in sorted({c["kind"] for c in timing}):
        vals = [next((c.get("seconds") for c in timing if c["kind"] == kind and c["stack"] == st), None)
                for st in ("rdocx", "builtin")]
        trows.append([kind] + bold_best(vals, best="min", fmt=secs))
    notes = ["\"right\": the heading's page agrees with Word's own PDF of the same file."]
    sp = load("speed")
    if sp:
        fits = [m for m in sp["measurements"] if m["op"] == "pptx.fit" and m["stack"] == "rdocx"]
        files = {m["file"] for m in fits}
        bad = {m["file"] for m in fits if m["status"] != "ok"}
        if bad:
            notes.append(f"The overflow report works through the Python API, but `rpptx fit` (CLI) refused "
                         f"{len(bad)} of the {len(files)} decks of the speed test (an rdocx finding).")
    return (table(["What the agent can see"] + [STACK_NAMES[s] for s in STACKS], rows) + "\n\n"
            + "\n".join(f"<sub>{n}</sub><br>" for n in notes) + "\n\n"
            + "Time to the first picture of page 1, one cold run\n\n"
            + table(["File", "rdocx-skills", "Built-in skill"], trows))


def block_visual_figures():
    v = load("visual")
    if not v:
        return "_visual results pending_"
    out = []
    for f in v["figures"]:
        out.append(f"![{f['caption']}]({f['path']})\n<sub>{f['caption']}"
                   + (f": {f['document']}, page {f['page']}" if f.get("page") else "") + "</sub>")
    return "\n\n".join(out)


# --- the scoreboard -----------------------------------------------------------------------
def block_scoreboard():
    n, ncommon, st = feature_stats()
    v, s, vis = load("validity"), load("speed"), load("visual")
    rows = [
        [f"\U0001f9f0 Tasks done, out of {n}"] + bold_best([st[x]["pass"] for x in STACKS]),
        [f"\u270d\ufe0f Lines of code an agent writes ({ncommon} tasks all three do)"] + bold_best(
            [st[x]["common_loc"] for x in STACKS], best="min"),
        [f"{XML} Tasks that need hand-written OOXML"] + bold_best([st[x]["xml"] for x in STACKS], best="min"),
    ]
    if v:
        ps = v["prevention_summary"]
        def refused(gens):
            tot = sum(sum(ps.get(g, {}).values()) for g in gens)
            return sum(ps.get(g, {}).get("refused_at_call", 0) for g in gens) / tot if tot else None
        rows.append(["\U0001f6d1 Known mistakes refused before a file is written"] + bold_best(
            [refused(("rdocx", "rpptx")), refused(("docx-js", "pptxgenjs")), refused(("python-docx", "python-pptx"))],
            fmt=pct))
        for fmt in ("docx", "pptx"):
            ch = v["summary"][fmt]["checkers"]
            rows.append([f"\U0001f50e Corrupt or invalid {fmt} caught by the stack's checks"] + bold_best(
                [ch[f"stack:{x}"]["detection_overall"]["rate"] for x in STACKS], fmt=pct))
    if s:
        for op, fid, label in (("docx.render", "docx-M", "\u26a1 Every page of a 52-page Word report to PNG"),
                               ("pptx.render", "pptx-M", "\u26a1 Every slide of a 60-slide deck to PNG"),
                               ("docx.replace", "docx-M", "\u270f\ufe0f Counted replacement in a 52-page Word report"),
                               ("docx.create", "docx-M", "\U0001f3d7\ufe0f Build a 52-page Word report from scratch")):
            r = speed_rows(s, op, fid)
            if r:
                r = r[0]
                vals = [r.get(x) if r.get(f"{x}_status", "ok") == "ok" else None for x in STACKS]
                rows.append([label] + [("\u2796" if c == "\u2013" else c) for c in bold_best(vals, best="min", fmt=secs)])
    if vis:
        d = vis["fidelity_summary"]["docx"]
        r, lo = d["rdocx"]["all_rendered"], d["libreoffice"]["all_rendered"]
        rows.append(["\U0001f5a8\ufe0f Word files with the same page count as Word's own PDF"] + [
            f"{x}" for x in bold_best([r["page_count_exact"], lo["page_count_exact"]],
                                      fmt=lambda k: f"{k} / {d['documents']}")] + ["\u2796"])
    return table(["", "rdocx-skills", "Built-in skill", "python-docx / python-pptx"], rows)


def block_versions():
    d = load("features_docx_authoring") or {}
    v = d.get("versions", {})
    keys = ("date", "machine", "rdocx", "rpptx", "rdocx_pin", "python-docx", "python-pptx", "docx-js",
            "pptxgenjs", "pandoc", "libreoffice", "node", "builtin_skills_sha256")
    return table(["", "Version"], [[k, v.get(k, "\u2013")] for k in keys], ["---", "---"])


def block_date():
    d = load("features_docx_authoring") or {}
    v = d.get("versions", {})
    return (f"Run of {v.get('date', '?')}: rdocx {v.get('rdocx', '?').split()[-1]} and rpptx "
            f"{v.get('rpptx', '?').split()[-1]} (pin `{v.get('rdocx_pin', '?')}`) against the built-in skills "
            f"(docx-js {v.get('docx-js', '?')}, pptxgenjs {v.get('pptxgenjs', '?')}, "
            f"{' '.join(v.get('libreoffice', '?').split()[:2])}, {v.get('pandoc', '?')}) and "
            f"python-docx {v.get('python-docx', '?')} / python-pptx {v.get('python-pptx', '?')}.")


BLOCKS = {
    "date": block_date,
    "scoreboard": block_scoreboard,
    "features-summary": block_features_summary,
    "features-matrix": block_features_matrix,
    "validity-detection": block_validity_detection,
    "validity-prevention": block_validity_prevention,
    "validity-atomic": block_validity_atomic,
    "speed": block_speed,
    "speed-details": block_speed_details,
    "visual-fidelity": block_visual_fidelity,
    "visual-capabilities": block_visual_capabilities,
    "visual-figures": block_visual_figures,
    "versions": block_versions,
}


def apply(path: Path, blocks: dict) -> int:
    text = path.read_text()
    n = 0
    for name, body in blocks.items():
        pat = re.compile(rf"(<!-- bench:{re.escape(name)}:start -->\n).*?(<!-- bench:{re.escape(name)}:end -->)", re.S)
        text, k = pat.subn(lambda m: m.group(1) + body + "\n" + m.group(2), text)
        n += k
    path.write_text(text)
    return n


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--apply", nargs="*", default=[], type=Path)
    a = ap.parse_args()
    blocks = {name: fn() for name, fn in BLOCKS.items()}
    chart = chart_speed(BENCH / "figures" / "speed-render.svg")
    if chart:
        print(f"wrote {chart}")
    (BENCH / "RESULTS.md").write_text(
        "# Results (generated by report.py)\n\n" + "\n\n".join(f"## {k}\n\n{v}" for k, v in blocks.items()) + "\n")
    for p in a.apply:
        print(f"{p}: {apply(p, blocks)} block(s) filled")
    print(f"wrote {BENCH / 'RESULTS.md'}")


if __name__ == "__main__":
    sys.exit(main())
