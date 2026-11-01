"""Validity section: which checks catch a corrupt or invalid .docx/.pptx, and which generators refuse to write one.

Regenerates ``results/validity.json`` from scratch:

0. base files: one document and one deck per generator (rdocx/rpptx, docx-js/pptxgenjs, python-docx/python-pptx),
   audited with the spec oracle (``oracle.py``) and run through every checker;
1. detection: every mutation of ``mutations.py`` applied to every base, then every checker of ``checkers.py`` run
   on the result (and the oracle, to confirm the label of the zip/opc/xml/schema faults);
2. prevention (``prevention.py``): the generators' documented footguns and their analogues, refused at the call
   or written, and what each checker says about what was written; atomic save under a failing write.

Usage: .cache/venv/bin/python validity/run.py [--jobs N] [--skip-prevention]
"""

from __future__ import annotations

import argparse
import json
import re
import statistics
import sys
import time
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parent))
sys.path.insert(0, str(HERE))
import common  # noqa: E402
import checkers  # noqa: E402
import mutations  # noqa: E402
import oracle  # noqa: E402

GENERATORS = {
    # base id: (format, stack, generator, command builder)
    "rdocx": {"docx": ("rdocx", lambda o, png: ([common.R_PY, HERE / "gen" / "docx_rdocx.py", o, png], common.env_rdocx())),
              "pptx": ("rpptx", lambda o, png: ([common.R_PY, HERE / "gen" / "pptx_rdocx.py", o, png], common.env_rdocx()))},
    "builtin": {"docx": ("docx-js", lambda o, png: ([common.NODE, HERE / "gen" / "docx_builtin.js", o, png], common.env_builtin())),
                "pptx": ("pptxgenjs", lambda o, png: ([common.NODE, HERE / "gen" / "pptx_builtin.js", o, png], common.env_builtin()))},
    "pylib": {"docx": ("python-docx", lambda o, png: ([common.PY, HERE / "gen" / "docx_pylib.py", o, png], common.env_pylib())),
              "pptx": ("python-pptx", lambda o, png: ([common.PY, HERE / "gen" / "pptx_pylib.py", o, png], common.env_pylib()))},
}
LAYER_ORDER = ["zip", "package", "xml", "schema", "reference", "semantic"]


def norm(issue: str) -> str:
    """An oracle issue without line numbers or counts, to compare a mutated file with its base."""
    return re.sub(r":\d+:", ":", re.sub(r"\d+ more error", "N more error", issue))


def make_picture(path: Path):
    r = common.run([common.PY, "-c", (
        "import sys; from PIL import Image, ImageDraw; im = Image.new('RGB', (160, 120), (128, 128, 128)); "
        "ImageDraw.Draw(im).rectangle([20, 20, 140, 100], fill=(30, 39, 97)); im.save(sys.argv[1])"), path],
        check=True)
    return r


def run_checkers(jobs, files):
    """files: list of (key, fmt, path[, original]). Returns {key: {checker: result}}; a checker that needs an
    original is left out where none is given."""
    tasks = [(f[0], name, f[2], f[3] if len(f) > 3 else None) for f in files for name in checkers.CHECKERS[f[1]]]
    out: dict = {}
    with ThreadPoolExecutor(max_workers=jobs) as pool:
        for (key, name, _, _), res in zip(tasks, pool.map(lambda t: checkers.check(t[1], t[2], t[3]), tasks)):
            out.setdefault(key, {})
            if res is not None:
                out[key][name] = res
    return out


def build_bases(work: Path):
    png = work / "picture.png"
    make_picture(png)
    bases = {}
    for base, per_fmt in GENERATORS.items():
        for fmt, (generator, cmd) in per_fmt.items():
            out = work / "bases" / f"{base}.{fmt}"
            out.parent.mkdir(parents=True, exist_ok=True)
            argv, env = cmd(out, png)
            r = common.run(argv, env=env, timeout=300)
            if r["rc"] != 0 or not out.exists():
                raise RuntimeError(f"base {base}.{fmt} ({generator}) failed: {r['stderr'][-1500:]}")
            bases[(fmt, base)] = {"format": fmt, "stack": base, "generator": generator, "path": out,
                                  "script": str((HERE / "gen" / Path(argv[1]).name).relative_to(common.BENCH)),
                                  "seconds": round(r["seconds"], 3)}
    return bases, png


def detect(jobs, work: Path, bases: dict):
    files, meta = [], {}
    for fmt in ("docx", "pptx"):
        for m in mutations.catalogue(fmt):
            for (bfmt, base), info in bases.items():
                if bfmt != fmt:
                    continue
                key = (m.id, base)
                try:
                    data = mutations.apply(m, info["path"].read_bytes())
                except mutations.NotApplicable as e:
                    meta[key] = {"applied": False, "reason": str(e)}
                    continue
                out = work / "mutated" / base / f"{m.id}{m.ext or '.' + fmt}"
                out.parent.mkdir(parents=True, exist_ok=True)
                out.write_bytes(data)
                base_issues = {norm(i) for i in info["oracle"]}
                new = [i for i in oracle.audit(out) if norm(i) not in base_issues]
                meta[key] = {"applied": True, "file": str(out.relative_to(common.BENCH)), "oracle_new": new}
                files.append((key, fmt, out, info["path"]))
    results = run_checkers(jobs, files)
    records = []
    for fmt in ("docx", "pptx"):
        for m in mutations.catalogue(fmt):
            rec = m.record()
            rec["format"] = rec.pop("fmt")
            rec["bases"] = {}
            for (bfmt, base) in bases:
                if bfmt != fmt:
                    continue
                entry = dict(meta[(m.id, base)])
                if entry["applied"]:
                    entry["checkers"] = results[(m.id, base)]
                    if m.basis in ("zip", "opc", "xml", "schema", "mce"):
                        confirmed = bool(entry["oracle_new"]) == (m.expected == "invalid")
                        entry["label_confirmed_by_oracle"] = confirmed
                rec["bases"][base] = entry
            records.append(rec)
    return records


def rate(hits, total):
    return {"detected": hits, "of": total, "rate": round(hits / total, 3) if total else None}


# What each stack's own workflow runs on a file it hands over (the union of its checks)
STACK_CHECKS = {
    "docx": {"stack:rdocx": ["rdocx_validate", "rdocx_open"],
             "stack:builtin": ["builtin_validate", "libreoffice", "pandoc"],
             "stack:pylib": ["pylib_open"]},
    "pptx": {"stack:rdocx": ["rdocx_validate", "rdocx_open"],
             "stack:builtin": ["builtin_validate", "libreoffice"],
             "stack:pylib": ["pylib_open"]},
}


def summarize(records, base_checks):
    """Detection per checker (and per stack: the union of the checks its workflow runs), counted on the
    (mutation, base) pairs whose base the checker (every check of the stack) passes."""
    summary = {}
    for fmt in ("docx", "pptx"):
        groups = {name: [name] for name in checkers.CHECKERS[fmt]} | STACK_CHECKS[fmt]

        def flagged(res, members):
            return any(res[m]["flagged"] for m in members)

        per = {}
        for name, members in groups.items():
            clean_bases = [b for (f, b), r in base_checks.items() if f == fmt and not flagged(r["checkers"], members)]
            layers, bases_of_rule = {}, {}
            inv_hit = inv_tot = 0
            fp_hit = fp_tot = 0
            office = []
            missed = []
            fps = []
            for rec in records:
                if rec["format"] != fmt:
                    continue
                pairs = [(b, e) for b, e in rec["bases"].items() if e["applied"] and b in clean_bases]
                flags = [flagged(e["checkers"], members) for _, e in pairs]
                if rec["office_only"]:
                    office.append({"id": rec["id"], "flagged": f"{sum(flags)}/{len(flags)}"})
                    continue
                if rec["expected"] == "invalid":
                    for table, k in ((layers, rec["layer"]), (bases_of_rule, rec["basis"])):
                        cell = table.setdefault(k, [0, 0])
                        cell[0] += sum(flags)
                        cell[1] += len(flags)
                    inv_hit += sum(flags)
                    inv_tot += len(flags)
                    if flags and not any(flags):
                        missed.append(rec["id"])
                else:
                    fp_hit += sum(flags)
                    fp_tot += len(flags)
                    if any(flags):
                        fps.append(rec["id"])
            base_fp = [b for (f, b), r in base_checks.items()
                       if f == fmt and r["spec_valid"] and flagged(r["checkers"], members)]
            base_tp = [b for (f, b), r in base_checks.items()
                       if f == fmt and not r["spec_valid"] and flagged(r["checkers"], members)]
            secs = [e["checkers"][name]["seconds"] for rec in records if rec["format"] == fmt
                    for e in rec["bases"].values() if e["applied"]] if name in checkers.CHECKERS[fmt] else []
            per[name] = {
                "label": checkers.LABELS.get(name, " + ".join(members)),
                "members": members,
                "bases_used": clean_bases,
                "detection_overall": rate(inv_hit, inv_tot),
                "detection_by_layer": {k: rate(*layers[k]) for k in LAYER_ORDER if k in layers},
                "detection_by_basis": {k: rate(*v) for k, v in sorted(bases_of_rule.items())},
                "false_positives_on_controls": rate(fp_hit, fp_tot),
                "false_positive_controls": fps,
                "flags_spec_valid_base": base_fp,
                "flags_spec_invalid_base": base_tp,
                "office_only": office,
                "missed_invalid_mutations": missed,
                "median_seconds": round(statistics.median(secs), 3) if secs else None,
            }
        uncaught = []
        for rec in records:
            if rec["format"] != fmt or rec["expected"] != "invalid" or rec["office_only"]:
                continue
            caught = any(
                e["checkers"][name]["flagged"]
                for b, e in rec["bases"].items() if e["applied"]
                for name in checkers.CHECKERS[fmt]
                if not base_checks[(fmt, b)]["checkers"][name]["flagged"]
            )
            if not caught:
                uncaught.append(rec["id"])
        summary[fmt] = {"checkers": per, "caught_by_no_checker": uncaught}
    return summary


def prevention_summary(cases):
    out = {}
    for c in cases:
        g = out.setdefault(c["generator"], {})
        key = c["outcome"]
        if key == "written_valid":
            key += "_same_as_control" if c["same_as_control"] else "_different"
        g[key] = g.get(key, 0) + 1
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--jobs", type=int, default=12)
    ap.add_argument("--skip-prevention", action="store_true")
    args = ap.parse_args()
    t0 = time.time()
    work = common.fresh_dir("validity")

    bases, png = build_bases(work)
    for info in bases.values():
        info["oracle"] = oracle.audit(info["path"])
        info["spec_valid"] = not info["oracle"]
    checked = run_checkers(args.jobs, [(k, k[0], v["path"], v["path"]) for k, v in bases.items()])
    for k, info in bases.items():
        info["checkers"] = checked[k]
    print(f"bases built and checked ({time.time() - t0:.0f}s)", flush=True)

    records = detect(args.jobs, work, bases)
    print(f"mutations checked ({time.time() - t0:.0f}s)", flush=True)
    summary = summarize(records, bases)

    prevention = {"cases": [], "atomic_save": []}
    if not args.skip_prevention:
        import prevention as prev
        prevention = prev.run_all(work / "prevention", png, args.jobs)
        print(f"prevention done ({time.time() - t0:.0f}s)", flush=True)
    for case in prevention["cases"]:
        case.pop("checks_files", None)

    data = {
        "section": "validity",
        "versions": common.versions(),
        "ground_truth": "ECMA-376 Part 1 (ISO/IEC 29500-4:2016 transitional schemas and prose), Part 2 (OPC), Part 3 (MCE)",
        "checkers": checkers.LABELS,
        "bases": [
            {k: (str(v.relative_to(common.BENCH)) if isinstance(v, Path) else v) for k, v in info.items()}
            for info in bases.values()
        ],
        "mutations": records,
        "summary": summary,
        "prevention": prevention["cases"],
        "prevention_summary": prevention_summary(prevention["cases"]),
        "atomic_save": prevention["atomic_save"],
        "seconds_total": round(time.time() - t0, 1),
    }
    out = common.save_results("validity", data)
    print(f"wrote {out}")
    for fmt in ("docx", "pptx"):
        print(f"\n{fmt}: detection on invalid / false positives on controls")
        for name, s in summary[fmt]["checkers"].items():
            d, f = s["detection_overall"], s["false_positives_on_controls"]
            print(f"  {name:26s} {d['detected']:3d}/{d['of']:<3d} ({d['rate']})  FP {f['detected']}/{f['of']}  "
                  f"bases used {s['bases_used']}")
        print("  caught by no checker:", summary[fmt]["caught_by_no_checker"])
    print("\nprevention outcomes:", json.dumps(data["prevention_summary"], indent=1))
    for s in data["atomic_save"]:
        print(f"  atomic save {s['stack']:7s} {s['save'][:60]:60s} input intact={s['input_intact']} target={s['target_state'][:60]}")


if __name__ == "__main__":
    main()
