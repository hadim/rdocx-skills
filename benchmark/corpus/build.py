"""Build the benchmark corpus into .cache/corpus/ and write corpus/manifest.json.

    .cache/venv/bin/python corpus/build.py            # synthetic files (no network)
    .cache/venv/bin/python corpus/build.py --real     # public files pinned in corpus/real.json
    .cache/venv/bin/python corpus/build.py --all      # both
    options: --no-checks (skip the open/validate checks), --jobs N

Synthetic (deterministic: same inputs, same bytes):
- docx-S, docx-M, docx-L: docx-js (the built-in docx skill's generator), ~5 / 50 / 200 A4 pages;
  docx-M-clean: docx-M's plan without comments and tracked changes; docx-M-lo: docx-M saved
  again by LibreOffice (soffice.py --convert-to docx), a second producer's XML habits.
- pptx-S, pptx-M, pptx-L: pptxgenjs structured decks (the built-in pptx skill's recipe, then
  its apply_theme.js), 15 / 60 / 150 slides; pptx-M-pythonpptx: the same plan with python-pptx.

Real (fetched, never committed): each file of real.json is downloaded into its own folder
under .cache/corpus/real/<id>/, its SHA-256 checked (skipped with a warning when offline or
gone), and read only by programs run with ``python -I`` from outside that folder.

Every file is then checked: it opens with rdocx / rpptx (CLI inspect; rdocx page count),
rdocx / rpptx validate, python-docx / python-pptx, LibreOffice (PDF export through the
built-in skill's soffice.py; the PDF is kept in .cache/corpus/lo_pdf/), and the built-in
skill's office/validate.py. A failing check is recorded in the manifest, never hidden.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import shutil
import subprocess
import sys
import time
import urllib.request
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

sys.dont_write_bytecode = True
HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parent))
sys.path.insert(0, str(HERE))
import common  # noqa: E402
import content  # noqa: E402
import images  # noqa: E402

OUT = common.CACHE / "corpus"
MANIFEST = HERE / "manifest.json"
REAL_PINS = HERE / "real.json"
N_IMAGES = 48
WORK = OUT / "work"


def rel(p: Path) -> str:
    return str(Path(p).resolve().relative_to(common.BENCH))


def sha256(p: Path) -> str:
    h = hashlib.sha256()
    with open(p, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def tail(text: str, n: int = 600) -> str:
    text = (text or "").strip().replace(str(common.BENCH) + "/", "")
    return text if len(text) <= n else "..." + text[-n:]


def fix_zip_times(src: Path, dest: Path) -> None:
    """Copy a zip with every entry's timestamp set to 2026-01-15 09:00 (same names, order, data and
    compression method). LibreOffice stamps entries with the current time and ignores
    SOURCE_DATE_EPOCH for OOXML; this keeps its output reproducible without touching any part."""
    import zipfile

    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(dest, "w") as zout:
        for info in zin.infolist():
            zi = zipfile.ZipInfo(info.filename, date_time=(2026, 1, 15, 9, 0, 0))
            zi.compress_type = info.compress_type
            zi.external_attr = info.external_attr
            zout.writestr(zi, zin.read(info.filename))


def node_env() -> dict:
    env = common.env_builtin()
    env["TZ"] = "UTC"
    return env


# ------------------------------------------------------------------------------------------
# Synthetic corpus
SYNTH = {
    "docx-S": {"kind": "docx", "size_class": "S", "producer": "docx-js"},
    "docx-M": {"kind": "docx", "size_class": "M", "producer": "docx-js"},
    "docx-L": {"kind": "docx", "size_class": "L", "producer": "docx-js"},
    "docx-M-clean": {"kind": "docx", "size_class": "M", "producer": "docx-js", "variant": "no comments, no tracked changes"},
    "docx-M-lo": {"kind": "docx", "size_class": "M", "producer": "docx-js, then LibreOffice", "variant": "docx-M round-tripped through LibreOffice (soffice.py --convert-to docx)"},
    "pptx-S": {"kind": "pptx", "size_class": "S", "producer": "pptxgenjs"},
    "pptx-M": {"kind": "pptx", "size_class": "M", "producer": "pptxgenjs"},
    "pptx-L": {"kind": "pptx", "size_class": "L", "producer": "pptxgenjs"},
    "pptx-M-pythonpptx": {"kind": "pptx", "size_class": "M", "producer": "python-pptx", "variant": "pptx-M's plan built with python-pptx on its default template"},
}


def path_of(fid: str) -> Path:
    kind = SYNTH[fid]["kind"]
    return OUT / kind / f"{fid}.{kind}"


def build_synthetic(jobs: int) -> dict:
    t0 = time.perf_counter()
    for d in ("assets", "plans", "docx", "pptx"):
        (OUT / d).mkdir(parents=True, exist_ok=True)
    pool = images.make_pool(N_IMAGES, OUT / "assets", common.FONTS_DIR)
    for im in pool:
        im["path"] = str(Path(im["path"]).resolve())
    print(f"[corpus] {len(pool)} pictures in {time.perf_counter() - t0:.1f}s")

    plans = {
        "docx-S": content.docx_plan("S", pool),
        "docx-M": content.docx_plan("M", pool),
        "docx-L": content.docx_plan("L", pool),
        "docx-M-clean": content.docx_plan("M", pool, clean=True),
        "pptx-S": content.deck_plan("S", pool),
        "pptx-M": content.deck_plan("M", pool),
        "pptx-L": content.deck_plan("L", pool),
    }
    plans["pptx-M-pythonpptx"] = plans["pptx-M"]
    for fid, plan in plans.items():
        (OUT / "plans" / f"{fid}.json").write_text(json.dumps(plan, ensure_ascii=False, indent=1))

    apply_theme = common.BUILTIN_PPTX / "scripts" / "apply_theme.js"

    def gen(fid: str):
        out = path_of(fid)
        out.unlink(missing_ok=True)
        plan = OUT / "plans" / f"{fid}.json"
        if fid.startswith("docx"):
            cmd = [common.NODE, HERE / "gen_docx.js", plan, out]
        elif fid == "pptx-M-pythonpptx":
            cmd = [common.PY, HERE / "gen_pptx_pythonpptx.py", plan, out]
        else:
            cmd = [common.NODE, HERE / "gen_pptx.js", plan, out, apply_theme]
        r = common.run(cmd, env=node_env(), cwd=HERE, timeout=600)
        return fid, r

    builds = {}
    with ThreadPoolExecutor(jobs) as ex:
        for fid, r in ex.map(gen, [f for f in SYNTH if f != "docx-M-lo"]):
            builds[fid] = {"rc": r["rc"], "seconds": round(r["seconds"], 2)}
            if r["rc"] != 0:
                builds[fid]["error"] = tail(r["stderr"] or r["stdout"])
            print(f"[corpus] {fid}: rc={r['rc']} {r['seconds']:.1f}s")

    # docx-M through LibreOffice: a second producer's XML habits.
    tmp = common.fresh_dir("corpus", "lo-roundtrip")
    src = tmp / "docx-M.docx"
    shutil.copy(path_of("docx-M"), src)
    outdir = tmp / "out"
    outdir.mkdir()
    r = common.run([common.PY, common.BUILTIN_DOCX / "scripts" / "office" / "soffice.py", "--headless",
                    "--convert-to", "docx", "--outdir", outdir, src], env=node_env(), cwd=tmp, timeout=600)
    produced = outdir / "docx-M.docx"
    path_of("docx-M-lo").unlink(missing_ok=True)
    if produced.exists():
        fix_zip_times(produced, path_of("docx-M-lo"))
    builds["docx-M-lo"] = {"rc": r["rc"] if produced.exists() else 1, "seconds": round(r["seconds"], 2)}
    if not produced.exists():
        builds["docx-M-lo"]["error"] = tail(r["stderr"] or r["stdout"])
    print(f"[corpus] docx-M-lo: rc={r['rc']} {r['seconds']:.1f}s")
    print(f"[corpus] synthetic build {time.perf_counter() - t0:.1f}s")
    return builds


# ------------------------------------------------------------------------------------------
# Real corpus
def download(url: str, dest: Path, timeout: int = 180) -> str | None:
    """Fetch url into dest (a new, empty folder holds it). Returns an error string or None."""
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0 (benchmark corpus fetch)"})
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp, open(dest, "wb") as f:
            shutil.copyfileobj(resp, f)
    except Exception as e:  # offline, gone, refused
        dest.unlink(missing_ok=True)
        return f"{type(e).__name__}: {e}"
    return None


def fetch_pinned(item: dict, folder: Path) -> dict:
    """Make sure folder holds item's file with the pinned hash. Returns status dict."""
    name = item["url"].split("?")[0].rstrip("/").split("/")[-1] or "download"
    if not name.lower().endswith((".docx", ".pptx", ".pdf")):
        name += "." + item.get("kind", "bin")
    dest = folder / name
    if dest.exists() and sha256(dest) == item["sha256"]:
        return {"status": "ok", "path": dest, "cached": True}
    if folder.exists():
        shutil.rmtree(folder)
    folder.mkdir(parents=True)
    err = download(item["url"], dest)
    if err:
        print(f"[corpus] WARNING: {item['url']}: {err} (skipped)")
        return {"status": "missing", "error": err}
    got = sha256(dest)
    if got != item["sha256"]:
        print(f"[corpus] WARNING: {item['url']}: sha256 {got} != pinned {item['sha256']} (skipped)")
        dest.unlink()
        return {"status": "hash_mismatch", "error": f"sha256 {got}"}
    return {"status": "ok", "path": dest, "cached": False}


def pdfinfo(p: Path) -> dict:
    r = common.run(["pdfinfo", p], timeout=60)
    out = {}
    for line in r["stdout"].splitlines():
        k, _, v = line.partition(":")
        if k in ("Producer", "Creator", "Pages"):
            out[k.lower()] = int(v) if k == "Pages" else v.strip()
    return out


def build_real(jobs: int) -> dict:
    pins = json.loads(REAL_PINS.read_text())
    status = {}

    def one(item):
        base = OUT / "real" / item["id"]
        st = fetch_pinned(item, base / "office")
        if item.get("pdf"):
            pst = fetch_pinned({**item["pdf"], "kind": "pdf"}, base / "pdf")
            if pst["status"] == "ok":
                pst["pdfinfo"] = pdfinfo(pst["path"])
            st["pdf"] = pst
        return item["id"], st

    with ThreadPoolExecutor(min(jobs, 6)) as ex:
        for fid, st in ex.map(one, pins["files"]):
            status[fid] = st
            print(f"[corpus] real {fid}: {st['status']}" + (f", pdf {st['pdf']['status']}" if "pdf" in st else ""))
    return status


# ------------------------------------------------------------------------------------------
# Checks (each one a separate process; downloaded files are only read by `python -I` scripts
# run from a neutral working folder, and by the CLIs / LibreOffice as data)
def checks_for(fid: str, path: Path, kind: str) -> list[tuple[str, list, dict]]:
    """(check name, argv, run options) for one file."""
    neutral = WORK / "neutral"
    neutral.mkdir(parents=True, exist_ok=True)
    lo_out = OUT / "lo_pdf"
    lo_out.mkdir(parents=True, exist_ok=True)
    office = common.BUILTIN_DOCX if kind == "docx" else common.BUILTIN_PPTX
    out = []
    if kind == "docx":
        out.append(("rdocx_open", [common.RDOCX, "inspect", "--json", path], {}))
        out.append(("rdocx_pages", [common.R_PY, "-I", "-c",
                                    "import sys, rdocx; print(rdocx.Document(sys.argv[1]).page_count(font_dir=sys.argv[2]))",
                                    path, common.FONTS_DIR], {}))
        out.append(("rdocx_validate", [common.RDOCX, "validate", path], {}))
        out.append(("python_docx_open", [common.PY, "-I", "-c",
                                         "import sys, docx; d = docx.Document(sys.argv[1]); print(len(d.paragraphs))", path], {}))
    else:
        out.append(("rpptx_open", [common.RPPTX, "inspect", "--json", path], {}))
        out.append(("rpptx_validate", [common.RPPTX, "validate", path], {}))
        out.append(("python_pptx_open", [common.PY, "-I", "-c",
                                         "import sys, pptx; p = pptx.Presentation(sys.argv[1]); print(len(p.slides))", path], {}))
    lo_dir = common.fresh_dir("corpus", "lo", fid)
    out.append(("libreoffice_pdf", [common.PY, office / "scripts" / "office" / "soffice.py", "--headless",
                                    "--convert-to", "pdf", "--outdir", lo_dir, path], {"lo_dir": lo_dir, "lo_out": lo_out}))
    out.append(("builtin_validate", [common.PY, office / "scripts" / "office" / "validate.py", path], {}))
    out.append(("stats", [common.PY, "-I", HERE / "ooxml_stats.py", path], {}))
    return out


def run_check(fid: str, name: str, argv: list, opts: dict) -> tuple[str, str, dict]:
    r = common.run(argv, env=common.env_builtin(), cwd=WORK / "neutral", timeout=600)
    res = {"ok": r["rc"] == 0, "rc": r["rc"], "seconds": round(r["seconds"], 1)}
    if name == "stats":
        try:
            return fid, name, {"ok": True, "data": json.loads(r["stdout"])}
        except json.JSONDecodeError:
            return fid, name, {"ok": False, "error": tail(r["stderr"] or r["stdout"])}
    if name == "rdocx_pages" and r["rc"] == 0:
        res["pages"] = int(r["stdout"].strip().splitlines()[-1])
    if name in ("rdocx_open", "rpptx_open") and r["rc"] == 0:
        try:
            d = json.loads(r["stdout"])
            res["summary"] = {k: d.get(k) for k in ("page_count", "page_count_source", "paragraphs", "tables", "slides", "layouts") if k in d}
            if isinstance(d.get("pictures"), list):
                res["summary"]["pictures"] = len(d["pictures"])
            if d.get("diagnostics"):
                res["diagnostics"] = d["diagnostics"][:5]
        except json.JSONDecodeError:
            res["ok"] = False
            res["error"] = "inspect did not print JSON"
    if name == "libreoffice_pdf":
        pdfs = list(opts["lo_dir"].glob("*.pdf"))
        res["ok"] = r["rc"] == 0 and len(pdfs) == 1
        if pdfs:
            info = pdfinfo(pdfs[0])
            res["pages"] = info.get("pages")
            dest = opts["lo_out"] / f"{fid}.pdf"
            shutil.move(pdfs[0], dest)
            res["pdf"] = rel(dest)
        shutil.rmtree(opts["lo_dir"], ignore_errors=True)
    if name in ("rdocx_validate", "rpptx_validate", "builtin_validate"):
        res["output"] = tail(r["stdout"] + "\n" + r["stderr"], 800)
    if not res["ok"] and "output" not in res:
        res["error"] = tail(r["stderr"] or r["stdout"])
    return fid, name, res


def run_checks(files: dict[str, tuple[Path, str]], jobs: int) -> dict:
    tasks = []
    slow = ("rdocx_pages", "libreoffice_pdf", "builtin_validate")
    for fid, (path, kind) in files.items():
        for name, argv, opts in checks_for(fid, path, kind):
            tasks.append((-path.stat().st_size * (2 if name in slow else 1), fid, name, argv, opts))
    tasks.sort(key=lambda t: t[0])  # longest first: the slow checks of the big files
    tasks = [t[1:] for t in tasks]
    out: dict[str, dict] = {fid: {} for fid in files}
    t0 = time.perf_counter()
    with ThreadPoolExecutor(jobs) as ex:
        for fid, name, res in ex.map(lambda t: run_check(*t), tasks):
            out[fid][name] = res
            if not res.get("ok"):
                print(f"[corpus] check {fid} {name}: FAILED {res.get('error') or res.get('output', '')[:200]}")
    print(f"[corpus] {len(tasks)} checks in {time.perf_counter() - t0:.1f}s")
    return out


# ------------------------------------------------------------------------------------------
# Manifest
COUNT_KEYS = ("paragraphs", "tables", "pictures", "comments", "revisions", "charts", "footnotes", "sections",
              "landscape_sections", "notes", "slides", "layouts", "masters", "media_files", "media_bytes")


def entry_for(fid: str, path: Path, kind: str, checks: dict) -> dict:
    stats = checks.get("stats", {}).get("data", {})
    e = {"id": fid, "path": rel(path), "kind": kind, "bytes": path.stat().st_size, "sha256": sha256(path)}
    app = stats.get("app", {})
    e["producer_application"] = " ".join(x for x in (app.get("Application"), app.get("AppVersion")) if x) or None
    if kind == "docx":
        e["pages"] = {"libreoffice": checks.get("libreoffice_pdf", {}).get("pages"),
                      "rdocx": checks.get("rdocx_pages", {}).get("pages"),
                      "app_xml": int(app["Pages"]) if app.get("Pages", "").isdigit() else None}
    else:
        e["slides"] = stats.get("slides")
    e["counts"] = {k: stats[k] for k in COUNT_KEYS if k in stats}
    for k in ("revisions_by_kind", "has_toc", "fields", "headers", "footers", "fonts"):
        if k in stats:
            e["counts"][k] = stats[k]
    if "error" in stats or not checks.get("stats", {}).get("ok", False):
        e["counts"]["error"] = stats.get("error") or checks.get("stats", {}).get("error")
    e["checks"] = {k: {kk: vv for kk, vv in v.items() if kk not in ("seconds",)} for k, v in sorted(checks.items()) if k != "stats"}
    e["checks_failed"] = sorted(k for k, v in checks.items() if not v.get("ok"))
    return e


def size_class_from_pages(kind: str, n: int | None) -> str | None:
    if n is None:
        return None
    if kind == "docx":
        return "S" if n <= 10 else "M" if n <= 80 else "L"
    return "S" if n <= 25 else "M" if n <= 80 else "L"


def write_manifest(entries: list[dict], rebuilt: set[str]) -> None:
    """Write the manifest: the entries just built, plus the old entries of the origins not rebuilt."""
    old = json.loads(MANIFEST.read_text()) if MANIFEST.exists() else {"files": []}
    keep = {e["id"]: e for e in old.get("files", []) if e.get("origin") not in rebuilt}
    for e in entries:
        keep[e["id"]] = e
    files = sorted(keep.values(), key=lambda e: (e["origin"], e["kind"], e["id"]))
    v = common.versions()
    v.pop("date", None)
    manifest = {
        "about": "Benchmark corpus: synthetic files built by corpus/build.py (deterministic) and public files "
                 "pinned in corpus/real.json (fetched into .cache/corpus/real/, never committed). Paths are relative "
                 "to benchmark/. pages.libreoffice: pages of LibreOffice's PDF export (kept in "
                 ".cache/corpus/lo_pdf/); pages.rdocx: rdocx page_count with the benchmark fonts; pages.app_xml: what "
                 "the producing application stored (Word's own count for real files). Counts come from corpus/ooxml_stats.py "
                 "(zipfile + lxml): docx paragraphs = every w:p of word/document.xml (table cells and TOC included), "
                 "pictures = pic:pic + v:imagedata in the body, headers, footers and notes, revisions = w:ins + w:del + "
                 "moves + property changes in those stories; pptx paragraphs = every a:p of the slides. checks: rdocx/rpptx "
                 "inspect, rdocx page_count (font_dir), rdocx/rpptx validate, python-docx/python-pptx open, LibreOffice PDF "
                 "export (soffice.py), the built-in skill's office/validate.py; checks_failed lists those that failed.",
        "versions": v,
        "summary": {
            "files": len(files),
            "synthetic": sum(1 for e in files if e["origin"] == "synthetic"),
            "real": sum(1 for e in files if e["origin"] == "real" and e.get("status", "ok") == "ok"),
            "real_missing": [e["id"] for e in files if e["origin"] == "real" and e.get("status", "ok") != "ok"],
            "real_pdf_twins": {e["id"]: e["pdf_twin"].get("route") for e in files
                               if e["origin"] == "real" and (e.get("pdf_twin") or {}).get("status") == "ok"},
            "with_failed_checks": {e["id"]: e["checks_failed"] for e in files if e.get("checks_failed")},
        },
        "files": files,
    }
    MANIFEST.write_text(common.scrub(json.dumps(manifest, indent=2, ensure_ascii=True)) + "\n")
    print(f"[corpus] wrote {rel(MANIFEST)} ({len(files)} files)")


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--real", action="store_true", help="only the public files of real.json")
    ap.add_argument("--all", action="store_true", help="synthetic and real")
    ap.add_argument("--no-checks", action="store_true")
    ap.add_argument("--jobs", type=int, default=8)
    a = ap.parse_args()
    do_synth, do_real = (not a.real) or a.all, a.real or a.all
    WORK.mkdir(parents=True, exist_ok=True)
    entries = []

    if do_synth:
        builds = build_synthetic(a.jobs)
        files = {fid: (path_of(fid), SYNTH[fid]["kind"]) for fid in SYNTH if path_of(fid).exists()}
        checks = {fid: {} for fid in files}
        if a.no_checks:
            for fid, (p, _) in files.items():
                checks[fid] = {"stats": {"ok": True, "data": json.loads(common.run([common.PY, "-I", HERE / "ooxml_stats.py", p])["stdout"])}}
        else:
            checks = run_checks(files, a.jobs)
        gen_versions = common.versions()
        for fid, meta in SYNTH.items():
            if fid not in files:
                entries.append({"id": fid, "origin": "synthetic", "kind": meta["kind"], "status": "build_failed", "build": builds.get(fid)})
                continue
            e = entry_for(fid, files[fid][0], meta["kind"], checks[fid])
            gen = {"docx-js": f"docx-js {gen_versions['docx-js']} (corpus/gen_docx.js)",
                   "docx-js, then LibreOffice": f"docx-js {gen_versions['docx-js']}, then {gen_versions['libreoffice']} (soffice.py --convert-to docx)",
                   "pptxgenjs": f"pptxgenjs {gen_versions['pptxgenjs']} (corpus/gen_pptx.js) + built-in pptx scripts/apply_theme.js",
                   "python-pptx": f"python-pptx {gen_versions['python-pptx']} (corpus/gen_pptx_pythonpptx.py)"}[meta["producer"]]
            e.update({"origin": "synthetic", "size_class": meta["size_class"], "generator": gen,
                      "plan": rel(OUT / "plans" / f"{fid if fid != 'docx-M-lo' else 'docx-M'}.json"), "licence": "synthetic (this repository)"})
            if "variant" in meta:
                e["variant"] = meta["variant"]
            if fid.startswith("docx"):
                e["target_pages"] = content.SIZES_DOCX[meta["size_class"]]["pages"]
            entries.append(dict(sorted(e.items(), key=lambda kv: list(ENTRY_ORDER).index(kv[0]) if kv[0] in ENTRY_ORDER else 99)))

    if do_real:
        if not REAL_PINS.exists():
            print("[corpus] no corpus/real.json")
        else:
            pins = {p["id"]: p for p in json.loads(REAL_PINS.read_text())["files"]}
            status = build_real(a.jobs)
            present = {fid: (st["path"], pins[fid]["kind"]) for fid, st in status.items() if st["status"] == "ok"}
            checks = {fid: {} for fid in present}
            if not a.no_checks and present:
                checks = run_checks(present, a.jobs)
            for fid, pin in pins.items():
                st = status[fid]
                base = {"id": fid, "origin": "real", "kind": pin["kind"], "status": st["status"],
                        "url": pin["url"], "source_page": pin.get("source_page"), "licence": pin.get("licence"),
                        "publisher": pin.get("publisher"), "title": pin.get("title"), "why": pin.get("why")}
                if st["status"] != "ok":
                    base["error"] = st.get("error")
                    base["sha256"] = pin["sha256"]
                    entries.append(base)
                    continue
                e = entry_for(fid, st["path"], pin["kind"], checks.get(fid, {}))
                e.update({k: v for k, v in base.items() if k not in e})
                e["generator"] = e.get("producer_application")
                if pin["kind"] == "docx":
                    n = (st.get("pdf", {}).get("pdfinfo") or {}).get("pages") or (e.get("pages") or {}).get("app_xml")
                else:
                    n = e.get("slides")
                e["size_class"] = size_class_from_pages(pin["kind"], n)
                e["band"] = pin.get("band")
                if pin.get("notes"):
                    e["notes"] = pin["notes"]
                if "pdf" in st:
                    p = st["pdf"]
                    twin = {"status": p["status"], "url": pin["pdf"]["url"], "route": pin["pdf"].get("route")}
                    if p["status"] == "ok":
                        twin.update({"path": rel(p["path"]), "sha256": pin["pdf"]["sha256"], "bytes": p["path"].stat().st_size,
                                     "page_size": pin["pdf"].get("page_size"), **p.get("pdfinfo", {})})
                        # The PDF's own words on who made it (route is the pin's reading of them).
                        said = (twin.get("producer", "") + " " + twin.get("creator", "")).lower()
                        twin["creator_is_office"] = any(w in said for w in ("microsoft", "powerpoint")) or "word" in twin.get("creator", "").lower()
                        if pin["kind"] == "docx":
                            twin["pages_match"] = {"app_xml": twin.get("pages") == (e.get("pages") or {}).get("app_xml"),
                                                   "libreoffice": twin.get("pages") == (e.get("pages") or {}).get("libreoffice"),
                                                   "rdocx": twin.get("pages") == (e.get("pages") or {}).get("rdocx")}
                        else:
                            twin["pages_match"] = twin.get("pages") == e.get("slides")
                    e["pdf_twin"] = twin
                    e["pages_office_pdf"] = twin.get("pages")
                entries.append(dict(sorted(e.items(), key=lambda kv: list(ENTRY_ORDER).index(kv[0]) if kv[0] in ENTRY_ORDER else 99)))

    write_manifest(entries, ({"synthetic"} if do_synth else set()) | ({"real"} if do_real else set()))
    return 0


ENTRY_ORDER = ("id", "origin", "kind", "status", "size_class", "band", "path", "bytes", "sha256", "pages", "pages_office_pdf", "slides",
               "target_pages", "generator", "producer_application", "variant", "plan", "title", "publisher", "url",
               "source_page", "licence", "why", "notes", "pdf_twin", "counts", "checks_failed", "checks")

if __name__ == "__main__":
    sys.exit(main())
