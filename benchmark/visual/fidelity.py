"""Part 1 of the visual section: how close each renderer's pages are to the PDF Office made of the same file.

For every real file with a PDF twin (corpus/manifest.json), three PDFs are compared page by page:

- ``office``: the PDF the publisher released (Word or PowerPoint laid it out; the ground truth);
- ``rdocx``: ``rdocx convert F --to pdf --font-dir FONTS`` (``rpptx`` for decks), the rdocx stack's PDF route;
  the PNG route (``rdocx render --dpi 60 --font-dir FONTS``) is compared too, as a check that the picture an
  agent looks at is the same layout;
- ``libreoffice``: the built-in skills' route, ``scripts/office/soffice.py --headless --convert-to pdf`` with the
  benchmark's fonts (``FONTCONFIG_FILE``).

All three PDFs are rasterised by the same pdftoppm at the same DPI, so that only the layout differs.
"""

from __future__ import annotations

import json
import shutil
from collections import Counter
from pathlib import Path

import numpy as np

import common
import metrics as M

DPI = 60
WORK = common.WORK / "visual" / "fidelity"
STACKS = ("rdocx", "libreoffice")
MIN_PAGE_WORDS_FIG = 80  # a page needs this many words to be picked for a figure


def twins() -> list[dict]:
    man = json.loads((common.BENCH / "corpus" / "manifest.json").read_text())
    out = []
    for f in man["files"]:
        tw = f.get("pdf_twin") or {}
        if f.get("origin") == "real" and tw.get("status") == "ok" and (common.BENCH / tw["path"]).exists():
            out.append(f)
    return sorted(out, key=lambda f: (f["kind"], tw_pages(f)))


def tw_pages(f):
    return f["pdf_twin"]["pages"]


def soffice_py(kind: str) -> Path:
    skill = common.BUILTIN_DOCX if kind == "docx" else common.BUILTIN_PPTX
    return skill / "scripts" / "office" / "soffice.py"


def _diagnostics(stderr: str) -> dict:
    lines = [ln.strip() for ln in stderr.splitlines() if ln.strip() and not ln.startswith("Written to")]
    return dict(Counter(lines).most_common())


# --- rendering ----------------------------------------------------------------------------
def render_rdocx_pdf(f: dict, d: Path) -> dict:
    cli = common.RDOCX if f["kind"] == "docx" else common.RPPTX
    out = d / "rdocx.pdf"
    r = common.run([cli, "convert", common.BENCH / f["path"], "--to", "pdf", "--font-dir", common.FONTS_DIR,
                    "-o", out, "--force"], env=common.env_rdocx(), timeout=600)
    ok = r["rc"] == 0 and out.exists()
    return {"ok": ok, "rc": r["rc"], "seconds": round(r["seconds"], 2), "pdf": str(out) if ok else None,
            "diagnostics": _diagnostics(r["stderr"]),
            "error": None if ok else (r["stderr"] or r["stdout"]).strip()[-400:]}


def render_rdocx_png(f: dict, d: Path) -> dict:
    out = d / "rdocx_png"
    if out.exists():
        shutil.rmtree(out)
    out.mkdir(parents=True)
    cli = common.RDOCX if f["kind"] == "docx" else common.RPPTX
    r = common.run([cli, "render", common.BENCH / f["path"], "-o", out, "--dpi", DPI, "--font-dir", common.FONTS_DIR],
                   env=common.env_rdocx(), timeout=600)
    pngs = sorted(out.glob("*.png"), key=lambda p: int("".join(ch for ch in p.stem.rsplit("_", 1)[1] if ch.isdigit())))
    return {"ok": r["rc"] == 0 and bool(pngs), "rc": r["rc"], "seconds": round(r["seconds"], 2),
            "pngs": [str(p) for p in pngs], "error": None if r["rc"] == 0 else r["stderr"].strip()[-400:]}


def render_libreoffice_pdf(f: dict, d: Path) -> dict:
    out_dir = d / "lo"
    if out_dir.exists():
        shutil.rmtree(out_dir)
    out_dir.mkdir(parents=True)
    src = common.BENCH / f["path"]
    r = common.run([common.PY, soffice_py(f["kind"]), "--headless", "--convert-to", "pdf", "--outdir", out_dir, src],
                   env=common.env_builtin(), cwd=out_dir, timeout=600)
    pdfs = list(out_dir.glob("*.pdf"))
    ok = bool(pdfs)
    out = d / "libreoffice.pdf"
    if ok:
        shutil.move(pdfs[0], out)
    return {"ok": ok, "rc": r["rc"], "seconds": round(r["seconds"], 2), "pdf": str(out) if ok else None,
            "error": None if ok else (r["stderr"] or r["stdout"]).strip()[-400:]}


def render(f: dict) -> dict:
    d = WORK / f["id"]
    d.mkdir(parents=True, exist_ok=True)
    return {"rdocx": render_rdocx_pdf(f, d), "rdocx_png": render_rdocx_png(f, d),
            "libreoffice": render_libreoffice_pdf(f, d)}


# --- measuring ------------------------------------------------------------------------------
def measure(f: dict, rend: dict) -> dict:
    d = WORK / f["id"]
    office_pdf = common.BENCH / f["pdf_twin"]["path"]
    o_sizes = M.page_sizes(office_pdf)
    o_imgs = M.rasterize(office_pdf, d / "ras_office", DPI, "p")
    o_arr = [M.load_gray(p) for p in o_imgs]
    o_words = M.words(office_pdf)
    o_count = Counter(w[0] for w in o_words[1])
    blank = [M.ssim(a, np.full_like(a, 255)) for a in o_arr]
    entry = {
        "id": f["id"],
        "kind": f["kind"],
        "producer": f["pdf_twin"].get("producer") or f["pdf_twin"].get("creator"),
        "route": f["pdf_twin"]["route"],
        "office_pdf": f["pdf_twin"]["path"],
        "source_page": f.get("source_page"),
        "licence": f.get("licence"),
        "pdf_group": "office-exported" if f["pdf_twin"]["route"] == "office-native" else "other-producer",
        "pages": {"office": len(o_sizes)},
        "office_page_size_pt": [round(x, 1) for x in o_sizes[0]],
        "office_words": len(o_words[1]),
        "ssim_blank_page": M.describe(blank),
        "stacks": {},
        "per_page": {"office_words": [o_count.get(k, 0) for k in range(len(o_sizes))],
                     "office_ink": [round(M.ink(a), 4) for a in o_arr]},
    }
    o_chars = M.font_chars(office_pdf)
    total_chars = sum(o_chars.values()) or 1
    entry["fonts_office"] = {k: round(v / total_chars, 4) for k, v in o_chars.most_common()}
    rfonts = {}
    for stack in STACKS:
        r = rend[stack]
        if not r["ok"]:
            entry["pages"][stack] = None
            entry["stacks"][stack] = {"status": "fail", "error": r["error"]}
            continue
        pdf = Path(r["pdf"])
        sizes = M.page_sizes(pdf)
        entry["pages"][stack] = len(sizes)
        imgs = M.rasterize(pdf, d / f"ras_{stack}", DPI, "p")
        paired = min(len(o_arr), len(imgs))
        ss, f1 = [], []
        for k in range(paired):
            b = M.load_gray(imgs[k])
            ss.append(M.ssim(o_arr[k], b))
            f1.append(M.ink_f1(o_arr[k], b))
        spans = M.actual_text_spans(pdf)
        words_pdf = M.drop_actual_text(pdf, d / f"{stack}_flat.pdf") if spans else pdf
        wm = M.match_words(o_words, M.words(words_pdf))
        size_ok = abs(sizes[0][0] - o_sizes[0][0]) <= 2 and abs(sizes[0][1] - o_sizes[0][1]) <= 2
        st = {
            "status": "ok",
            "seconds": r["seconds"],
            "page_diff": len(sizes) - len(o_sizes),
            "pages_paired": paired,
            "page_size_pt": [round(x, 1) for x in sizes[0]],
            "page_size_matches": size_ok,
            "ssim": M.describe(ss),
            "ink_f1": M.describe(f1),
            "words": {k: v for k, v in wm.items() if k != "per_page_same_page"},
            "actual_text_spans": spans,
            "words_read_from": "pdftocairo round trip (ActualText dropped)" if spans else "the PDF itself",
        }
        if f["kind"] == "pptx" and stack == "libreoffice":
            # check that the cairo round trip used for rpptx's PDFs does not move words: same metric on
            # LibreOffice's PDF read directly and through the round trip
            wc = M.match_words(o_words, M.words(M.drop_actual_text(pdf, d / "libreoffice_flat.pdf")))
            st["words_cairo_roundtrip_check"] = {k: wc[k] for k in ("matched", "same_page", "within_5pt", "within_20pt")}
        if stack == "rdocx":
            st["diagnostics"] = r["diagnostics"]
            png = rend["rdocx_png"]
            if png["ok"]:
                pp = [Path(p) for p in png["pngs"]]
                n = min(len(o_arr), len(pp))
                st["png_route"] = {"pages": len(pp), "ssim": M.describe([M.ssim(o_arr[k], M.load_gray(pp[k]))
                                                                       for k in range(n)]),
                                   "ssim_vs_own_pdf": M.describe([M.ssim(M.load_gray(imgs[k]), M.load_gray(pp[k]))
                                                                  for k in range(min(len(imgs), len(pp)))])}
            else:
                st["png_route"] = {"pages": 0, "error": png["error"]}
        entry["stacks"][stack] = st
        entry["per_page"][f"{stack}_ssim"] = [round(x, 4) for x in ss]
        entry["per_page"][f"{stack}_ink_f1"] = [None if x is None else round(x, 4) for x in f1]
        entry["per_page"][f"{stack}_same_page"] = [wm["per_page_same_page"].get(k) for k in range(len(o_sizes))]
        rfonts[stack] = M.pdf_fonts(pdf)
    # fonts: which Office families each renderer drew with the family itself or a metric-compatible clone
    fonts = {"renderer_fonts": {s: sorted(v) for s, v in rfonts.items()}}
    for stack, have in rfonts.items():
        miss = [fam for fam in entry["fonts_office"] if not M.font_available(fam, have)]
        fonts[f"missing_{stack}"] = miss
        fonts[f"missing_share_{stack}"] = round(sum(entry["fonts_office"][m] for m in miss), 4)
    if rfonts:
        both = [fam for fam in entry["fonts_office"] if all(fam in fonts[f"missing_{s}"] for s in rfonts)]
        fonts["missing_all_renderers"] = both
        fonts["missing_share_all_renderers"] = round(sum(entry["fonts_office"][m] for m in both), 4)
        fonts["missing_fonts_dominate"] = fonts["missing_share_all_renderers"] > 0.5
    entry["fonts"] = fonts
    return entry


# --- aggregation ----------------------------------------------------------------------------
def summarise(docs: list[dict]) -> dict:
    groups = {
        "docx": lambda e: e["kind"] == "docx",
        "pptx": lambda e: e["kind"] == "pptx",
        "pptx_office_exported": lambda e: e["kind"] == "pptx" and e["pdf_group"] == "office-exported",
        "pptx_other_producer": lambda e: e["kind"] == "pptx" and e["pdf_group"] == "other-producer",
        "office_exported": lambda e: e["pdf_group"] == "office-exported",
        "other_producer": lambda e: e["pdf_group"] == "other-producer",
    }
    out = {}
    for g, sel in groups.items():
        es = [e for e in docs if sel(e)]
        if not es:
            continue
        both = [e for e in es if all(e["stacks"].get(s, {}).get("status") == "ok" for s in STACKS)]
        gs = {"documents": len(es), "documents_all_renderers_ok": len(both)}
        for stack in STACKS:
            for subset_name, subset in (("all_rendered", es), ("common_subset", both)):
                ok = [e for e in subset if e["stacks"].get(stack, {}).get("status") == "ok"]
                sts = [e["stacks"][stack] for e in ok]
                gs.setdefault(stack, {})[subset_name] = {
                    "rendered": len(ok),
                    "failed": len([e for e in subset if e["stacks"].get(stack, {}).get("status") != "ok"]),
                    "page_count_exact": sum(1 for s in sts if s["page_diff"] == 0),
                    "abs_page_diff": M.describe([abs(s["page_diff"]) for s in sts]),
                    "ssim_median_per_doc": M.describe([s["ssim"]["median"] for s in sts]),
                    "ink_f1_median_per_doc": M.describe([s["ink_f1"]["median"] for s in sts]),
                    "words_same_page": M.describe([s["words"]["same_page"] for s in sts]),
                    "words_within_5pt": M.describe([s["words"]["within_5pt"] for s in sts]),
                    "words_within_20pt": M.describe([s["words"]["within_20pt"] for s in sts]),
                    "text_pages_aligned_share": M.describe([s["words"]["text_pages_aligned"] / s["words"]["text_pages"]
                                                            for s in sts if s["words"]["text_pages"]]),
                }
        gs["ssim_blank_page_median_per_doc"] = M.describe([e["ssim_blank_page"]["median"] for e in es])
        out[g] = gs
    return out
