"""Visual section: how an agent gets visual feedback on a document, and how faithful that picture is.

Regenerates ``results/visual.json`` and ``visual/figures/`` from scratch:

1. fidelity (``fidelity.py``): every real file with a PDF twin made by Office, rendered by rdocx/rpptx and by
   LibreOffice (the built-in skills' route), compared with the Office PDF: page counts, SSIM and ink agreement of
   index-paired pages, where the Office PDF's words land (same page, within 5 / 20 points), fonts;
2. capabilities (``capabilities.py``): the feedback loop of each stack, every row tested on the corpus;
3. figures (``figures.py``): labelled side-by-side pages "Office PDF | rdocx | LibreOffice", tracked changes,
   thumbnail grids.

4. rdocx findings (``findings.py``): each rendering fault found above, reduced to a neutral reproduction that
   builds its own input, re-run on the pinned build (``reproduced``).

Usage: .cache/venv/bin/python visual/run.py [--jobs N] [--only fidelity,capabilities,findings,figures]
(``--only`` reuses the other parts from .cache/work/visual/partial.json)
"""

from __future__ import annotations

import argparse
import json
import sys
import time
from concurrent.futures import ProcessPoolExecutor, ThreadPoolExecutor
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parent))
sys.path.insert(0, str(HERE))
import common  # noqa: E402
import fidelity  # noqa: E402

CACHE = common.WORK / "visual" / "partial.json"
ALL = ("fidelity", "capabilities", "findings", "figures")


def _measure(args):
    f, rend = args
    return fidelity.measure(f, rend)


def run_fidelity(jobs: int) -> dict:
    files = fidelity.twins()
    t0 = time.time()
    with ThreadPoolExecutor(jobs) as ex:
        rends = list(ex.map(fidelity.render, files))
    print(f"fidelity: rendered {len(files)} files in {time.time() - t0:.0f} s", flush=True)
    t0 = time.time()
    with ProcessPoolExecutor(jobs) as ex:
        docs = list(ex.map(_measure, zip(files, rends)))
    print(f"fidelity: measured in {time.time() - t0:.0f} s", flush=True)
    return {"fidelity": docs, "fidelity_summary": fidelity.summarise(docs)}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--jobs", type=int, default=8)
    ap.add_argument("--only", default=",".join(ALL))
    a = ap.parse_args()
    parts = set(a.only.split(","))
    CACHE.parent.mkdir(parents=True, exist_ok=True)
    data = json.loads(CACHE.read_text()) if CACHE.exists() and parts != set(ALL) else {}
    if "fidelity" in parts:
        data.update(run_fidelity(a.jobs))
        CACHE.write_text(json.dumps(data, indent=1))
    if "capabilities" in parts:
        import capabilities
        cap = capabilities.run()
        data["capabilities"] = cap.pop("rows")
        data["capabilities_detail"] = cap
        CACHE.write_text(json.dumps(data, indent=1))
    if "findings" in parts:
        import findings
        data["rdocx_findings"] = findings.run()
        CACHE.write_text(json.dumps(data, indent=1))
    if "figures" in parts:
        import figures
        data["figures"] = figures.make(data)
        CACHE.write_text(json.dumps(data, indent=1))
    out = {"section": "visual", "versions": common.versions(), **data}
    out["method"] = {
        "dpi": fidelity.DPI,
        "ssim": "mean SSIM (skimage, Gaussian window sigma 1.5) of grayscale pages rasterised by pdftoppm at the "
                "same DPI, pages paired by index; ssim_blank_page is the same score of each Office page against "
                "a white page, the floor to read SSIM against",
        "ink_f1": "share of ink pixels (grey < 160) of each page with ink of the other page within 3 px (3.6 pt at "
                  "60 dpi), both ways, F1; a white page scores 0",
        "words": "pdftotext -bbox words of the Office PDF matched to the renderer's (difflib on normalised tokens, "
                 "all pages in reading order); same_page: share of the Office words found on the same page index; "
                 "within_Tpt: same page and left edge and vertical centre within T points; text_pages_aligned: "
                 "pages (>= 5 words) with >= 90% of their words on the same page",
        "fonts": "share of the Office PDF's letters and digits per font family (pdftohtml -xml); a family is "
                 "missing for a renderer when its PDF (pdffonts) holds neither the family nor a metric-compatible "
                 "clone (Carlito, Caladea, Liberation / Arimo, Tinos, Cousine)",
    }
    p = common.save_results("visual", out)
    print(f"wrote {p}")


if __name__ == "__main__":
    main()
