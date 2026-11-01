"""Part 3 of the visual section: labelled side-by-side figures, made with Pillow.

Pages are picked from the fidelity results (``fidelity.py``) by rule, not by eye: among the pages of the Word
reports that both renderers rendered, with at least 80 words, not the first page (covers carry logos) and no
picture in the Office PDF (no third-party image), the page where each renderer is closest to Office (highest SSIM)
and farthest (lowest SSIM). Only GOV.UK files under the Open Government Licence and synthetic files are used
(attribution in NOTES.md). Each PNG is kept under ``MAX_BYTES``.
"""

from __future__ import annotations

import io
import json
import re
import shutil
import subprocess
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

import common
import metrics as M

HERE = Path(__file__).resolve().parent
OUT = HERE / "figures"
WORK = common.WORK / "visual" / "figures"
MAX_BYTES = 290_000
FIG_DPI = 72
LABEL_FONT = common.FONTS_DIR / "Carlito-Bold.ttf"
TEXT_FONT = common.FONTS_DIR / "Carlito-Regular.ttf"
STACK_LABEL = {"office": "Office PDF (ground truth)", "rdocx": "rdocx / rpptx", "libreoffice": "LibreOffice (built-in skills)"}


def font(size, bold=True):
    return ImageFont.truetype(str(LABEL_FONT if bold else TEXT_FONT), size)


def raster_page(pdf: Path, page: int, dpi=FIG_DPI, gray=True) -> Image.Image:
    WORK.mkdir(parents=True, exist_ok=True)
    prefix = WORK / "tmp"
    for f in WORK.glob("tmp-*"):
        f.unlink()
    cmd = ["pdftoppm", "-r", str(dpi), "-f", str(page), "-l", str(page), "-png", str(pdf), str(prefix)]
    if gray:
        cmd.insert(1, "-gray")
    subprocess.run(cmd, check=True, capture_output=True)
    f = next(WORK.glob("tmp-*.png"))
    im = Image.open(f).convert("L" if gray else "RGB")
    im.load()
    return im


def compose(panels, labels, title, captions=None, gap=14, pad=12):
    """Panels side by side, a label above each and a title line on top."""
    mode = "RGB" if any(p.mode == "RGB" for p in panels) else "L"
    panels = [p.convert(mode) for p in panels]
    h = max(p.height for p in panels)
    w = sum(p.width for p in panels) + gap * (len(panels) - 1) + 2 * pad
    top = 58 if captions else 46
    im = Image.new(mode, (w, h + top + pad), "white" if mode == "RGB" else 255)
    dr = ImageDraw.Draw(im)
    dr.text((pad, 6), title, font=font(17), fill="black" if mode == "RGB" else 0)
    x = pad
    for k, p in enumerate(panels):
        dr.text((x, 28), labels[k], font=font(14), fill=(30, 30, 30) if mode == "RGB" else 30)
        if captions and captions[k]:
            dr.text((x, 44), captions[k], font=font(11, bold=False), fill=(80, 80, 80) if mode == "RGB" else 80)
        im.paste(p, (x, top))
        dr.rectangle([x - 1, top - 1, x + p.width, top + p.height], outline=(150, 150, 150) if mode == "RGB" else 150)
        x += p.width + gap
    return im


def save_small(im: Image.Image, path: Path, colors=32) -> int:
    """PNG, palette-quantised, scaled down until it is under MAX_BYTES."""
    path.parent.mkdir(parents=True, exist_ok=True)
    scale = 1.0
    while True:
        cur = im if scale == 1.0 else im.resize((int(im.width * scale), int(im.height * scale)), Image.LANCZOS)
        q = cur.convert("RGB").quantize(colors=colors, method=Image.Quantize.MEDIANCUT)
        buf = io.BytesIO()
        q.save(buf, "PNG", optimize=True)
        if buf.tell() <= MAX_BYTES or scale < 0.4:
            path.write_bytes(buf.getvalue())
            return buf.tell()
        scale *= 0.9


def has_images(pdf: Path, page: int) -> bool:
    r = subprocess.run(["pdfimages", "-list", "-f", str(page), "-l", str(page), str(pdf)], capture_output=True, text=True)
    return len(r.stdout.strip().splitlines()) > 2


def candidates(doc: dict, kind="docx", min_words=80):
    pp = doc["per_page"]
    office_pdf = common.BENCH / doc["office_pdf"]
    out = []
    n = min(len(pp.get("rdocx_ssim", [])), len(pp.get("libreoffice_ssim", [])))
    for k in range(1, n):
        if pp["office_words"][k] < min_words or has_images(office_pdf, k + 1):
            continue
        out.append(k)
    return out


def page_figure(doc, k, title, path, gray=True, dpi=FIG_DPI):
    d = common.WORK / "visual" / "fidelity" / doc["id"]
    pdfs = {"office": common.BENCH / doc["office_pdf"], "rdocx": d / "rdocx.pdf", "libreoffice": d / "libreoffice.pdf"}
    panels = [raster_page(pdfs[s], k + 1, dpi=dpi, gray=gray) for s in ("office", "rdocx", "libreoffice")]
    pp = doc["per_page"]
    caps = ["page %d of %d" % (k + 1, doc["pages"]["office"])]
    for s in ("rdocx", "libreoffice"):
        sp = pp[f"{s}_same_page"][k]
        caps.append("SSIM %.2f, ink F1 %.2f, words on this page %s" % (
            pp[f"{s}_ssim"][k], pp[f"{s}_ink_f1"][k] or 0, "-" if sp is None else f"{sp:.0%}"))
    im = compose(panels, [STACK_LABEL[s] for s in ("office", "rdocx", "libreoffice")], title, caps)
    size = save_small(im, path)
    return {"path": str(path.relative_to(common.BENCH)), "bytes": size, "document": doc["id"], "page": k + 1,
            "ssim": {s: pp[f"{s}_ssim"][k] for s in ("rdocx", "libreoffice")},
            "ink_f1": {s: pp[f"{s}_ink_f1"][k] for s in ("rdocx", "libreoffice")},
            "words_same_page": {s: pp[f"{s}_same_page"][k] for s in ("rdocx", "libreoffice")}}


# --- tracked changes -------------------------------------------------------------------------------------
def locate(pdf: Path, phrase: str):
    """(page, y-centre) of the first place the phrase's first words occur in a row (pdftotext -bbox)."""
    toks = [M.norm_token(t) for t in phrase.split()]
    toks = [t for t in toks if t][:5]
    pages, ws = M.words(pdf)
    seq = [w[5] for w in ws]
    for i in range(len(seq) - len(toks) + 1):
        if seq[i:i + len(toks)] == toks:
            w = ws[i]
            return w[0] + 1, (w[2] + w[4]) / 2, pages[w[0]]
    return None


def tracked_figure(ctx: dict, path: Path):
    wins = ctx["windows"]
    pdfs = [("rdocx --revision-view tracked", Path(ctx["tracked_pdfs"]["tracked"])),
            ("rdocx --revision-view tracked --revision-style legal", Path(ctx["tracked_pdfs"]["tracked-legal"])),
            ("LibreOffice, default PDF export", Path(ctx["lo_pdf"]))]
    # the first replacement (a deletion followed by an insertion) found in all three PDFs, so that both marks show
    for w in sorted(wins, key=lambda w: not w.get("replacement")):
        locs = [locate(p, w["tracked"]) for _, p in pdfs]
        if all(locs):
            break
    else:
        return None
    dpi = 110
    panels = []
    for (label, pdf), (page, y, (pw, ph)) in zip(pdfs, locs):
        im = raster_page(pdf, page, dpi=dpi, gray=False)
        s = dpi / 72
        top = max(0, int((y - 70) * s))
        bottom = min(im.height, int((y + 70) * s))
        panels.append((label, im.crop((0, top, int(pw * s), bottom)), page))
    width = max(p.width for _, p, _ in panels)
    rows = []
    for label, p, page in panels:
        cap = Image.new("RGB", (width, 22), "white")
        ImageDraw.Draw(cap).text((4, 3), f"{label} (page {page})", font=font(14), fill="black")
        rows += [cap, p]
    h = sum(r.height for r in rows) + 8 * len(rows) + 30
    im = Image.new("RGB", (width + 16, h), "white")
    ImageDraw.Draw(im).text((8, 6), "Tracked changes in the render: docx-M (synthetic), the same revision in each view",
                            font=font(16), fill="black")
    yy = 30
    for r in rows:
        im.paste(r, (8, yy))
        yy += r.height + 8
    size = save_small(im, path, colors=64)
    return {"path": str(path.relative_to(common.BENCH)), "bytes": size, "document": "docx-M",
            "deleted_text": w["deleted"], "pages": [p for _, _, p in panels]}


def grid_figure(cap_detail: dict, path: Path):
    rows = {r["stack"]: r for r in cap_detail if r["capability"] == "thumbnail grid" and r["kind"] == "pptx"}
    b = rows.get("builtin", {}).get("files") or []
    r = rows.get("rdocx", {}).get("files") or []
    if not b or not r:
        return None
    ib = Image.open(b[0]).convert("RGB")
    ir = Image.open(r[0]).convert("RGB")
    h = 520
    ib = ib.resize((int(ib.width * h / ib.height), h), Image.LANCZOS)
    ir = ir.resize((int(ir.width * h / ir.height), h), Image.LANCZOS)
    im = compose([ib, ir], ["built-in: scripts/thumbnail.py (first grid)", "rpptx: render_all_handouts(9) (first page)"],
                 "Thumbnail grids of pptx-M (synthetic, 60 slides)")
    size = save_small(im, path, colors=64)
    return {"path": str(path.relative_to(common.BENCH)), "bytes": size, "document": "pptx-M"}


def make(data: dict) -> list[dict]:
    if OUT.exists():
        shutil.rmtree(OUT)
    OUT.mkdir(parents=True)
    figs = []
    docs = [d for d in data["fidelity"] if all(d["stacks"].get(s, {}).get("status") == "ok" for s in ("rdocx", "libreoffice"))]
    docx = [(d, k) for d in docs if d["kind"] == "docx" for k in candidates(d)]
    used = set()

    def pick(stack, best):
        key = lambda dk: dk[0]["per_page"][f"{stack}_ssim"][dk[1]]
        for d, k in sorted(docx, key=key, reverse=best):
            if (d["id"], k) not in used:
                used.add((d["id"], k))
                return d, k
        return None

    plan = [("rdocx", True, "fig1_docx_rdocx_closest.png", "Word page where rdocx is closest to Office (highest SSIM)"),
            ("rdocx", False, "fig2_docx_rdocx_farthest.png", "Word page where rdocx is farthest from Office (lowest SSIM)"),
            ("libreoffice", True, "fig3_docx_libreoffice_closest.png", "Word page where LibreOffice is closest to Office"),
            ("libreoffice", False, "fig4_docx_libreoffice_farthest.png", "Word page where LibreOffice is farthest from Office")]
    for stack, best, name, title in plan:
        got = pick(stack, best)
        if got:
            d, k = got
            f = page_figure(d, k, f"{title}: {d['id']}, page {k + 1}", OUT / name)
            f["caption"] = title
            figs.append(f)
    # one slide: the office-native deck, the slide with rpptx's median SSIM among slides without pictures
    deck = next((d for d in docs if d["kind"] == "pptx" and d["route"] == "office-native"), None)
    if deck:
        ks = candidates(deck, min_words=20)
        ks.sort(key=lambda k: deck["per_page"]["rdocx_ssim"][k])
        if ks:
            k = ks[len(ks) // 2]
            f = page_figure(deck, k, f"A slide at rpptx's median SSIM: {deck['id']}, slide {k + 1}",
                            OUT / "fig5_pptx_slide_median.png", gray=False, dpi=60)
            f["caption"] = "Slide where rpptx scores its median SSIM (PowerPoint's own PDF)"
            figs.append(f)
    ctx_path = common.WORK / "visual" / "capabilities" / "context.json"
    if ctx_path.exists():
        f = tracked_figure(json.loads(ctx_path.read_text())["review"], OUT / "fig6_tracked_changes.png")
        if f:
            f["caption"] = "The same tracked deletion in rdocx's tracked view (plain, legal) and LibreOffice's default export"
            figs.append(f)
    if "capabilities" in data:
        f = grid_figure(data["capabilities"], OUT / "fig7_thumbnail_grids.png")
        if f:
            f["caption"] = "Thumbnail grid of each stack for a 60-slide deck"
            figs.append(f)
    return figs
