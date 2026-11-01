"""Pure helpers of the visual section: read PDFs with Poppler, compare pages, match word positions, fonts.

Everything here reads PDFs with the Poppler command-line tools (``pdfinfo``, ``pdftoppm``, ``pdftotext -bbox``,
``pdftohtml -xml``, ``pdffonts``), the same tools for the Office PDF and for each renderer's PDF, so that no
renderer is read by its own reader.
"""

from __future__ import annotations

import difflib
import html
import re
import shutil
import statistics
import subprocess
import unicodedata
from collections import Counter
from pathlib import Path

import numpy as np
from PIL import Image
from skimage.metrics import structural_similarity

PDFINFO = shutil.which("pdfinfo") or "pdfinfo"
PDFTOPPM = shutil.which("pdftoppm") or "pdftoppm"
PDFTOTEXT = shutil.which("pdftotext") or "pdftotext"
PDFTOHTML = shutil.which("pdftohtml") or "pdftohtml"
PDFFONTS = shutil.which("pdffonts") or "pdffonts"


def _out(cmd, timeout=600) -> str:
    p = subprocess.run([str(c) for c in cmd], capture_output=True, text=True, timeout=timeout)
    if p.returncode != 0:
        raise RuntimeError(f"{cmd[0]} failed ({p.returncode}): {p.stderr[-500:]}")
    return p.stdout


# --- pages ---------------------------------------------------------------------------------
def page_sizes(pdf: Path) -> list[tuple[float, float]]:
    """(width, height) in points of every page (rotation applied)."""
    txt = _out([PDFINFO, "-f", "1", "-l", "100000", pdf])
    sizes, rots = {}, {}
    for m in re.finditer(r"^Page\s+(\d+) size:\s+([\d.]+) x ([\d.]+) pts", txt, re.M):
        sizes[int(m.group(1))] = (float(m.group(2)), float(m.group(3)))
    for m in re.finditer(r"^Page\s+(\d+) rot:\s+(\d+)", txt, re.M):
        rots[int(m.group(1))] = int(m.group(2))
    out = []
    for i in sorted(sizes):
        w, h = sizes[i]
        out.append((h, w) if rots.get(i, 0) in (90, 270) else (w, h))
    return out


def rasterize(pdf: Path, outdir: Path, dpi: int, prefix: str = "p") -> list[Path]:
    """Grayscale PNG of every page with pdftoppm, sorted by page."""
    outdir.mkdir(parents=True, exist_ok=True)
    _out([PDFTOPPM, "-r", str(dpi), "-gray", "-png", pdf, outdir / prefix], timeout=900)
    return sorted(outdir.glob(f"{prefix}-*.png"), key=lambda p: int(p.stem.rsplit("-", 1)[1]))


def load_gray(path: Path) -> np.ndarray:
    im = Image.open(path)
    if im.mode in ("RGBA", "LA", "P"):
        im = im.convert("RGBA")
        bg = Image.new("RGBA", im.size, (255, 255, 255, 255))
        im = Image.alpha_composite(bg, im)
    return np.asarray(im.convert("L"), dtype=np.uint8)


def _fit(img: np.ndarray, shape) -> np.ndarray:
    if img.shape == shape:
        return img
    return np.asarray(Image.fromarray(img).resize((shape[1], shape[0]), Image.LANCZOS), dtype=np.uint8)


def ssim(a: np.ndarray, b: np.ndarray) -> float:
    """Mean SSIM (Wang et al. 2004: Gaussian window sigma 1.5, K1 0.01, K2 0.03) of two grayscale pages;
    ``b`` is resampled to ``a``'s pixel size when they differ."""
    b = _fit(b, a.shape)
    return float(structural_similarity(a, b, data_range=255, gaussian_weights=True, sigma=1.5,
                                       use_sample_covariance=False))


INK_LEVEL = 160  # a pixel darker than this grey is ink


def ink(a: np.ndarray) -> float:
    """Share of pixels darker than mid-grey (a crude measure of how much is on the page)."""
    return float((a < INK_LEVEL).mean())


def _disk(r: int) -> np.ndarray:
    y, x = np.ogrid[-r:r + 1, -r:r + 1]
    return (x * x + y * y) <= r * r


def ink_f1(a: np.ndarray, b: np.ndarray, radius: int = 3):
    """Ink agreement tolerant to small shifts: recall = share of ``a``'s ink pixels with ink of ``b`` within
    ``radius`` pixels, precision the same the other way, and their F1. A blank ``b`` scores 0; two empty pages
    score None."""
    from scipy.ndimage import binary_dilation

    b = _fit(b, a.shape)
    ia, ib = a < INK_LEVEL, b < INK_LEVEL
    na, nb = int(ia.sum()), int(ib.sum())
    if na == 0 and nb == 0:
        return None
    if na == 0 or nb == 0:
        return 0.0
    k = _disk(radius)
    recall = float((ia & binary_dilation(ib, k)).sum()) / na
    precision = float((ib & binary_dilation(ia, k)).sum()) / nb
    return 0.0 if recall + precision == 0 else 2 * recall * precision / (recall + precision)


# --- words ---------------------------------------------------------------------------------
_WORD = re.compile(r'<word xMin="([\d.\-]+)" yMin="([\d.\-]+)" xMax="([\d.\-]+)" yMax="([\d.\-]+)">(.*?)</word>')
_PAGE = re.compile(r'<page width="([\d.]+)" height="([\d.]+)">')
_QUOTES = str.maketrans({"\u2018": "'", "\u2019": "'", "\u201c": '"', "\u201d": '"', "\u2013": "-", "\u2014": "-",
                         "\u00a0": " ", "\u2010": "-", "\u2011": "-"})


def norm_token(t: str) -> str:
    """A word as compared: NFKC (ligatures split), quotes and dashes unified, lower case, outer punctuation
    stripped. Tokens with no letter or digit (bullets, rules, symbols) become ''."""
    t = unicodedata.normalize("NFKC", html.unescape(t)).translate(_QUOTES).lower()
    t = t.strip(".,;:!?()[]{}\"'*-/\\|<>")
    if not any(ch.isalnum() for ch in t):
        return ""
    return t


def words(pdf: Path):
    """Word boxes of a PDF: (pages [(w, h)], words [(page, x0, y0, x1, y1, token)]) in pdftotext's reading order,
    tokens normalised, empty tokens dropped."""
    txt = _out([PDFTOTEXT, "-bbox", "-enc", "UTF-8", pdf, "-"], timeout=900)
    pages, ws = [], []
    pos = 0
    page_starts = [(m.start(), float(m.group(1)), float(m.group(2))) for m in _PAGE.finditer(txt)]
    page_starts.append((len(txt), 0, 0))
    for k in range(len(page_starts) - 1):
        start, w, h = page_starts[k]
        end = page_starts[k + 1][0]
        pages.append((w, h))
        for m in _WORD.finditer(txt, start, end):
            tok = norm_token(m.group(5))
            if tok:
                ws.append((k, float(m.group(1)), float(m.group(2)), float(m.group(3)), float(m.group(4)), tok))
        pos = end
    return pages, ws


def actual_text_spans(pdf: Path) -> int:
    """Number of marked-content ``/ActualText`` entries in the PDF's (Flate) content streams."""
    import zlib

    data = Path(pdf).read_bytes()
    n = 0
    for m in re.finditer(rb"stream\r?\n", data):
        end = data.find(b"endstream", m.end())
        try:
            n += zlib.decompress(data[m.end():end]).count(b"/ActualText")
        except Exception:
            continue
    return n


def drop_actual_text(pdf: Path, out: Path) -> Path:
    """Re-write a PDF through poppler's cairo backend (``pdftocairo -pdf``), which keeps every glyph where it is
    and its ToUnicode text but drops marked-content ``/ActualText``. Used when a PDF puts a whole line's text on
    its first word (rpptx): pdftotext then gives each word its own box again."""
    _out([shutil.which("pdftocairo") or "pdftocairo", "-pdf", pdf, out], timeout=900)
    return out


def match_words(office, other, tolerances=(5.0, 20.0), min_page_words=5, aligned_share=0.9):
    """Where the Office PDF's words land in another PDF.

    The two word streams (all pages, pdftotext reading order) are aligned with difflib's SequenceMatcher on the
    normalised tokens, so that a word is paired with the same word in the same context, wherever it moved. For
    each Office word: ``matched`` (found in the other PDF), ``same_page`` (paired word on the same page index),
    ``within_T`` (same page and both the left edge and the vertical centre of the box within T points, after
    scaling the other page to the Office page's size). A page with at least ``min_page_words`` words is
    ``aligned`` when at least ``aligned_share`` of its words are on the same page in the other PDF (its page
    breaks agree).
    """
    (opages, ow), (rpages, rw) = office, other
    a = [w[5] for w in ow]
    b = [w[5] for w in rw]
    sm = difflib.SequenceMatcher(None, a, b, autojunk=True)
    pair = {}
    for blk in sm.get_matching_blocks():
        for k in range(blk.size):
            pair[blk.a + k] = blk.b + k
    n = len(ow)
    per_page_total, per_page_same = Counter(), Counter()
    same = 0
    within = {t: 0 for t in tolerances}
    dxs, dys = [], []
    for i, w in enumerate(ow):
        per_page_total[w[0]] += 1
        j = pair.get(i)
        if j is None:
            continue
        v = rw[j]
        if v[0] != w[0]:
            continue
        same += 1
        per_page_same[w[0]] += 1
        ow_, oh_ = opages[w[0]]
        rw_, rh_ = rpages[v[0]] if v[0] < len(rpages) else (ow_, oh_)
        sx = ow_ / rw_ if rw_ else 1.0
        sy = oh_ / rh_ if rh_ else 1.0
        sdx = v[1] * sx - w[1]
        sdy = (v[2] + v[4]) / 2 * sy - (w[2] + w[4]) / 2
        dxs.append(sdx)
        dys.append(sdy)
        dx, dy = abs(sdx), abs(sdy)
        for t in tolerances:
            if dx <= t and dy <= t:
                within[t] += 1
    text_pages = [p for p, c in per_page_total.items() if c >= min_page_words]
    aligned = [p for p in text_pages if per_page_same[p] >= aligned_share * per_page_total[p]]
    res = {
        "words_office": n,
        "words_other": len(rw),
        "matched": round(len(pair) / n, 4) if n else None,
        "same_page": round(same / n, 4) if n else None,
    }
    for t in tolerances:
        res[f"within_{int(t)}pt"] = round(within[t] / n, 4) if n else None
    # signed median offset of the words found on the same page (other minus Office, points; y grows downwards)
    res["median_dx_pt"] = round(statistics.median(dxs), 2) if dxs else None
    res["median_dy_pt"] = round(statistics.median(dys), 2) if dys else None
    res["text_pages"] = len(text_pages)
    res["text_pages_aligned"] = len(aligned)
    res["per_page_same_page"] = {int(p): round(per_page_same[p] / per_page_total[p], 3) for p in sorted(per_page_total)}
    return res


# --- fonts ---------------------------------------------------------------------------------
_STYLE = re.compile(r"(bolditalic|boldoblique|bold|italic|oblique|regular|roman|book|normal)$")
# metric-compatible clones (same advance widths) that the benchmark or the system provides
CLONES = {
    "calibri": {"carlito"},
    "cambria": {"caladea"},
    "arial": {"liberationsans", "arimo"},
    "helvetica": {"liberationsans", "arimo"},
    "timesnewroman": {"liberationserif", "tinos"},
    "times": {"liberationserif", "tinos"},
    "couriernew": {"liberationmono", "cousine"},
    "arialnarrow": {"liberationsansnarrow"},
}
SYMBOL_FONTS = {"symbol", "wingdings", "wingdings2", "wingdings3", "opensymbol", "standardsymbolsps", "webdings"}


def canon_font(name: str) -> str:
    """A PDF font name reduced to a family key: subset tag, style suffix and PostScript MT/PS tags removed,
    lower case, letters and digits only (``ABCDEE+TimesNewRomanPS-BoldMT`` -> ``timesnewroman``,
    ``Calibri Light`` and ``Calibri-Light`` -> ``calibrilight``)."""
    name = re.sub(r"^[A-Z]{6}\+", "", name.strip())
    name = name.split(",")[0]
    parts = name.split("-")
    base = parts[0]
    rest = "".join(parts[1:])
    rest_key = re.sub(r"[^a-z0-9]", "", rest.lower())
    rest_key = re.sub(r"(psmt|mt|ps)$", "", rest_key)
    rest_key = _STYLE.sub("", rest_key)
    rest_key = _STYLE.sub("", rest_key)
    key = re.sub(r"[^a-z0-9]", "", base.lower())
    key = re.sub(r"(psmt|mt)$", "", key)
    key = re.sub(r"ps$", "", key) if key.endswith("romanps") else key
    key = _STYLE.sub("", key) if len(key) > 6 else key
    return key + rest_key


_FONTSPEC = re.compile(r'<fontspec id="(\d+)" size="[^"]*" family="([^"]*)"')
_TEXT = re.compile(r'<text [^>]*font="(\d+)"[^>]*>(.*?)</text>', re.S)


def font_chars(pdf: Path) -> Counter:
    """Characters (letters and digits) drawn in each font family of a PDF (pdftohtml -xml)."""
    xml = _out([PDFTOHTML, "-xml", "-i", "-stdout", "-q", pdf], timeout=900)
    spec = {m.group(1): canon_font(html.unescape(m.group(2))) for m in _FONTSPEC.finditer(xml)}
    counts = Counter()
    for m in _TEXT.finditer(xml):
        txt = html.unescape(re.sub(r"<[^>]+>", "", m.group(2)))
        n = sum(ch.isalnum() for ch in txt)
        if n:
            counts[spec.get(m.group(1), "?")] += n
    return counts


def pdf_fonts(pdf: Path) -> set[str]:
    """Font families a PDF carries (pdffonts), as family keys."""
    txt = _out([PDFFONTS, pdf])
    out = set()
    for line in txt.splitlines()[2:]:
        if line.strip():
            out.add(canon_font(line.split()[0]))
    return out


def font_available(office_family: str, renderer_fonts: set[str]) -> bool:
    """The renderer drew this Office family with the family itself or a metric-compatible clone."""
    return office_family in renderer_fonts or bool(CLONES.get(office_family, set()) & renderer_fonts)


def describe(values):
    vals = [v for v in values if v is not None]
    if not vals:
        return {"n": 0, "median": None, "mean": None, "min": None, "max": None}
    return {
        "n": len(vals),
        "median": round(statistics.median(vals), 4),
        "mean": round(statistics.fmean(vals), 4),
        "min": round(min(vals), 4),
        "max": round(max(vals), 4),
    }
