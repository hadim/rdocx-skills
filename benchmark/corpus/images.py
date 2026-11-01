"""Deterministic PNG pictures for the synthetic corpus, drawn with Pillow.

Most are chart-like figures (bars, lines, areas, a heat map, scatter) with axis labels in
Carlito; a few are photo-like (smooth gradient plus noise) so that the picture sizes spread
from tens of KB to a few hundred KB, as in real reports. Same index, same bytes.
"""

from __future__ import annotations

import math
import random
from pathlib import Path

from PIL import Image, ImageDraw, ImageFilter, ImageFont

PALETTE = [(31, 78, 121), (84, 130, 53), (191, 144, 0), (197, 90, 17), (112, 48, 160), (38, 166, 154)]
KINDS = ["bar", "line", "area", "heat", "scatter", "photo", "bar", "line"]


def _font(font_dir: Path, size: int, bold: bool = False):
    name = "Carlito-Bold.ttf" if bold else "Carlito-Regular.ttf"
    p = font_dir / name
    return ImageFont.truetype(str(p), size) if p.exists() else ImageFont.load_default()


def _axes(d, box, rng, font, ymax, ylabel_fmt="{:,.0f}"):
    x0, y0, x1, y1 = box
    d.line([(x0, y0), (x0, y1), (x1, y1)], fill=(60, 60, 60), width=3)
    for k in range(6):
        y = y1 - (y1 - y0) * k / 5
        d.line([(x0, y), (x1, y)], fill=(225, 225, 225), width=1)
        d.text((x0 - 12, y), ylabel_fmt.format(ymax * k / 5), fill=(70, 70, 70), font=font, anchor="rm")


def make_image(index: int, out: Path, font_dir: Path, width: int = 1500, height: int = 900) -> dict:
    rng = random.Random(1000 + index)
    kind = KINDS[index % len(KINDS)]
    img = Image.new("RGB", (width, height), (255, 255, 255))
    d = ImageDraw.Draw(img)
    f_title, f_lab, f_small = _font(font_dir, 44, True), _font(font_dir, 28), _font(font_dir, 24)
    box = (170, 140, width - 80, height - 120)
    x0, y0, x1, y1 = box
    title = {
        "bar": "Capital spend by district",
        "line": "Monthly demand against forecast",
        "area": "Supply mix by source",
        "heat": "Burst frequency by month and zone",
        "scatter": "Pressure against age of main",
        "photo": "Site survey",
    }[kind]
    if kind == "photo":
        # Smooth two-colour gradient, soft shapes, then per-pixel noise: compresses like a photo.
        base = Image.linear_gradient("L").resize((width, height))
        c1 = PALETTE[rng.randrange(len(PALETTE))]
        c2 = (230, 225, 210)
        img = Image.composite(Image.new("RGB", (width, height), c1), Image.new("RGB", (width, height), c2), base)
        d = ImageDraw.Draw(img)
        for _ in range(14):
            cx, cy, r = rng.randrange(width), rng.randrange(height), rng.randrange(40, 260)
            col = tuple(min(255, max(0, v + rng.randrange(-60, 60))) for v in c1)
            d.ellipse([cx - r, cy - r, cx + r, cy + r], fill=col)
        img = img.filter(ImageFilter.GaussianBlur(18))
        noise = Image.effect_noise((width // 2, height // 2), 18).resize((width, height)).convert("RGB")
        img = Image.blend(img, noise, 0.08)
        img = img.resize((width // 2, height // 2))
    elif kind == "bar":
        n = rng.randint(6, 10)
        vals = [rng.uniform(2, 30) for _ in range(n)]
        ymax = math.ceil(max(vals) / 5) * 5
        _axes(d, box, rng, f_small, ymax, "{:.0f}")
        bw = (x1 - x0) / n
        for k, v in enumerate(vals):
            bx = x0 + k * bw + bw * 0.18
            by = y1 - (y1 - y0) * v / ymax
            d.rectangle([bx, by, bx + bw * 0.64, y1], fill=PALETTE[0])
            d.text((bx + bw * 0.32, by - 8), f"{v:.1f}", fill=(40, 40, 40), font=f_small, anchor="mb")
            d.text((bx + bw * 0.32, y1 + 14), f"D{k + 1}", fill=(70, 70, 70), font=f_small, anchor="mt")
    elif kind in ("line", "area"):
        n = 24
        series = []
        for s in range(3 if kind == "area" else 2):
            level = rng.uniform(20, 60)
            pts = []
            for k in range(n):
                level = max(5, level + rng.uniform(-6, 6) + 8 * math.sin(k / 3.0 + s))
                pts.append(level)
            series.append(pts)
        if kind == "area":
            stacked = [[sum(series[j][k] for j in range(s + 1)) for k in range(n)] for s in range(len(series))]
            ymax = math.ceil(max(stacked[-1]) / 20) * 20
        else:
            ymax = math.ceil(max(max(p) for p in series) / 20) * 20
        _axes(d, box, rng, f_small, ymax)
        sx = (x1 - x0) / (n - 1)
        if kind == "area":
            prev = [y1] * n
            for s, pts in enumerate(stacked):
                cur = [y1 - (y1 - y0) * v / ymax for v in pts]
                poly = [(x0 + k * sx, cur[k]) for k in range(n)] + [(x0 + k * sx, prev[k]) for k in reversed(range(n))]
                d.polygon(poly, fill=PALETTE[s + 1])
                prev = cur
        else:
            for s, pts in enumerate(series):
                xy = [(x0 + k * sx, y1 - (y1 - y0) * v / ymax) for k, v in enumerate(pts)]
                d.line(xy, fill=PALETTE[s], width=5, joint="curve")
        for k in range(0, n, 3):
            d.text((x0 + k * sx, y1 + 14), f"M{k + 1}", fill=(70, 70, 70), font=f_small, anchor="mt")
    elif kind == "heat":
        rows, cols = 8, 12
        cw, ch = (x1 - x0) / cols, (y1 - y0) / rows
        for r in range(rows):
            for c in range(cols):
                v = rng.random()
                col = tuple(int(255 - v * (255 - p)) for p in PALETTE[3])
                d.rectangle([x0 + c * cw, y0 + r * ch, x0 + (c + 1) * cw - 2, y0 + (r + 1) * ch - 2], fill=col)
                d.text((x0 + (c + 0.5) * cw, y0 + (r + 0.5) * ch), f"{int(v * 40)}", fill=(30, 30, 30), font=f_small, anchor="mm")
            d.text((x0 - 12, y0 + (r + 0.5) * ch), f"Zone {r + 1}", fill=(70, 70, 70), font=f_small, anchor="rm")
    else:  # scatter
        ymax = 100
        _axes(d, box, rng, f_small, ymax)
        for _ in range(220):
            age = rng.uniform(0, 1)
            v = max(2, min(98, 80 - 50 * age + rng.gauss(0, 9)))
            px, py = x0 + (x1 - x0) * age, y1 - (y1 - y0) * v / ymax
            d.ellipse([px - 7, py - 7, px + 7, py + 7], fill=PALETTE[4])
    if kind != "photo":
        d.text((x0, 50), f"{title} (fig. {index + 1})", fill=(20, 20, 20), font=f_title)
        d.text((width - 80, height - 40), "Source: synthetic data", fill=(110, 110, 110), font=f_small, anchor="rm")
    out.parent.mkdir(parents=True, exist_ok=True)
    img.save(out, format="PNG", optimize=False)
    return {"path": str(out), "kind": kind, "w": img.width, "h": img.height}


def make_pool(n: int, out_dir: Path, font_dir: Path) -> list[dict]:
    return [make_image(i, out_dir / f"img_{i:03d}.png", font_dir) for i in range(n)]
