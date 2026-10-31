"""R11 oracle: OUT/accepted and OUT/tracked each hold three PNG pages of a Letter page at 100 dpi (850 x 1100,
2 px tolerance), none blank; at least one tracked page differs from its accepted page (the marks are visible)."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import review_oracle as o  # noqa: E402


def pages(folder):
    return sorted(Path(folder).glob("*.png"))


def main(out):
    acc, trk = pages(Path(out) / "accepted"), pages(Path(out) / "tracked")
    info = {p.parent.name + "/" + p.name: o.png_pixels(p)[:3] for p in acc + trk}
    size_ok = all(abs(w - 850) <= 2 and abs(h - 1100) <= 2 for w, h, _ in info.values())
    diffs = [o.png_differs(a, t) for a, t in zip(acc, trk)] if len(acc) == len(trk) else []
    o.report({
        "accepted_three_pages": len(acc) == 3,
        "tracked_three_pages": len(trk) == 3,
        "100_dpi_letter": bool(info) and size_ok,
        "not_blank": bool(info) and all(ink > 0.005 for _, _, ink in info.values()),
        "tracked_marks_visible": any(x > 0.001 for x in diffs),
    }, {"pages": {k: [w, h, round(ink, 4)] for k, (w, h, ink) in info.items()}, "row_diff": diffs})


o.guard(lambda: main(sys.argv[1]))
