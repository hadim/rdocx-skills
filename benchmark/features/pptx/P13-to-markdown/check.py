"""P13 oracle: reads the Markdown as CommonMark text (entities and backslash escapes decoded, inline HTML and
emphasis markers dropped): every title as a heading, in slide order; the slide-2 bullets as list items with the
level-2 bullet indented under its parent; the slide-3 bullets as list items; the table as a GFM table (header,
delimiter row, the marker cell); each slide's notes inside that slide's section."""
import html
import json
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "lib"))
import oracle  # noqa: E402


def plain(line):
    line = re.sub(r"<[^>]+>", "", line)
    line = re.sub(r"\\([!-/:-@\[-`{-~])", r"\1", line)
    line = html.unescape(line)
    return line.replace("**", "").replace("__", "")


def check(out, inputs):
    truth = json.loads((inputs / "truth.json").read_text())
    lines = [plain(x) for x in out.read_text().splitlines()]
    errors = []
    heads = []
    for k, title in enumerate(truth["titles"]):
        at = next((i for i, x in enumerate(lines) if re.match(r"^#{1,6}\s+", x)
                   and re.sub(r"^#{1,6}\s+", "", x).strip().strip("*_") == title), None)
        if at is None:
            errors.append(f"no heading for slide {k + 1} title {title!r}")
        heads.append(at)
    if any(h is None for h in heads):
        return errors
    if heads != sorted(heads):
        errors.append("headings out of slide order")
    sections = [lines[h:(heads[k + 1] if k + 1 < len(heads) else len(lines))] for k, h in enumerate(heads)]

    def item(sec, text):
        for x in sec:
            m = re.match(r"^(\s*)[-*+]\s+(.*)$", x)
            if m and text in m.group(2):
                return len(m.group(1).expandtabs(4))
        return None

    parent, child = item(sections[1], "mk-body-2"), item(sections[1], "mk-sub-bullet-2")
    if parent is None or child is None:
        errors.append(f"slide 2 bullets not list items (parent {parent}, child {child})")
    elif child <= parent:
        errors.append(f"slide 2 level-2 bullet not nested (indent {child} vs {parent})")
    for m in ("mk-left-3", "mk-right-3"):
        if item(sections[2], m) is None:
            errors.append(f"slide 3 bullet {m} not a list item")
    tab = [x for x in sections[3] if x.strip().startswith("|")]
    if not tab or not all(c in tab[0] for c in ("Region", "Revenue", "Plan", "Delta")):
        errors.append("slide 4 table header row missing")
    elif len(tab) < 2 or not re.match(r"^\|(\s*:?-{3,}:?\s*\|)+$", tab[1].strip()):
        errors.append("slide 4 table has no delimiter row")
    elif not any("mk-cell-4" in x for x in tab[2:]):
        errors.append("slide 4 table row with mk-cell-4 missing")
    for n, note in truth["notes"].items():
        if note not in " ".join(sections[int(n) - 1]).replace("> ", ""):
            errors.append(f"slide {n} notes missing from its section")
    return errors


oracle.main(check, pptx_output=False)
