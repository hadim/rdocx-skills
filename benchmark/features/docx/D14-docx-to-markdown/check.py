"""D14 oracle: headings at the right level, every list item with the right marker, every table row on one line
with its cells in order, emphasis kept."""
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from authoring_oracle import Check  # noqa: E402

ck = Check()
md = Path(sys.argv[1]).read_text(encoding="utf-8")
lines = md.splitlines()


def heading(text, level):
    if any(re.fullmatch(rf"#{{{level}}}\s+{re.escape(text)}(\s+#*)?\s*(\{{[^}}]*\}})?", l) for l in lines):
        return True
    if level <= 2:
        under = "=" if level == 1 else "-"
        return any(l.strip() == text and k + 1 < len(lines) and re.fullmatch(rf"{under}{{3,}}\s*", lines[k + 1])
                   for k, l in enumerate(lines))
    return False


for text, level in (("Field report", 1), ("Findings", 2), ("Actions", 2), ("Costs", 3)):
    ck.require(heading(text, level), f"no level-{level} heading '{text}'")
for text in ("Cracked kerb", "Loose railing", "Blocked drain"):
    ck.require(any(re.fullmatch(rf"\s*[-*+]\s+{text}\s*", l) for l in lines), f"no bullet item '{text}'")
for text in ("Repair the kerb", "Fix the railing", "Clear the drain"):
    ck.require(any(re.fullmatch(rf"\s*\d+[.)]\s+{text}\s*", l) for l in lines), f"no numbered item '{text}'")
for row in (("Item", "Hours", "Cost"), ("Kerb", "6", "480"), ("Railing", "3", "210")):
    pat = r"[\s|*_]+".join(re.escape(c) for c in row)
    ck.require(any(re.search(rf"(^|[\s|*_]){pat}([\s|*_]|$)", l) for l in lines), f"no table row {row}")
flat = re.sub(r"\s+", " ", md)
ck.require(re.search(r"(\*\*|__)three defects(\*\*|__)", flat), "'three defects' is not bold")
ck.require(re.search(r"(?<![*_])[*_]open question[*_](?![*_])", flat), "'open question' is not italic")
ck.require("Report closed." in flat, "last paragraph missing")
ck.require("<w:" not in md, "raw XML in the output")
ck.done()
