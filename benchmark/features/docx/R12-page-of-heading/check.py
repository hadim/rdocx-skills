"""R12 oracle: one record per heading, in order, with its level, text and page equal to the expected ones."""
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
sys.path.insert(0, str(Path(__file__).resolve().parent))
import review_oracle as o  # noqa: E402
from make_input import EXPECTED  # noqa: E402


def main(out):
    rows = json.loads(Path(out).read_text())["headings"]
    got = [(int(r["level"]), r["text"], int(r["page"])) for r in rows]
    checks = {"headings_listed": [(lv, t) for lv, t, _ in got] == [(lv, t) for lv, t, _ in EXPECTED]}
    for lv, t, page in EXPECTED:
        checks[f"page_of_{t}"] = any(g[1] == t and g[2] == page for g in got)
    o.report(checks, {"got": got})


o.guard(lambda: main(sys.argv[1]))
