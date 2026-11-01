"""R03 oracle: one record per comment with its author, date, text (paragraphs joined by a newline) and anchored
text; a reply may give its parent's anchor or null."""
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
sys.path.insert(0, str(Path(__file__).resolve().parent))
import review_oracle as o  # noqa: E402
from make_input import COMMENTS, EXPECTED  # noqa: E402


def main(out):
    rows = {int(r["id"]): r for r in json.loads(Path(out).read_text())["comments"]}
    checks = {"all_listed": sorted(rows) == sorted(c["id"] for c in COMMENTS)}
    for c in COMMENTS:
        r = rows.get(c["id"], {})
        anchor, parent = EXPECTED[c["id"]]
        checks[f"{c['id']}_author"] = r.get("author") == c["author"]
        checks[f"{c['id']}_date"] = str(r.get("date") or "").rstrip("Z")[:19] == c["date"].rstrip("Z")
        checks[f"{c['id']}_text"] = (r.get("text") or "").strip() == "\n".join(c["paragraphs"])
        ok = [anchor, None] if parent is not None else [anchor]
        checks[f"{c['id']}_anchor"] = r.get("anchor") in ok
    o.report(checks, {"rows": rows})


o.guard(lambda: main(sys.argv[1]))
