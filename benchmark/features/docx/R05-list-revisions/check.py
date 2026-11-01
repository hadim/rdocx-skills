"""R05 oracle: the listed revisions, as a multiset of (type, author, date, text), equal the input's five."""
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
sys.path.insert(0, str(Path(__file__).resolve().parent))
import review_oracle as o  # noqa: E402
from make_input import EXPECTED  # noqa: E402


def key(row):
    return (row["type"], row["author"], str(row["date"] or "").rstrip("Z")[:19], (row["text"] or "").strip())


def main(out):
    rows = json.loads(Path(out).read_text())["revisions"]
    got = sorted(key(r) for r in rows)
    want = sorted(key(dict(zip(("type", "author", "date", "text"), e))) for e in EXPECTED)
    checks = {f"{e[0]}:{e[3].strip()}": key(dict(zip(("type", "author", "date", "text"), e))) in got
              for e in EXPECTED}
    checks["nothing_else"] = len(got) == len(want)
    o.report(checks, {"got": got})


o.guard(lambda: main(sys.argv[1]))
