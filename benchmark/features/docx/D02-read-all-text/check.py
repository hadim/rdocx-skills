"""D02 oracle: each marker is in its story's key, once, and in no other key; body order is kept."""
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from authoring_oracle import Check  # noqa: E402

ck = Check()
data = json.loads(Path(sys.argv[1]).read_text())
keys = ["body", "headers", "footers", "footnotes", "text_boxes", "comments"]
ck.require(isinstance(data, dict) and all(isinstance(data.get(k), list) for k in keys), f"keys {keys} expected")
data = {k: [str(x) for x in data.get(k) or []] for k in keys}
markers = {"MARK_BODY_1": "body", "MARK_TABLE_1": "body", "MARK_BODY_2": "body", "MARK_HEADER": "headers",
           "MARK_FOOTER": "footers", "MARK_FOOTNOTE": "footnotes", "MARK_TEXTBOX": "text_boxes",
           "MARK_COMMENT": "comments"}
for marker, key in markers.items():
    n = sum(t.count(marker) for t in data[key])
    ck.require(n == 1, f"{marker}: found {n} time(s) in {key!r}, expected 1")
    for other in keys:
        if other != key:
            ck.require(not any(marker in t for t in data[other]), f"{marker} leaks into {other!r}")


def first(text):
    return next((i for i, t in enumerate(data["body"]) if text in t), None)


order = [first(t) for t in ("Read test", "MARK_BODY_1", "Cell A", "MARK_TABLE_1", "Cell D", "MARK_BODY_2",
                            "Anchor of the box.", "Last body paragraph.")]
ck.require(None not in order and order == sorted(order), f"body order wrong: {order}")
ck.done()
