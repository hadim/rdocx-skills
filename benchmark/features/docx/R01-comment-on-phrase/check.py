"""R01 oracle: one new comment by Dana Reviewer with the text, initials and date asked, anchored on exactly the
phrase, with its reference mark; the earlier comment and the body text unchanged."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
sys.path.insert(0, str(Path(__file__).resolve().parent))
import review_oracle as o  # noqa: E402
from make_input import BODY, OLD_ANCHOR, OLD_COMMENT, PHRASE  # noqa: E402


def main(out):
    d = o.Docx(out)
    cs = o.comments(d)
    new = [cid for cid, c in cs.items() if c["author"] == "Dana Reviewer"]
    c = cs[new[0]] if len(new) == 1 else {}
    cid = new[0] if new else None
    old = cs.get(OLD_COMMENT["id"], {})
    o.report({
        "one_new_comment": len(new) == 1 and len(cs) == 2,
        "text": c.get("text") == "Should this be 45 days?",
        "initials": c.get("initials") == "DR",
        "date": (c.get("date") or "").startswith("2026-10-30T09:00:00"),
        "anchor_exact": cid is not None and o.anchored_text(d, cid) == PHRASE,
        "reference_mark": cid is not None and o.has_reference(d, cid),
        "old_comment_kept": old.get("text") == OLD_COMMENT["paragraphs"][0]
        and o.anchored_text(d, OLD_COMMENT["id"]) == OLD_ANCHOR,
        "body_unchanged": o.body(d) == o.norm(BODY),
    }, {"new": c, "anchor": o.anchored_text(d, cid) if cid is not None else None})


o.guard(lambda: main(sys.argv[1]))
