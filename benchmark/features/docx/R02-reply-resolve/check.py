"""R02 oracle: a reply by Dana Reviewer with the text and date asked, threaded under Alex Author's comment in
commentsExtended (paraIdParent), that thread resolved (done="1"), the other thread still open, the body text
unchanged. The reply must also be referenced from the document, next to its parent: Word writes a range and a
reference mark for every reply, and LibreOffice drops a comment that the document never references."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
sys.path.insert(0, str(Path(__file__).resolve().parent))
import review_oracle as o  # noqa: E402
from make_input import BODY, THREADS  # noqa: E402


def main(out):
    d = o.Docx(out)
    cs, ex = o.comments(d), o.comments_ex(d)
    new = [cid for cid, c in cs.items() if c["author"] == "Dana Reviewer"]
    reply = cs[new[0]] if len(new) == 1 else {"para_ids": [None]}
    parent, other = cs.get(0, {"para_ids": [None]}), cs.get(1, {"para_ids": [None]})
    rex = ex.get(reply["para_ids"][-1], (None, None))
    o.report({
        "one_reply": len(new) == 1 and len(cs) == 3,
        "text": reply.get("text") == "Agreed, changed to 45 days.",
        "date": (reply.get("date") or "").startswith("2026-10-30T10:00:00"),
        "threaded_under_parent": rex[0] is not None and rex[0] == parent["para_ids"][-1],
        "thread_resolved": ex.get(parent["para_ids"][-1], (None, None))[1] in ("1", "true"),
        "other_thread_open": ex.get(other["para_ids"][-1], (None, "0"))[1] in ("0", "false", None),
        "reply_referenced": bool(new) and o.has_reference(d, new[0]),
        "parents_unchanged": parent.get("text") == THREADS[0][0]["paragraphs"][0]
        and o.anchored_text(d, 0) == THREADS[0][2] and o.anchored_text(d, 1) == THREADS[1][2],
        "body_unchanged": o.body(d) == o.norm(BODY),
    }, {"reply": reply, "commentsExtended": ex})


o.guard(lambda: main(sys.argv[1]))
