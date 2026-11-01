"""R04 oracle: the accepted view is the input with the word replaced, the rejected view is the input; the
deleted text is the old word and the inserted text the new one; every revision is Dana Reviewer's, dated."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
sys.path.insert(0, str(Path(__file__).resolve().parent))
import review_oracle as o  # noqa: E402
from make_input import ACCEPTED, BODY, NEW, OLD  # noqa: E402


def main(out):
    d = o.Docx(out)
    revs = o.all_revisions(d)
    dels = "".join(t.text or "" for _, r in revs if r.tag == o.q("w:del") for t in r.iter(o.q("w:delText")))
    inss = "".join(t.text or "" for _, r in revs if r.tag == o.q("w:ins") for t in r.iter(o.q("w:t")))
    o.report({
        "accepted_view": o.body(d, "accepted") == o.norm(ACCEPTED),
        "rejected_view": o.body(d, "rejected") == o.norm(BODY),
        "deleted_is_old_word": dels.strip() == OLD,
        "inserted_is_new_word": inss.strip() == NEW,
        "only_ins_del": {r.tag.split("}")[1] for _, r in revs} == {"ins", "del"},
        "author_and_date": all(r.get(o.q("w:author")) == "Dana Reviewer"
                               and (r.get(o.q("w:date")) or "").startswith("2026-10-30T11:00:00") for _, r in revs),
    }, {"deleted": dels, "inserted": inss, "revisions": [r.tag.split("}")[1] for _, r in revs]})


o.guard(lambda: main(sys.argv[1]))
