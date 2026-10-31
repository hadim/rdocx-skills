"""R08 oracle: accepted view == v2, rejected view == v1 (styles too); every revision by Dana Reviewer, dated;
word level: a word that a word diff of v1 and v2 leaves unchanged, and that is not next to a changed word, is in
no revision (a whole-run or whole-paragraph redline fails, and so does an unchanged word deleted and inserted
again); the removed paragraph is deleted as a whole."""
import difflib
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
sys.path.insert(0, str(Path(__file__).resolve().parent))
import review_oracle as o  # noqa: E402
from make_input import AUTHOR, V1, V2, flat  # noqa: E402


def deleted_mask(p):
    """(text, mask) of paragraph p in the rejected view; mask[i] is True when char i is in a w:del."""
    text, mask = [], []
    for el in p.iter(o.q("w:t"), o.q("w:delText"), o.q("w:tab")):
        if o._in(el, "w:ins", p):
            continue
        s = el.text or "" if el.tag != o.q("w:tab") else "\t"
        text.append(s)
        mask += [o._in(el, "w:del", p)] * len(s)
    return "".join(text), mask


def stray_words(d):
    """Unchanged words, not next to a change, found inside a deletion."""
    v2 = [t for _, t in flat(V2)]
    stray = []
    for p in d.document.iter(o.q("w:p")):
        text, mask = deleted_mask(p)
        if not text or all(mask):
            continue
        accepted = o.run_text(p, "accepted")
        target = difflib.get_close_matches(accepted, v2, n=1, cutoff=0.0)
        a, b = text.split(" "), (target[0] if target else "").split(" ")
        keep = set()
        ops = difflib.SequenceMatcher(None, a, b, autojunk=False).get_opcodes()
        for k, (op, i1, i2, _, _) in enumerate(ops):
            if op == "equal":
                lo = i1 + (1 if k > 0 else 0)
                hi = i2 - (1 if k < len(ops) - 1 else 0)
                keep.update(range(lo, hi))
        pos = 0
        for i, w in enumerate(a):
            if i in keep and any(mask[pos:pos + len(w)]):
                stray.append(w)
            pos += len(w) + 1
    return stray


def main(out):
    d = o.Docx(out)
    revs = [r for _, r in o.all_revisions(d)]
    stray = stray_words(d)
    o.report({
        "accepted_is_v2": o.body(d, "accepted") == o.norm(flat(V2)),
        "rejected_is_v1": o.body(d, "rejected") == o.norm(flat(V1)),
        "has_revisions": bool(revs),
        "author_and_date": all(r.get(o.q("w:author")) == AUTHOR and r.get(o.q("w:date")) for r in revs),
        "word_level": not stray,
    }, {"unchanged_words_in_revisions": stray, "accepted": o.body(d, "accepted"), "rejected": o.body(d, "rejected")})


o.guard(lambda: main(sys.argv[1]))
