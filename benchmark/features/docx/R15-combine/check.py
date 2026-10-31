"""R15 oracle: A's paragraphs then B's, in order; each part keeps its own effective formatting (A's headings blue
Calibri Light and body Calibri, B's headings red Georgia, body Georgia, BNote green italic); A's list items stay
decimal "1.", B's lower roman "i)"; each picture shows its own image bytes."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
sys.path.insert(0, str(Path(__file__).resolve().parent))
import review_oracle as o  # noqa: E402
from make_input import A_BODY, A_COLOR, A_PNG, B_BODY, B_COLOR, B_PNG, NOTE_COLOR, NUMBERING  # noqa: E402


def main(out):
    d = o.Docx(out)
    paras = [p for p in d.document.find("w:body", o.NS).iter(o.q("w:p"))]
    texts = [o.run_text(p) for p in paras]
    want = [t for _, t in A_BODY + B_BODY if t != "PICTURE"]
    seq = [t for t in texts if t]

    def runs_props(p):
        return [o.effective_run(d, p, r) for r in p.iter(o.q("w:r")) if o.run_text(r)]

    fmt = {}
    for p, t in zip(paras, texts):
        if t:
            fmt[t] = (runs_props(p), o.numbering_format(d, p), o.style_of(p))
    side = lambda spec, style: [t for s, t in spec if s == style and t != "PICTURE"]  # noqa: E731
    all_runs = lambda ts, key, val: bool(ts) and all(  # noqa: E731
        t in fmt and fmt[t][0] and all(r.get(key) == val for r in fmt[t][0]) for t in ts)
    pics = []
    rels = d.rels(d.main_part())
    for p, t in zip(paras, texts):
        blip = next(p.iter(o.q("a:blip")), None)
        if blip is not None:
            target = rels.get(blip.get(o.q("r:embed")), (None, None))[1]
            pics.append(d.files.get(target))
    o.report({
        "a_then_b": seq == want,
        "a_headings_kept": all_runs(side(A_BODY, "Heading1"), "color", A_COLOR)
        and all_runs(side(A_BODY, "Heading1"), "font", "Calibri Light"),
        "a_body_font_kept": all_runs(side(A_BODY, None), "font", "Calibri"),
        "b_headings_kept": all_runs(side(B_BODY, "Heading1"), "color", B_COLOR)
        and all_runs(side(B_BODY, "Heading1"), "font", "Georgia"),
        "b_body_font_kept": all_runs(side(B_BODY, None), "font", "Georgia"),
        "b_own_style_kept": all_runs(side(B_BODY, "BNote"), "color", NOTE_COLOR)
        and all_runs(side(B_BODY, "BNote"), "i", True),
        "a_numbering_kept": all(fmt.get(t, (0, None))[1] == NUMBERING["A"] for t in side(A_BODY, "ListParagraph")),
        "b_numbering_kept": all(fmt.get(t, (0, None))[1] == NUMBERING["B"] for t in side(B_BODY, "ListParagraph")),
        "pictures_kept": pics == [A_PNG, B_PNG],
    }, {"sequence": seq, "formats": {t: [v[0][:1], v[1], v[2]] for t, v in fmt.items()},
        "pictures": [len(x) if x else None for x in pics]})


o.guard(lambda: main(sys.argv[1]))
