"""R10 oracle: the body's DATE, NUMPAGES and PAGE fields are still fields (complex or simple) and show today's date,
3 and 3; the rest of the body text unchanged."""
import datetime
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
sys.path.insert(0, str(Path(__file__).resolve().parent))
import review_oracle as o  # noqa: E402
from make_input import FIELDS  # noqa: E402


def fields(root):
    """[(field name, result text)] of the complex and simple fields under root, in order."""
    out, stack = [], []
    for el in root.iter(o.q("w:fldChar"), o.q("w:instrText"), o.q("w:t"), o.q("w:fldSimple")):
        if el.tag == o.q("w:fldSimple"):
            out.append(((el.get(o.q("w:instr")) or "").split()[0], o.run_text(el)))
        elif el.tag == o.q("w:fldChar"):
            kind = el.get(o.q("w:fldCharType"))
            if kind == "begin":
                stack.append(["", "", "instr"])
            elif kind == "separate" and stack:
                stack[-1][2] = "result"
            elif kind == "end" and stack:
                instr, result, _ = stack.pop()
                out.append((instr.split()[0] if instr.split() else "", result))
        elif el.tag == o.q("w:instrText") and stack:
            stack[-1][0] += el.text or ""
        elif el.tag == o.q("w:t") and stack and stack[-1][2] == "result" and not o._in(el, "w:fldSimple", root):
            stack[-1][1] += el.text or ""
    return out


def main(out):
    d = o.Docx(out)
    got = fields(d.document.find("w:body", o.NS))
    today = datetime.date.today().isoformat()
    want = [(instr.split()[0], expected or today) for instr, _, expected in FIELDS]
    text = [t for _, t in o.body(d)]
    o.report({
        "still_fields": [n for n, _ in got] == [n for n, _ in want],
        "date": ("DATE", today) in got,
        "numpages": ("NUMPAGES", "3") in got,
        "page": ("PAGE", "3") in got,
        "text_around_kept": text[1].startswith("Report date: ") and text[-1].startswith("This paragraph is on page "),
    }, {"fields": got})


o.guard(lambda: main(sys.argv[1]))
