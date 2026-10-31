"""A comment anchored on a phrase with python-docx 1.2 (doc.add_comment anchors on whole runs: the runs at the
phrase's edges are split by copying their XML, python-docx has no API for it).
    python docx_pylib_comment.py IN OUT PHRASE"""
import copy
import sys

import docx
from docx.text.run import Run

src, out, phrase = sys.argv[1:4]
doc = docx.Document(src)
p = next(p for p in doc.paragraphs if phrase in "".join(r.text for r in p.runs))
start = "".join(r.text for r in p.runs).index(phrase)
end = start + len(phrase)


def split(run, at):
    """Split a run at a character offset; return the right part."""
    right = copy.deepcopy(run._r)
    run._r.addnext(right)
    tail = Run(right, run._parent)
    tail.text, run.text = run.text[at:], run.text[:at]
    return tail


pos, covered = 0, []
for run in list(p.runs):
    a, b = pos, pos + len(run.text)
    pos = b
    if b <= start or a >= end:
        continue
    if a < start:
        run, a = split(run, start - a), start
    if b > end:
        split(run, end - a)
    covered.append(run)
doc.add_comment(covered, text="Speed benchmark comment", author="Speed Bench", initials="SB")
doc.save(out)
