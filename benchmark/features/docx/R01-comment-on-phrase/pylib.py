# bench: {"tools": ["python-docx", "lxml (run split, comment date)"], "processes": 1, "raw_xml": true, "notes": "python-docx anchors on whole runs only: the runs are split by copying their XML; the comment gets no w14:paraId and no commentsExtended entry; the date has no public setter"}
import copy
import sys
from datetime import datetime, timezone

import docx
from docx.text.run import Run

PHRASE = "thirty days after delivery"
doc = docx.Document(sys.argv[1])
p = next(p for p in doc.paragraphs if PHRASE in p.text)
start = p.text.index(PHRASE)
end = start + len(PHRASE)


def split(run, at):
    """Split a run at a character offset (python-docx has no API for it); return the right part."""
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
comment = doc.add_comment(covered, text="Should this be 45 days?", author="Dana Reviewer", initials="DR")
comment._comment_elm.date = datetime(2026, 10, 30, 9, 0, tzinfo=timezone.utc)  # no public setter
doc.save(sys.argv[2])
