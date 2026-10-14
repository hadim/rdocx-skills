# compare fails its acceptance check when paragraphs holding only a line break are removed

Pinned build: rdocx 0.16.0, commit `52e7f0b8e400239e96e8ac530d990fc038c54624` (ref `integration/open-prs-2026-10-12-4`), installed by rdocx-skills `a41d8553b2db`; reproduced on linux-x86_64 and linux-aarch64.

## What fails

Consider two paragraphs holding only a line break (`<w:r><w:br/></w:r>`, the answer boxes of a form), followed by a text paragraph. When they become one break-only paragraph, `Document.compare` raises `RdocxError: comparison acceptance does not reproduce the edited stories at body story item[1]`. The ignore options fail the same way. Smaller variants compare fine (4c, 4d). In a real form this was the second failure, hidden behind compare-hyperlink-paragraph until the hyperlinks were removed from the inputs.

Met while redlining two versions of a set of ten Word documents with `rdocx compare --granularity word --ignore-fields --ignore-formatting`: eight compared and round-trip, two were refused. The cases below were reduced from the real pairs by greedy deletion of body blocks keeping the same error, then rebuilt from scratch.

## Expected

The comparison succeeds.

## Acceptance criterion

4a and 4b print OK.

## Reproduction

Builds its own input; run with the pinned rdocx Python (`$RDOCX_HOME/current/bin/python repro.py`).

```python
import os
import re
import tempfile

import rdocx
from rdocx import Pt

TMP = tempfile.mkdtemp()


def build(*blocks):
    """A string is a paragraph; TABLE:<text> a 1x1 table; LINK a paragraph holding a hyperlink; BREAK a paragraph
    holding only a line break."""
    def f(doc):
        for b in blocks:
            if b.startswith("TABLE:"):
                doc.add_table(1, 1).cell(0, 0).text = b[len("TABLE:"):]
            elif b == "LINK":
                p = doc.add_paragraph("Write to ")
                p.add_hyperlink("the team", "https://example.com")
                p.add_run(" for help.")
            elif b == "BREAK":
                doc.add_paragraph().add_run("").add_break()
            else:
                doc.add_paragraph(b)
    return f


def save(builder, name):
    doc = rdocx.Document()
    builder(doc)
    path = os.path.join(TMP, name)
    doc.save(path)
    return path


def check(case, build_a, build_b, **options):
    """Expected: compare succeeds, accept_all() gives B's text, reject_all() gives A's text."""
    a, b = save(build_a, case + "-a.docx"), save(build_b, case + "-b.docx")
    doc = rdocx.Document(a)
    try:
        doc.compare(rdocx.Document(b), "Reviewer", "2026-01-01T00:00:00Z", granularity="word", **options)
    except Exception as e:
        print("[FAIL] %s %s\n       %s: %s" % (case, options or "", type(e).__name__, e))
        return False
    out = os.path.join(TMP, case + "-redline.docx")
    doc.save(out)
    acc, rej = rdocx.Document(out), rdocx.Document(out)
    acc.accept_all()
    rej.reject_all()
    ok = acc.text() == rdocx.Document(b).text() and rej.text() == rdocx.Document(a).text()
    print("[%s] %s %s" % (" OK " if ok else "FAIL", case, options or ""))
    return ok


check("4a [break, break, text] -> [break]", build("BREAK", "BREAK", "Heading."), build("BREAK"))
check("4b same with ignore options", build("BREAK", "BREAK", "Heading."), build("BREAK"), ignore_fields=True, ignore_formatting=True)
check("4c control: [break, break] -> [break]", build("BREAK", "BREAK"), build("BREAK"))
check("4d control: [break, text] -> [break]", build("BREAK", "Heading."), build("BREAK"))
```

## Output of a real run

```
[FAIL] 4a [break, break, text] -> [break] 
       RdocxError: comparison acceptance does not reproduce the edited stories at body story item[1]
[FAIL] 4b same with ignore options {'ignore_fields': True, 'ignore_formatting': True}
       RdocxError: comparison acceptance does not reproduce the edited stories at body story item[1]
[ OK ] 4c control: [break, break] -> [break] 
[ OK ] 4d control: [break, text] -> [break]
```
