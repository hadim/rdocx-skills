# compare fails its acceptance check when an empty paragraph gains paragraph properties

Pinned build: rdocx 0.16.0, commit `52e7f0b8e400239e96e8ac530d990fc038c54624` (ref `integration/open-prs-2026-10-12-4`), installed by rdocx-skills `a41d8553b2db`; reproduced on linux-x86_64 and linux-aarch64.

## What fails

When an empty paragraph gains paragraph properties (spacing after, alignment), `Document.compare` with default options raises `RdocxError: comparison acceptance does not reproduce the edited stories at body story item[0]`. In a header it raises the same error at `header story /word/header1.xml item[0]`. A non-empty paragraph with the same change compares fine, and so does the same change with `ignore_formatting=True`. Seen in a real header, where an empty paragraph gained `w:spacing w:after="0"`.

Met while redlining two versions of a set of ten Word documents with `rdocx compare --granularity word --ignore-fields --ignore-formatting`: eight compared and round-trip, two were refused. The cases below were reduced from the real pairs by greedy deletion of body blocks keeping the same error, then rebuilt from scratch.

## Expected

The comparison succeeds, with a tracked paragraph-property change (`w:pPrChange`) on the empty paragraph.

## Acceptance criterion

3a to 3c print OK.

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


def empty_then_text(space_after=None, centre=False):
    def f(doc):
        p = doc.add_paragraph("")
        doc.add_paragraph("Body.")
        if space_after is not None:
            p.paragraph_format.space_after = Pt(space_after)
        if centre:
            p.paragraph_format.alignment = 1
    return f


def header_paragraph(space_after=None):
    def f(doc):
        doc.add_paragraph("Body.")
        p = doc.sections[0].header.paragraphs[0]  # the header story exists on both sides
        if space_after is not None:
            p.paragraph_format.space_after = Pt(space_after)
    return f


def text_then_text(space_after=None):
    def f(doc):
        p = doc.add_paragraph("x")
        doc.add_paragraph("Body.")
        if space_after is not None:
            p.paragraph_format.space_after = Pt(space_after)
    return f


check("3a empty body paragraph gains space_after=0", empty_then_text(), empty_then_text(space_after=0))
check("3b empty body paragraph gains centring", empty_then_text(), empty_then_text(centre=True))
check("3c empty header paragraph gains space_after=0", header_paragraph(), header_paragraph(space_after=0))
check("3d control: one-character paragraph gains space_after=0", text_then_text(), text_then_text(space_after=0))
check("3e control: 3a with ignore_formatting", empty_then_text(), empty_then_text(space_after=0), ignore_formatting=True)
```

## Output of a real run

```
[FAIL] 3a empty body paragraph gains space_after=0 
       RdocxError: comparison acceptance does not reproduce the edited stories at body story item[0]
[FAIL] 3b empty body paragraph gains centring 
       RdocxError: comparison acceptance does not reproduce the edited stories at body story item[0]
[FAIL] 3c empty header paragraph gains space_after=0 
       RdocxError: comparison acceptance does not reproduce the edited stories at header story /word/header1.xml item[0]
[ OK ] 3d control: one-character paragraph gains space_after=0 
[ OK ] 3e control: 3a with ignore_formatting {'ignore_formatting': True}
```
