# compare refuses a paragraph inserted or deleted right before a table

Pinned build: rdocx 0.16.0, commit `52e7f0b8e400239e96e8ac530d990fc038c54624` (ref `integration/open-prs-2026-10-12-4`), installed by rdocx-skills `a41d8553b2db`; reproduced on linux-x86_64 and linux-aarch64.

## What fails

When the paragraph right before a table is deleted (or inserted), `Document.compare` raises `RdocxError: comparison needs an adjacent paragraph for a final paragraph change`, although the changed paragraph is not the final one. Every granularity fails, and so do the ignore options. The failure depends on content: the controls 2e and 2f, nearly identical, compare fine. This blocked a real two-page letter.

Met while redlining two versions of a set of ten Word documents with `rdocx compare --granularity word --ignore-fields --ignore-formatting`: eight compared and round-trip, two were refused. The cases below were reduced from the real pairs by greedy deletion of body blocks keeping the same error, then rebuilt from scratch.

## Expected

The comparison succeeds, with the paragraph tracked as deleted (or inserted).

## Acceptance criterion

2a to 2d print OK.

## Workaround

Workaround (2g): add an empty paragraph before every table in both inputs. It works on the real pair, at the cost of a blank line before each table in the redline.

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


check("2a first paragraph, right before a table, deleted", build("Old.", "TABLE:Same cell.", "End."), build("TABLE:Same cell.", "End."))
check("2b same, inserted", build("TABLE:Same cell.", "End."), build("New.", "TABLE:Same cell.", "End."))
check("2c paragraph before a table deleted, cell edited", build("Intro.", "Old.", "TABLE:Old cell.", "End."), build("Intro.", "TABLE:New cell.", "End."))
check("2d same with ignore options", build("Intro.", "Old.", "TABLE:Old cell.", "End."), build("Intro.", "TABLE:New cell.", "End."), ignore_fields=True, ignore_formatting=True)
check("2e control: as 2a with an intro paragraph", build("Intro.", "Old.", "TABLE:Same cell.", "End."), build("Intro.", "TABLE:Same cell.", "End."))
check("2f control: as 2c with the cell unchanged", build("Intro.", "Old.", "TABLE:Cell text.", "End."), build("Intro.", "TABLE:Cell text.", "End."))
check("2g workaround: an empty paragraph before the table in both", build("Intro.", "Old.", "", "TABLE:Old cell.", "End."), build("Intro.", "", "TABLE:New cell.", "End."))
```

## Output of a real run

```
[FAIL] 2a first paragraph, right before a table, deleted 
       RdocxError: comparison needs an adjacent paragraph for a final paragraph change
[FAIL] 2b same, inserted 
       RdocxError: comparison needs an adjacent paragraph for a final paragraph change
[FAIL] 2c paragraph before a table deleted, cell edited 
       RdocxError: comparison needs an adjacent paragraph for a final paragraph change
[FAIL] 2d same with ignore options {'ignore_fields': True, 'ignore_formatting': True}
       RdocxError: comparison needs an adjacent paragraph for a final paragraph change
[ OK ] 2e control: as 2a with an intro paragraph 
[ OK ] 2f control: as 2c with the cell unchanged 
[ OK ] 2g workaround: an empty paragraph before the table in both
```
