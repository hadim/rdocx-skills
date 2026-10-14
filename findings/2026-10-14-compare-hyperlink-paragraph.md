# compare refuses a whole paragraph holding a hyperlink that is inserted or deleted

Pinned build: rdocx 0.16.0, commit `52e7f0b8e400239e96e8ac530d990fc038c54624` (ref `integration/open-prs-2026-10-12-4`), installed by rdocx-skills `a41d8553b2db`; reproduced on linux-x86_64 and linux-aarch64.

## What fails

When the edited document inserts or deletes a whole paragraph that contains a hyperlink, `Document.compare` (and `rdocx compare`) raises `RdocxError: comparison cannot track a whole paragraph's hyperlink or simple field outside a field it replaces at body`. Every granularity fails, and so do `ignore_fields` and `ignore_formatting`. The same paragraph edited in place compares fine. This blocked a real form whose how-to-answer paragraph, holding a mail link, was rewritten.

Met while redlining two versions of a set of ten Word documents with `rdocx compare --granularity word --ignore-fields --ignore-formatting`: eight compared and round-trip, two were refused. The cases below were reduced from the real pairs by greedy deletion of body blocks keeping the same error, then rebuilt from scratch.

## Expected

The comparison succeeds. The hyperlink's runs are tracked as inserted or deleted inside the `w:hyperlink`, as Word does.

## Acceptance criterion

1a and 1b print OK: `accept_all()` gives the edited text and `reject_all()` the original.

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


check("1a hyperlink paragraph inserted", build("One.", "Two."), build("One.", "LINK", "Two."))
check("1b hyperlink paragraph deleted", build("One.", "LINK", "Two."), build("One.", "Two."))
check("1c control: hyperlink paragraph edited in place", build("One.", "LINK", "Two."), build("One.", "LINK", "Two, changed."))
check("1d with ignore options", build("One.", "LINK", "Two."), build("One.", "Two."), ignore_fields=True, ignore_formatting=True)
```

## Output of a real run

```
[FAIL] 1a hyperlink paragraph inserted 
       RdocxError: comparison cannot track a whole paragraph's hyperlink or simple field outside a field it replaces at body
[FAIL] 1b hyperlink paragraph deleted 
       RdocxError: comparison cannot track a whole paragraph's hyperlink or simple field outside a field it replaces at body
[ OK ] 1c control: hyperlink paragraph edited in place 
[FAIL] 1d with ignore options {'ignore_fields': True, 'ignore_formatting': True}
       RdocxError: comparison cannot track a whole paragraph's hyperlink or simple field outside a field it replaces at body
```
