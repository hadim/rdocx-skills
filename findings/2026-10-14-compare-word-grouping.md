# compare at word granularity interleaves one-word changes, which makes a rewritten paragraph unreadable

Pinned build: rdocx 0.16.0, commit `52e7f0b8e400239e96e8ac530d990fc038c54624` (ref `integration/open-prs-2026-10-12-4`), installed by rdocx-skills `a41d8553b2db`; reproduced on linux-x86_64 and linux-aarch64.

## What fails

Feature request, not a failure. At `granularity="word"`, a stretch of consecutive changed words comes out as alternating one-word pairs, `[-the-]{+to+} [-site-]{+the+} [-of-]{+laboratory+}`, instead of one deletion followed by one insertion. Short common words (`the`, `and`, `,`) kept as anchors inside a rewritten sentence split the changes further. On a paragraph that was largely rewritten the redline becomes word salad: one real five-page annex gives 1251 revisions at word granularity, against 76 at run granularity. Run granularity avoids the salad only by replacing the whole run, which hides light edits.

## Expected

(a) Consecutive changed words are grouped into one deletion then one insertion, as Word and legal comparison tools display them. (b) Optionally, a semantic clean-up like diff-match-patch's `diff_cleanupSemantic`: an unchanged stretch shorter than the changes around it is folded into them. A similarity threshold under which a sentence or paragraph is replaced whole would serve too.

## Acceptance criterion

6 word prints `... and [-notifies the site of-]{+reports to the laboratory+} any specimen that [-does not pass-]{+fails+} inspection[-, with the reason-].`

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


def markup(xml):
    """The tracked paragraph as text: [-deleted-] {+inserted+}."""
    out = []
    for m in re.finditer(r"<w:(ins|del)\b[^>]*>(.*?)</w:\1>|<w:r\b[^>]*>(.*?)</w:r>", xml, re.S):
        kind, inner = (m.group(1), m.group(2)) if m.group(1) else (None, m.group(3))
        text = "".join(re.findall(r"<w:(?:t|delText)\b[^>]*>([^<]*)</w:(?:t|delText)>", inner))
        if text:
            out.append({"ins": "{+%s+}", "del": "[-%s-]", None: "%s"}[kind] % text)
    return "".join(out)


OLD = "On receipt, the Sponsor inspects the material and notifies the site of any specimen that does not pass inspection, with the reason."
NEW = "After receipt, the Sponsor confirms the delivery and reports to the laboratory any specimen that fails inspection."
for granularity in ("word", "run"):
    a, b = save(build(OLD), "6-a.docx"), save(build(NEW), "6-b.docx")
    doc = rdocx.Document(a)
    doc.compare(rdocx.Document(b), "Reviewer", "2026-01-01T00:00:00Z", granularity=granularity)
    xml = doc.paragraphs[0].xml
    xml = xml.decode() if isinstance(xml, bytes) else xml
    print("6 %s: %d revisions\n  %s" % (granularity, len(doc.revisions), markup(xml)))
```

## Output of a real run

```
6 word: 18 revisions
  [-On-]{+After+} receipt, the Sponsor [-inspects-]{+confirms+} the [-material-]{+delivery+} and [-notifies-]{+reports+} [-the-]{+to+} [-site-]{+the+} [-of-]{+laboratory+} any specimen that [-does-]{+fails+}[- not pass-] inspection[-, with the reason-].
6 run: 2 revisions
  [-On receipt, the Sponsor inspects the material and notifies the site of any specimen that does not pass inspection, with the reason.-]{+After receipt, the Sponsor confirms the delivery and reports to the laboratory any specimen that fails inspection.+}
```
