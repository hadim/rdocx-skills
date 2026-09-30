# add_style refuses several default styles of one type

**What fails.** When `word/styles.xml` marks two styles of one type as default with different ids, as Google
Docs writes them (a `TableNormal` and a localized `TableauNormal`, both `w:default="1"`), `Document.add_style`
raises `RdocxError: invalid style graph: style type 'table' has more than one default`. It is the sibling of the
gap `styles-duplicate-ids`: a file repaired for duplicate ids only is still refused. Met on real files on
30/09/2026; the documented workaround (drop later duplicate `w:style` elements) is not enough on its own.

**Reproduction** (builds its own input with python-docx; run with rdocx and python-docx importable):

```python
"""rdocx add_style refuses a package whose styles.xml marks two table styles as default (different ids), as
Google Docs writes them (a TableNormal and a localized TableauNormal, both w:default="1"). Builds its own input."""
import io, sys, zipfile
from docx import Document
sys.path.insert(0, sys.argv[1]) if len(sys.argv) > 1 else None
import rdocx
buf = io.BytesIO(); Document().save(buf)
zin = zipfile.ZipFile(io.BytesIO(buf.getvalue()))
extra = b'<w:style w:type="table" w:default="1" w:styleId="TableauNormal"><w:name w:val="Tableau Normal"/></w:style></w:styles>'
out = io.BytesIO()
with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zo:
    for i in zin.infolist():
        d = zin.read(i.filename)
        if i.filename == "word/styles.xml":
            assert d.count(b'w:type="table" w:default="1"') == 1
            d = d.replace(b"</w:styles>", extra)
        zo.writestr(i, d)
doc = rdocx.Document.from_bytes(out.getvalue())
try:
    doc.add_style("Probe", "paragraph", based_on="Normal")
    print("expected: add_style succeeds; observed: succeeds")
except rdocx.RdocxError as e:
    print("expected: add_style succeeds (first default of a type authoritative); observed: RdocxError:", e)
```

**Real output** (pinned commit d4d7c8af3fb1):

```
expected: add_style succeeds (first default of a type authoritative); observed: RdocxError: invalid style graph: style type 'table' has more than one default
```

**Acceptance criterion.** `add_style` succeeds on such a package, the first default style of each type being
authoritative (as `rebuild_toc` treats duplicate ids). Until then, the `styles-duplicate-ids` workaround in
`skills/docx/references/gaps.md` should also remove `w:default` from every later default style of a type
(done that way on real files: renders identical before and after, rdocx and LibreOffice).
