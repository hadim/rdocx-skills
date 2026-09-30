# An edit re-serializes the whole part it touches, indented and with repeated namespace declarations

**What fails.** Any edit of the body (a single `try_replace_text`, a run's text, `rdocx replace`) rewrites all of
`word/document.xml`, not only the edited paragraph: every element is indented on its own line, and every
element that carries a `w:rsid*` attribute (Word writes them on nearly every paragraph and run) declares
`xmlns:w` again. Every other part keeps its bytes and the content of the rewritten one is kept (text,
attributes, `w14:paraId`, `mc:Ignorable`), so nothing is lost; but the part grows, and an XML diff of an
rdocx-edited file shows the whole part instead of the edit. On a synthetic report with Word's traits (504
paragraphs), one word replaced 14 times took `word/document.xml` from 252,662 to 456,768 bytes, with 1,092
declarations of `xmlns:w` instead of 1. rpptx, on the same kind of edit, rewrites the slide compactly (the
XML declaration dropped, attributes reordered, one namespace declared again on `p:spTree`).

**Reproduction** (builds its own input with python-docx; run with rdocx and python-docx importable):

```python
"""rdocx re-serializes the whole part an edit touches: indented, and with xmlns:w declared again on every element
that carries a w:rsid* attribute, as Word writes on nearly every paragraph and run. Builds its own input."""
import io, zipfile
from docx import Document
import rdocx
d = Document()
for i in range(50):
    d.add_paragraph(f"Paragraph {i} alpha beta.")
buf = io.BytesIO(); d.save(buf)
zin = zipfile.ZipFile(io.BytesIO(buf.getvalue())); src = io.BytesIO()
with zipfile.ZipFile(src, "w", zipfile.ZIP_DEFLATED) as zo:
    for i in zin.infolist():
        data = zin.read(i.filename)
        if i.filename == "word/document.xml":
            data = data.replace(b"<w:p>", b'<w:p w:rsidR="00AB12CD">')
        zo.writestr(i, data)
doc = rdocx.Document.from_bytes(src.getvalue())
assert doc.try_replace_text("Paragraph 7 ", "Paragraph seven ") == 1
a = zipfile.ZipFile(src).read("word/document.xml")
b = zipfile.ZipFile(io.BytesIO(doc.to_bytes())).read("word/document.xml")
nl = b"\n"
print("expected: document.xml about", len(a), "bytes, 1 xmlns:w, only paragraph 7 differs")
print("observed:", len(a), "->", len(b), "bytes;", "xmlns:w", a.count(b"xmlns:w="), "->", b.count(b"xmlns:w="),
      "; line breaks", a.count(nl), "->", b.count(nl))
i = b.find(b"<w:body")
print(b[i:i + 190].decode())
```

**Real output** (pinned commit d4d7c8af3fb1):

```
expected: document.xml about 5338 bytes, 1 xmlns:w, only paragraph 7 differs
observed: 5338 -> 10668 bytes; xmlns:w 1 -> 52 ; line breaks 1 -> 260
<w:body>
    <w:p w:rsidR="00AB12CD" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:r>
        <w:t>Paragraph 0 alpha beta.</w:t>
      </w:r>
    </w:p>
```

**Acceptance criterion.** After an edit, the rewritten part declares each namespace once (on the root, as read),
adds no indentation, and differs from the original only around the edited elements: on the reproduction,
`xmlns:w` is declared once and `document.xml` stays within a few bytes of its original size. Until then,
compare rdocx-edited files by text (`rdocx text --json`, `rdocx diff`), not by the XML of the parts.
