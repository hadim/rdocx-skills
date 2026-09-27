`Document.add_picture` fails on a file that has a content control and a default namespace on its root

Google Docs exports declare a default namespace on the root of `document.xml`
(`xmlns="http://schemas.microsoft.com/office/tasks/2019/documenttasks"`, unused) and wrap text in
`goog_rdk_*` content controls. On such a file, `add_picture` raises
``cannot serialize a modified document with a shadowed `default` namespace`` (the fail-closed check of
`flush_document_to_package`, `rdocx/src/document.rs:12465`), while every other edit I tried on the same file
(`add_paragraph`, `insert_paragraph`, `add_table`, `try_replace_text`, `add_hyperlink`, `set_header`)
succeeds and saves. Without the content control, or without the default namespace, `add_picture` works.
The namespace is declared but never used, so there is nothing a canonical serialization could change.

```python
import io, zipfile
import rdocx
from docx import Document
from docx.oxml.ns import qn
from PIL import Image


def build(path, default_ns):
    d = Document()
    d.add_paragraph("Before the control.")
    body = d.element.body
    sdt = body.makeelement(qn("w:sdt"), {})
    pr = sdt.makeelement(qn("w:sdtPr"), {})
    tag = pr.makeelement(qn("w:tag"), {}); tag.set(qn("w:val"), "goog_rdk_0"); pr.append(tag); sdt.append(pr)
    content = sdt.makeelement(qn("w:sdtContent"), {})
    p = d.add_paragraph("Inside the control.")._p; body.remove(p); content.append(p); sdt.append(content)
    body.insert(1, sdt)
    d.save(path)
    if default_ns:  # as Google Docs exports write the root
        items = [(i, zipfile.ZipFile(path).read(i.filename)) for i in zipfile.ZipFile(path).infolist()]
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
            for info, data in items:
                if info.filename == "word/document.xml":
                    data = data.replace(b"<w:document ", b'<w:document xmlns="http://schemas.microsoft.com/office/tasks/2019/documenttasks" ', 1)
                z.writestr(info, data)
    return path


png = io.BytesIO(); Image.new("RGB", (60, 40)).save(png, "PNG")
for default_ns in (False, True):
    doc = rdocx.Document.open(build("sdt.docx", default_ns))
    try:
        doc.add_picture(png.getvalue(), "x.png", width=rdocx.Inches(1), height=rdocx.Inches(1))
        result = "ok"
    except rdocx.RdocxError as e:
        result = f"RdocxError: {e}"
    print(f"content control, default namespace on the root: {default_ns!s:<5} -> add_picture {result}")
```

```
content control, default namespace on the root: False -> add_picture ok
content control, default namespace on the root: True  -> add_picture RdocxError: cannot serialize a modified document with a shadowed `default` namespace
```

*Acceptance:* `add_picture` succeeds on this file and on a Google Docs export, and the producer-traits
matrix gains an `add_picture` column that passes on the "default namespace on the root" and "inline content
control" rows.

Environment: `main` at `9a7ed714` (S75), release build, linux x86_64, Python 3.11, python-docx 1.2.0.
