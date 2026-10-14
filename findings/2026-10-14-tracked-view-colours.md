# the tracked revision view has no colours, and raw XML stops short of headers, footers and some pictures

Pinned build: rdocx 0.16.0, commit `52e7f0b8e400239e96e8ac530d990fc038c54624` (ref `integration/open-prs-2026-10-12-4`), installed by rdocx-skills `a41d8553b2db`; reproduced on linux-x86_64 and linux-aarch64.

## What fails

Feature request with two gaps. `rdocx convert/render --revision-view tracked`, and `to_pdf`/`render_pages` with `revision_view="tracked"`, draw insertions as a black underline, deletions as a black strikethrough and a black change bar, with no option. Reviewers expect colour: the legal comparison tools (DeltaView, Workshare, now Litera) show insertions in blue underlined, deletions in red struck through and moved text in green. In Word these colours are application settings (Track Changes Options), not document content, so the renderer is where they belong. The workaround in use colours the runs inside `w:ins`/`w:del` of a throwaway copy through `Paragraph.xml` and `replace_xml`, then renders it. Two gaps cut it short: (5a) a header or footer paragraph has no raw XML (`NotImplementedError`), so revisions there stay black; (5b) `replace_xml` refuses a picture paragraph with an SVG twin (`asvg:svgBlip`) when the `asvg` namespace is declared on the `w:p` rather than on `svgBlip`. Both texts are the same XML infoset, and `xml.etree` hoists declarations to the root. Also seen on screen: `w:u w:val="double"` renders as a single underline.

## Expected

(5) A tracked-view option for the colour and underline style of insertions, the colour of deletions, and moves, for example a `legal` preset. (5a) Raw XML on header and footer paragraphs. (5b) `replace_xml` accepts the namespace of `svgBlip` declared on an ancestor.

## Acceptance criterion

5a and 5b-on-w:p print OK. A tracked render can be asked in colour.

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


import struct
import zlib


def tiny_png():
    def chunk(tag, data):
        return struct.pack(">I", len(data)) + tag + data + struct.pack(">I", zlib.crc32(tag + data) & 0xFFFFFFFF)
    ihdr = struct.pack(">IIBBBBB", 1, 1, 8, 2, 0, 0, 0)
    return b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", ihdr) + chunk(b"IDAT", zlib.compress(b"\x00\xff\x00\x00")) + chunk(b"IEND", b"")


doc = rdocx.Document()
doc.add_paragraph("Body.")
doc.sections[0].header.paragraphs[0].text = "Header text."
try:
    doc.sections[0].header.paragraphs[0].xml
    print("[ OK ] 5a header paragraph .xml")
except Exception as e:
    print("[FAIL] 5a header paragraph .xml\n       %s: %s" % (type(e).__name__, e))

png = os.path.join(TMP, "dot.png")
open(png, "wb").write(tiny_png())
NS = 'xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main"'
for where in ("svgBlip", "w:p"):
    doc = rdocx.Document()
    doc.add_picture(png, width=rdocx.Inches(0.5))
    doc.add_picture(png, width=rdocx.Inches(0.5))  # a second image part: its id stands in for the SVG part
    x0, x1 = (p.xml.decode() if isinstance(p.xml, bytes) else p.xml for p in doc.paragraphs[:2])
    rid = re.search(r'r:embed="([^"]+)"', x1).group(1)
    ext = ('<a:extLst><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}"><asvg:svgBlip %s r:embed="%s"/>'
           '</a:ext></a:extLst>' % (NS if where == "svgBlip" else "", rid))
    x2 = re.sub(r"(<a:blip\b[^>]*?)/>", lambda m: m.group(1) + ">" + ext + "</a:blip>", x0, count=1) \
        if "</a:blip>" not in x0 else x0.replace("</a:blip>", ext + "</a:blip>", 1)
    if where == "w:p":
        x2 = re.sub(r"^<w:p\b", "<w:p " + NS, x2, count=1)
    try:
        doc.paragraphs[0].replace_xml(x2)
        print("[ OK ] 5b svgBlip namespace declared on %s" % where)
    except Exception as e:
        print("[FAIL] 5b svgBlip namespace declared on %s\n       %s: %s" % (where, type(e).__name__, e))
```

## Output of a real run

```
[FAIL] 5a header paragraph .xml
       NotImplementedError: raw XML works on body and table-cell paragraphs, not yet in a header or footer
[ OK ] 5b svgBlip namespace declared on svgBlip
[FAIL] 5b svgBlip namespace declared on w:p
       ValueError: replace_xml expects one w:p element: rdocx would not keep svgBlip, so the replacement was refused
```
