import sys

import docx_ops
import rdocx

src, out = sys.argv[1:3]
doc, diagnostics = rdocx.Document.open_markdown(src)
for d in diagnostics:
    print(f"{d.location}: {d.message}", file=sys.stderr)
# the import writes links without the Hyperlink style: they read as plain text unless coloured and underlined
links = {h.text for h in doc.hyperlinks}
for p in doc.paragraphs:
    for run in p.runs:
        if run.text in links:
            run.font.underline = True
            run.font.color.rgb = rdocx.RGBColor(0x05, 0x63, 0xC1)
docx_ops.save_atomic(doc, out)
