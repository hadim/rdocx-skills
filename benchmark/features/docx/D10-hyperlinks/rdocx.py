import sys

import docx_ops
import rdocx


def look_like_a_link(run):
    run.font.underline = True
    run.font.color.rgb = rdocx.RGBColor(0x05, 0x63, 0xC1)


src, out = sys.argv[1:3]
doc = rdocx.Document.open(src)
i = next(i for i, p in enumerate(doc.paragraphs) if p.text.startswith("Further reading:"))
look_like_a_link(doc.paragraphs[i].add_hyperlink("Project site", "https://example.com/project"))
doc.paragraphs[i].add_run(" and ")
heading = next(h for h in doc.paragraphs if h.text == "Appendix A")
look_like_a_link(doc.add_heading_hyperlink(doc.paragraphs[i], heading, "Appendix A"))
doc.paragraphs[i].add_run(".")
docx_ops.save_atomic(doc, out, src)
