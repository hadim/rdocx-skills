# bench: tools=python-pptx; processes=1; raw_xml=no
import sys

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE


def shape_md(sh):
    if sh.shape_type == MSO_SHAPE_TYPE.GROUP:
        return [line for child in sh.shapes for line in shape_md(child)]
    if getattr(sh, "has_table", False) and sh.has_table:
        rows = [[c.text.replace("|", "\\|").replace("\n", " ") for c in r.cells] for r in sh.table.rows]
        lines = ["| " + " | ".join(r) + " |" for r in rows]
        return ["", lines[0], "|" + " --- |" * len(rows[0]), *lines[1:], ""]
    if sh.shape_type == MSO_SHAPE_TYPE.PICTURE or (sh.is_placeholder and hasattr(sh, "image")):
        return [f"![{sh.name}]()"]
    if sh.has_text_frame:
        return ["  " * p.level + "- " + p.text for p in sh.text_frame.paragraphs if p.text.strip()]
    return []


prs = Presentation(sys.argv[1])
out = []
for k, slide in enumerate(prs.slides):
    title = slide.shapes.title
    out += [f"## {title.text if title is not None else f'Slide {k + 1}'}", ""]
    for sh in sorted(slide.shapes, key=lambda s: (s.top or 0, s.left or 0)):
        if title is None or sh.shape_id != title.shape_id:
            out += shape_md(sh)
    if slide.has_notes_slide and slide.notes_slide.notes_text_frame.text.strip():
        out += ["", "> **Notes:** " + slide.notes_slide.notes_text_frame.text.replace("\n", "\n> "), ""]
    out.append("")
with open(sys.argv[2], "w") as f:
    f.write("\n".join(out))
