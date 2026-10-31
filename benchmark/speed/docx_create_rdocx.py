"""Create the corpus report from its plan with rdocx (the rdocx stack's creation path).

    $R/python docx_create_rdocx.py plan.json out.docx [--naive]      (docx_ops on PYTHONPATH)

Same content and features as corpus/gen_docx.js (docx-js): A4, Calibri body and Cambria headings, a title page,
a table of contents (filled with page numbers by rebuild_toc), headings on three levels, paragraphs of formatted
runs, bullet lists on two levels and numbered lists restarting per list, tables with column widths, a shaded
header row, zebra rows, merged cells and borders, pictures with alt text and captions, footnotes, comments,
tracked insertions and deletions (raw XML through replace_xml: rdocx has no API that writes a revision), a
header, a "Page X of Y" footer, and a landscape second section that inherits header and footer.

How, following the rdocx skill (references/gaps.md, edit costs): each chapter is built in a small document of its
own, a copy of the styled empty document, and inserted with one insert_document, since an edit costs more in a
larger body. Workarounds for this build (speed/NOTES.md, rdocx findings): the table cells' paragraph styles are
set by one replace_xml per table (every call through a cell's paragraph or run handle costs time that grows with
the square of the table's cells); every list instance is made up front (a definition id can go stale); numbered
lists get their instance again after each insert_document (it merges equal instances, so restarts were lost); the
table of contents is inserted last (one inserted before the headings exist cannot be rebuilt).

--naive: the same document built straight into one body, cell paragraphs styled through their handles, no
chunks: what the API invites first; minutes from tens of cells on (measured on docx-S only).
"""
import json
import sys
from xml.sax.saxutils import escape

import docx_ops
import rdocx
from rdocx import Emu, Pt, Twips
from rdocx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from rdocx.enum.text import WD_ALIGN_PARAGRAPH

plan_path, out = sys.argv[1:3]
NAIVE = "--naive" in sys.argv[3:]
plan = json.load(open(plan_path, encoding="utf-8"))
HEAD_FILL, ZEBRA, TOTAL_FILL = "1F4E79", "EAF1F8", "D9E2F3"
comments = {c["id"]: c for c in plan["comments"]}

# --- the styled, empty document every part starts from -------------------------------------------------------
base = rdocx.Document()
base.update_section(0, page_width=Twips(11906), page_height=Twips(16838), margin_top=Twips(1440),
                    margin_bottom=Twips(1440), margin_left=Twips(1300), margin_right=Twips(1300))
base.default_font_name = "Calibri"
base.default_font_size = Pt(11)
base.set_default_paragraph_format(space_after=Twips(120), line_spacing=1.1)
base.set_style("Title", font_name="Cambria", font_size=Pt(28), color=HEAD_FILL)
base.set_style("Heading1", font_name="Cambria", font_size=Pt(16), bold=True, color=HEAD_FILL,
               space_before=Twips(360), space_after=Twips(120), keep_with_next=True)
base.set_style("Heading2", font_name="Cambria", font_size=Pt(13), bold=True, color="2E74B5",
               space_before=Twips(240), space_after=Twips(80), keep_with_next=True)
base.set_style("Heading3", font_name="Calibri", font_size=Pt(11), bold=True, italic=True, color="404040",
               space_before=Twips(200), space_after=Twips(60), keep_with_next=True)
base.set_style("Subtitle", font_name="Cambria", font_size=Pt(16), color="595959", space_after=Twips(480))
base.set_style("Caption", font_size=Pt(9), italic=True, color="404040", space_before=Twips(60),
               space_after=Twips(160), keep_with_next=True)
base.add_style("TOC Heading", based_on="Normal", next_style="Normal", font_name="Cambria", font_size=Pt(16),
               bold=True, color=HEAD_FILL, space_after=Twips(240))
STYLE_ID = {
    "Table Text": base.add_style("Table Text", based_on="Normal", font_size=Pt(9), space_after=Twips(0)).style_id,
    "Table Number": base.add_style("Table Number", based_on="Table Text", alignment=WD_ALIGN_PARAGRAPH.RIGHT).style_id,
    "Table Heading": base.add_style("Table Heading", based_on="Table Text", bold=True, color="FFFFFF",
                                    alignment=WD_ALIGN_PARAGRAPH.CENTER).style_id,
    "Table Total": base.add_style("Table Total", based_on="Table Text", bold=True).style_id,
    "Table Total Number": base.add_style("Table Total Number", based_on="Table Total",
                                         alignment=WD_ALIGN_PARAGRAPH.RIGHT).style_id,
}
base.add_style("Figure", based_on="Normal", alignment=WD_ALIGN_PARAGRAPH.CENTER, keep_with_next=True)
bullets = base.add_numbering_definition([
    rdocx.ListLevel(format="bullet", text="\u2022", left_indent=Twips(720), hanging_indent=Twips(360)),
    rdocx.ListLevel(format="bullet", text="\u2013", left_indent=Twips(1440), hanging_indent=Twips(360)),
])
numbers = base.add_numbering_definition([
    rdocx.ListLevel(format="decimal", text="%1.", start=1, left_indent=Twips(720), hanging_indent=Twips(360)),
])
instances = {}  # every list instance up front (NOTES.md: a definition id can go stale)
for section in plan["sections"]:
    for b in section["blocks"]:
        if b["t"] == "li" and (b["list"], b["instance"]) not in instances:
            instances[(b["list"], b["instance"])] = base.add_numbering_instance(numbers if b["list"] == "num" else bullets)
BASE = base.to_bytes()


def rpr(x):
    return ("<w:rPr>" + ("<w:b/>" if x.get("b") else "") + ("<w:i/>" if x.get("i") else "") + "</w:rPr>"
            if x.get("b") or x.get("i") else "")


def run_xml(x, tag="w:t"):
    return f'<w:r>{rpr(x)}<{tag} xml:space="preserve">{escape(x["text"])}</{tag}></w:r>'


class Part:
    """Blocks appended to one document; counts its body blocks and paragraphs, and where its numbered lists are."""

    def __init__(self, doc):
        self.doc, self.blocks, self.flow, self.numbered, self.toc_at = doc, 0, 0, [], None

    def paragraph(self, inl, style=None, numbering=None):
        doc = self.doc
        p = doc.add_paragraph(style=style)
        if numbering:
            p.numbering = numbering
            if numbering[0] in NUMBERED:
                self.numbered.append((self.flow, numbering))
        self.blocks += 1
        self.flow += 1
        if any(x["k"] in ("ins", "del") for x in inl):  # no API writes a revision: the paragraph's XML
            xml = p.xml
            parts = []
            for x in inl:
                if x["k"] == "r":
                    parts.append(run_xml(x))
                elif x["k"] in ("ins", "del"):
                    attrs = f'w:id="{x["id"]}" w:author="{escape(x["author"])}" w:date="{x["date"]}"'
                    tag = "w:t" if x["k"] == "ins" else "w:delText"
                    parts.append(f'<w:{x["k"]} {attrs}>{run_xml(x, tag)}</w:{x["k"]}>')
            if xml.endswith("/>"):
                xml = xml[:-2] + "></w:p>"
            p.replace_xml(xml[:xml.rindex("</w:p>")] + "".join(parts) + "</w:p>")
        else:
            for x in inl:
                if x["k"] == "r":
                    r = p.add_run(x["text"])
                    if x.get("b"):
                        r.font.bold = True
                    if x.get("i"):
                        r.font.italic = True
        # footnote references after the run before them; comments over the runs between their markers
        k, opened = 0, {}
        for x in inl:
            if x["k"] in ("r", "ins"):
                k += 1
            elif x["k"] == "fn":
                doc.add_footnote(p.runs[k - 1] if k else p, plan["footnotes"][str(x["id"])])
                k += 1
            elif x["k"] == "cs":
                opened[x["id"]] = k
            elif x["k"] == "ce":
                c = comments[x["id"]]
                p = doc.paragraphs[-1]  # a comment retires every handle: fetch the paragraph again
                rng = rdocx.StoryRunRange(start=rdocx.StoryRunPosition(paragraph=p, run_index=opened[x["id"]]),
                                          end=rdocx.StoryRunPosition(paragraph=p, run_index=k))
                doc.add_comment(rng, author=c["author"], initials=c["initials"], text=c["text"], date=c["date"])
                p = doc.paragraphs[-1]

    def table(self, spec):
        doc = self.doc
        widths, ncols, nrows = spec["widths"], len(spec["widths"]), len(spec["rows"])
        grid, spans, vmerge, covered = [[None] * ncols for _ in range(nrows)], [], [], set()
        carry = [0] * ncols
        for r, row in enumerate(spec["rows"]):  # each cell on the grid, as docx-js places it
            col = 0
            for c in row["cells"]:
                while col < ncols and carry[col] > 0:
                    carry[col] -= 1
                    col += 1
                span = c.get("colSpan") or 1
                grid[r][col] = c
                if span > 1:
                    spans.append((r, col, span))
                    covered.update((r, k) for k in range(col + 1, col + span))
                if c.get("rowSpan"):
                    vmerge.append((r, col, c["rowSpan"]))
                    for k in range(col, col + span):
                        carry[k] = c["rowSpan"] - 1
                col += span
            while col < ncols and carry[col] > 0:
                carry[col] -= 1
                col += 1
        t = doc.add_table(nrows, ncols)
        self.blocks += 1
        t.grid_widths = [Twips(w) for w in widths]
        t.set_borders("single", size=4, color="8EAADB")
        t.set_cell_margins(top=Twips(40), bottom=Twips(40), left=Twips(80), right=Twips(80))
        cells = [[t.cell(r, c) for c in range(ncols)] for r in range(nrows)]
        styles = [[None] * ncols for _ in range(nrows)]
        for r, row in enumerate(spec["rows"]):
            kind = "Table Heading" if row.get("header") else "Table Total" if row.get("total") else "Table Text"
            fill = HEAD_FILL if row.get("header") else TOTAL_FILL if row.get("total") else ZEBRA if row.get("shade") else None
            for c in range(ncols):
                if (r, c) in covered:  # left empty for set_cell_grid_span
                    continue
                cell, spec_cell = cells[r][c], grid[r][c]
                if spec_cell is not None and spec_cell["text"]:
                    cell.text = spec_cell["text"]
                right = spec_cell is not None and spec_cell.get("align") == "right"
                styles[r][c] = {"Table Text": "Table Number", "Table Total": "Table Total Number"}.get(kind, kind) if right else kind
                if NAIVE:
                    cell.paragraphs[0].style = styles[r][c]
                cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                if fill:
                    cell.shading = fill
        for r, col, n in vmerge:
            for k in range(n):
                t.set_cell_vertical_merge(r + k, col, "restart" if k == 0 else "continue")
        for r, col, span in sorted(spans, key=lambda s: (s[0], -s[1])):  # right to left: cell indices stay put
            doc.tables[-1].set_cell_grid_span(r, col, span)
        doc.tables[-1].rows[0].is_header = True
        if nrows > 1 and spec["rows"][1].get("header"):
            doc.tables[-1].rows[1].is_header = True
        if NAIVE:
            return
        # the cell paragraphs' styles in one raw-XML edit of the table (NOTES.md: cell handles cost O(cells^2))
        span_at = {(r, col): n for r, col, n in spans}
        t = doc.tables[-1]
        xml = t.xml
        pieces, pos = [], 0
        for r in range(nrows):
            col = 0
            while col < ncols:
                tc = xml.index("<w:tc>", pos)
                props = xml.find("</w:tcPr>", tc, xml.index("</w:tc>", tc))
                para = xml.index("<w:p", props if props != -1 else tc)
                end = xml.index(">", para)
                ppr = f'<w:pPr><w:pStyle w:val="{STYLE_ID[styles[r][col]]}"/></w:pPr>'
                if xml[end - 1] == "/":  # an empty paragraph, <w:p/>
                    pieces.append(xml[pos:end - 1] + ">" + ppr + "</w:p>")
                else:
                    pieces.append(xml[pos:end + 1] + ppr)
                pos = end + 1
                col += span_at.get((r, col), 1)
        t.replace_xml("".join(pieces) + xml[pos:])

    def block(self, b):
        doc, t = self.doc, b["t"]
        if t == "title":
            doc.add_heading(b["text"], 0).paragraph_format.space_before = Twips(2400)
        elif t == "subtitle":
            doc.add_paragraph(b["text"], style="Subtitle")
        elif t == "pagebreak":
            doc.add_page_break()
        elif t == "toc":
            doc.add_paragraph("Contents", style="TOC Heading")
            self.toc_at = self.blocks + 1  # the TOC goes in last (NOTES.md)
        elif t == "h":
            doc.add_heading(b["text"], b["level"])
        elif t == "caption":
            doc.add_paragraph(b["text"], style="Caption")
        elif t == "img":
            doc.add_picture(b["path"], b["path"].rsplit("/", 1)[-1], Emu(b["w"] * 9525), Emu(b["h"] * 9525),
                            options=rdocx.PictureOptions(name=b["alt"], description=b["alt"], title=b["alt"]))
            doc.paragraphs[-1].style = "Figure"
        elif t == "p":
            return self.paragraph(b["inl"])
        elif t == "li":
            return self.paragraph(b["inl"], style="List Paragraph", numbering=(instances[(b["list"], b["instance"])], b["level"]))
        elif t == "table":
            return self.table(b)
        else:
            raise ValueError(f"unknown block {t}")
        self.blocks += 1
        self.flow += 1


NUMBERED = {num_id for (kind, _), num_id in instances.items() if kind == "num"}
doc = rdocx.Document.from_bytes(BASE)
head = doc.section(0).header.add_paragraph(plan["meta"]["header"])
head.alignment = WD_ALIGN_PARAGRAPH.RIGHT
run = doc.section(0).header.paragraphs[0].runs[0]
run.font.size, run.font.color = Pt(8), "595959"
foot = doc.section(0).footer.add_paragraph("Page ")
foot.alignment = WD_ALIGN_PARAGRAPH.CENTER
foot.runs[0].add_field("PAGE", "1")
foot = doc.section(0).footer.paragraphs[0]
foot.add_run(" of ")
foot.runs[-1].add_field("NUMPAGES", "1")
for r in doc.section(0).footer.paragraphs[0].runs:
    r.font.size = Pt(9)

main = Part(doc)
toc_at = None
for s, section in enumerate(plan["sections"]):
    if s > 0:
        doc.insert_section(len(doc.sections))
        main.blocks += 1
        main.flow += 1
        doc.update_section(s, orientation="landscape" if section.get("orientation") == "landscape" else "portrait")
    if NAIVE:
        for b in section["blocks"]:
            main.block(b)
            toc_at = main.toc_at if main.toc_at is not None else toc_at
        continue
    chapters, current = [], []
    for b in section["blocks"]:  # one part per chapter: the blocks from one Heading 1 to the next
        if b["t"] == "h" and b["level"] == 1 and current:
            chapters.append(current)
            current = []
        current.append(b)
    chapters.append(current)
    for blocks in chapters:
        part = Part(rdocx.Document.from_bytes(BASE))
        for b in blocks:
            part.block(b)
        if part.toc_at is not None:
            toc_at = main.blocks + part.toc_at
        doc.insert_document(part.doc, at=main.blocks)
        for flow, numbering in part.numbered:  # insert_document merged the equal list instances: set them again
            doc.paragraphs[main.flow + flow].numbering = numbering
        main.blocks += part.blocks
        main.flow += part.flow

doc.core_properties.title = plan["meta"]["title"]
doc.core_properties.author = plan["meta"]["creator"]
doc.core_properties.last_modified_by = plan["meta"]["creator"]
doc.core_properties.comments = "Synthetic benchmark document (fictional organisation)."
if toc_at is not None:
    doc.insert_toc(toc_at, max_level=3)
    doc.rebuild_toc()
docx_ops.save_atomic(doc, out)
