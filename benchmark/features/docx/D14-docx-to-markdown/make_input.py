"""D14 input: headings, bold and italic runs, a bulleted and a numbered list made as Word makes them (List
Paragraph with w:numPr to its own lists), a table with a header row."""
import sys
from pathlib import Path

from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from authoring_inputs import declare_rsids, new_document, para  # noqa: E402

doc = new_document("Field report")
numbering = doc.part.numbering_part.element
for aid, fmt, text in ((30, "bullet", "\u2022"), (31, "decimal", "%1.")):
    numbering.find(qn("w:num")).addprevious(parse_xml(
        f'<w:abstractNum {nsdecls("w")} w:abstractNumId="{aid}"><w:multiLevelType w:val="hybridMultilevel"/>'
        f'<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="{fmt}"/><w:lvlText w:val="{text}"/>'
        '<w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum>'))
    numbering.append(parse_xml(f'<w:num {nsdecls("w")} w:numId="{aid - 19}"><w:abstractNumId w:val="{aid}"/></w:num>'))


def item(text, num_id):
    p = para(doc, [(text, "")], style="List Paragraph")
    p._p.get_or_add_pPr().append(parse_xml(f'<w:numPr {nsdecls("w")}><w:ilvl w:val="0"/><w:numId w:val="{num_id}"/></w:numPr>'))


doc.add_heading("Field report", 1)
para(doc, [("The inspection found ", ""), ("three defects", "b"), (" and one ", ""), ("open question", "i"), (".", "")])
doc.add_heading("Findings", 2)
for text in ("Cracked kerb", "Loose railing", "Blocked drain"):
    item(text, 11)
doc.add_heading("Actions", 2)
for text in ("Repair the kerb", "Fix the railing", "Clear the drain"):
    item(text, 12)
doc.add_heading("Costs", 3)
table = doc.add_table(rows=3, cols=3, style="Table Grid")
for r, row in enumerate([("Item", "Hours", "Cost"), ("Kerb", "6", "480"), ("Railing", "3", "210")]):
    for c, text in enumerate(row):
        table.cell(r, c).text = text
        if r == 0:
            table.cell(r, c).paragraphs[0].runs[0].bold = True
para(doc, [("Report closed.", "")])
declare_rsids(doc)
doc.save(Path(sys.argv[1]) / "input.docx")
