"""D12 input: two 'Note:' paragraphs in Normal and two numbered lists that share one Word list (numId 10), so
that the second continues the first (4, 5, 6) until it is restarted."""
import sys
from pathlib import Path

from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from authoring_inputs import declare_rsids, new_document, para  # noqa: E402

doc = new_document("Procedure")
numbering = doc.part.numbering_part.element
levels = "".join(
    f'<w:lvl w:ilvl="{i}"><w:start w:val="1"/><w:numFmt w:val="{"decimal" if i % 3 == 0 else "lowerLetter" if i % 3 == 1 else "lowerRoman"}"/>'
    f'<w:lvlText w:val="%{i + 1}."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="{720 * (i + 1)}" w:hanging="360"/></w:pPr></w:lvl>'
    for i in range(9))
abstract = parse_xml(f'<w:abstractNum {nsdecls("w")} w:abstractNumId="20"><w:nsid w:val="1A2B3C4D"/>'
                     f'<w:multiLevelType w:val="hybridMultilevel"/><w:tmpl w:val="0409000F"/>{levels}</w:abstractNum>')
numbering.find(qn("w:num")).addprevious(abstract)  # abstractNum elements come before num elements
numbering.append(parse_xml(f'<w:num {nsdecls("w")} w:numId="10"><w:abstractNumId w:val="20"/></w:num>'))


def item(text):
    p = para(doc, [(text, "")], style="List Paragraph")
    p._p.get_or_add_pPr().append(parse_xml(f'<w:numPr {nsdecls("w")}><w:ilvl w:val="0"/><w:numId w:val="10"/></w:numPr>'))


doc.add_heading("Procedure", 1)
para(doc, [("Note:", "b"), (" read all steps first.", "")])
for text in ("Open the valve.", "Check the pressure.", "Close the valve."):
    item(text)
para(doc, [("Then repeat for the second circuit:", "")])
for text in ("Open the bypass.", "Check the flow.", "Close the bypass."):
    item(text)
para(doc, [("Note:", "b"), (" log every reading.", "")])
para(doc, [("End of procedure.", "")])
declare_rsids(doc)
doc.save(Path(sys.argv[1]) / "input.docx")
