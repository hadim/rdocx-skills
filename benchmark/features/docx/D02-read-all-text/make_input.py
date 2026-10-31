"""D02 input: one marker string in each story (body, table cell, header, footer, footnote, text box, comment)."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from authoring_inputs import add_comment, add_footnote, add_runs, add_text_box, declare_rsids, new_document, para, proof_error_before  # noqa: E402

doc = new_document("Read test")
doc.add_heading("Read test", 1)
p = para(doc, [("Intro text with ", ""), ("MARK_BO", "b"), ("DY_1", "i"), (" inside.", "")])
table = doc.add_table(rows=2, cols=2, style="Table Grid")
table.cell(0, 0).text = "Cell A"
cell = table.cell(0, 1).paragraphs[0]
runs = add_runs(cell, [("MARK_TA", "b"), ("BLE_1", "")])
proof_error_before(runs[1])
table.cell(1, 0).text = "Cell C"
table.cell(1, 1).text = "Cell D"
closing = para(doc, [("Closing ", "b"), ("MARK_BODY_2", ""), (" paragraph.", "i")])
add_footnote(doc, closing, "MARK_FOOTNOTE note text.")
add_comment(doc, closing.runs[0], "MARK_COMMENT remark")
anchor = para(doc, [("Anchor of the box.", "")])
add_text_box(anchor, "MARK_TEXTBOX in a box")
para(doc, [("Last body paragraph.", "")])
section = doc.sections[0]
section.header.paragraphs[0].text = "MARK_HEADER running head"
section.footer.paragraphs[0].text = "MARK_FOOTER running foot"
declare_rsids(doc)
doc.save(Path(sys.argv[1]) / "input.docx")
