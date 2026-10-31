"""D03 input: 'Service Level Agreement' 5 times: one plain run between bold runs, split across italic runs with
revision ids and a proofing mark, split across a bold and a plain run, split across red runs in a table cell,
split across runs in the header. A lowercase near-miss must stay."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from authoring_inputs import add_runs, declare_rsids, new_document, para, proof_error_before  # noqa: E402

doc = new_document("Contract summary")
doc.add_heading("Contract summary", 1)
para(doc, [("Scope: ", "b"), ("The ", ""), ("Service Level Agreement", ""), (" applies to ", ""),
           ("all customers", "b"), (".", "")])
runs = add_runs(doc.add_paragraph(), [("Under the ", "i"), ("Service Le", "i"), ("vel Agree", "i"), ("ment", "i"),
                                      (", response times are ", "i"), ("guaranteed", "bi"), (".", "i")])
proof_error_before(runs[2])
proof_error_before(runs[4], "spellEnd")
para(doc, [("Service ", "b"), ("Level Agreement", ""), (" terms are reviewed yearly.", "")])
para(doc, [("The service level is measured monthly.", "")])
table = doc.add_table(rows=2, cols=2, style="Table Grid")
table.cell(0, 0).text, table.cell(0, 1).text = "Document", "Owner"
add_runs(table.cell(1, 0).paragraphs[0], [("Service Level ", "r"), ("Agreement", "r")])
table.cell(1, 1).text = "Legal"
para(doc, [("End of summary.", "")])
add_runs(doc.sections[0].header.paragraphs[0], [("Service Level ", ""), ("Agreement - draft for review", "")])
declare_rsids(doc)
doc.save(Path(sys.argv[1]) / "input.docx")
