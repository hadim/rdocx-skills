"""D10 input: a 'Further reading: ' paragraph to append links to and a heading 'Appendix A' to link to."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from authoring_inputs import declare_rsids, go_back_bookmark, new_document, para  # noqa: E402

doc = new_document("Project note")
doc.add_heading("Introduction", 1)
para(doc, [("The project collects ", ""), ("field data", "i"), (" from three regions.", "")])
p = para(doc, [("Further reading:", "b"), (" ", "")])
go_back_bookmark(p)
doc.add_heading("Method", 1)
para(doc, [("Samples were taken ", ""), ("monthly", "b"), (".", "")])
doc.add_heading("Appendix A", 1)
para(doc, [("Supporting data.", "")])
declare_rsids(doc)
doc.save(Path(sys.argv[1]) / "input.docx")
