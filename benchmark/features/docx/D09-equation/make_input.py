"""D09 input: a short text with the anchor paragraph the equation goes after."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from authoring_inputs import declare_rsids, new_document, para  # noqa: E402

doc = new_document("Quadratic equations")
doc.add_heading("Quadratic equations", 1)
para(doc, [("An equation ", ""), ("ax\u00b2 + bx + c = 0", "i"), (" has at most two real roots.", "")])
para(doc, [("The roots are given by:", "")])
para(doc, [("where ", ""), ("a", "i"), (" is not zero.", "")])
declare_rsids(doc)
doc.save(Path(sys.argv[1]) / "input.docx")
