"""D08 input: a short report with the anchor paragraph the chart goes after."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from authoring_inputs import declare_rsids, new_document, para  # noqa: E402

doc = new_document("Sales review")
doc.add_heading("Sales review", 1)
para(doc, [("Sales grew in ", ""), ("three", "b"), (" quarters out of four.", "")])
para(doc, [("Chart anchor: ", "b"), ("sales by quarter.", "")])
para(doc, [("The fourth quarter of 2026 was the ", ""), ("strongest", "i"), (".", "")])
declare_rsids(doc)
doc.save(Path(sys.argv[1]) / "input.docx")
