"""D07 inputs: a document with one inline picture (400 x 300 px shown at 8 x 6 cm) and an anchor paragraph,
the picture to insert (new.png, 300 x 200 px) and the replacement image (replacement.png, 200 x 150 px)."""
import io
import sys
from pathlib import Path

from docx.shared import Cm

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from authoring_inputs import declare_rsids, new_document, para, png_bytes  # noqa: E402

out = Path(sys.argv[1])
doc = new_document("Site report")
doc.add_heading("Site report", 1)
para(doc, [("Figure 1 shows the ", ""), ("site", "b"), (" as surveyed.", "")])
doc.add_paragraph().add_run().add_picture(io.BytesIO(png_bytes(400, 300, (31, 78, 121))), width=Cm(8), height=Cm(6))
para(doc, [("Figure 1: site plan.", "i")])
para(doc, [("Figure placement ", ""), ("anchor", "b"), (".", "")])
para(doc, [("Closing paragraph.", "")])
declare_rsids(doc)
doc.save(out / "input.docx")
(out / "new.png").write_bytes(png_bytes(300, 200, (192, 0, 0)))
(out / "replacement.png").write_bytes(png_bytes(200, 150, (0, 128, 64)))
