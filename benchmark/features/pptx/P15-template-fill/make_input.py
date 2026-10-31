"""P15 inputs: template.potx (the Northwind master and layouts, 16:9, one instruction slide) and photo.png (4:3,
while the layout's picture placeholder is 16:9, so filling it means cropping)."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "lib"))
import inputs  # noqa: E402

out = Path(sys.argv[1])
out.mkdir(parents=True, exist_ok=True)
prs = inputs.base_deck()
s = prs.slides.add_slide(inputs.layout(prs, "content"))
s.shapes.title.text = "How to use this template"
inputs.bullets(s.placeholders[1].text_frame, [(0, ["Use the Title Slide layout first"]),
                                              (0, ["Delete this slide before sending"])])
inputs.save(prs, out / "template.pptx")
inputs.to_template(out / "template.pptx", out / "template.potx")
(out / "template.pptx").unlink()
(out / "photo.png").write_bytes(inputs.png(800, 600, 1, "hq"))
