"""P09 input: a 4-slide deck; slides 2 and 3 show the same image (one shared media part); on slide 3 it is cropped
on four sides, outlined, named "Site photo" with alt text. new.png has the same 4:3 aspect ratio."""
import io
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "lib"))
import inputs  # noqa: E402
from inputs import Inches, Pt, RGBColor  # noqa: E402

out = Path(sys.argv[1])
out.mkdir(parents=True, exist_ok=True)
old, new = inputs.png(800, 600, 2, "site"), inputs.png(800, 600, 3, "new")
(out / "old.png").write_bytes(old)
(out / "new.png").write_bytes(new)
prs = inputs.base_deck()
s = prs.slides.add_slide(inputs.layout(prs, "title"))
s.shapes.title.text = "Site progress"
s = prs.slides.add_slide(inputs.layout(prs, "title_only"))
s.shapes.title.text = "Overview"
pic = s.shapes.add_picture(io.BytesIO(old), Inches(0.8), Inches(1.6), width=Inches(5.0))
pic.name = "Overview photo"
s = prs.slides.add_slide(inputs.layout(prs, "title_only"))
s.shapes.title.text = "Detail"
pic = s.shapes.add_picture(io.BytesIO(old), Inches(6.2), Inches(1.5), Inches(6.0), Inches(4.2))
pic.name = "Site photo"
pic._element.nvPicPr.cNvPr.set("descr", "Construction site, September")
pic.crop_left, pic.crop_right, pic.crop_top, pic.crop_bottom = 0.10, 0.05, 0.08, 0.12
pic.line.color.rgb = RGBColor(0x1E, 0x27, 0x61)
pic.line.width = Pt(2)
inputs.textbox(s.shapes, 0.8, 1.8, 5.0, 2.0, ["Foundations done, ", ("framing started", {"bold": True})], size=20)
s = prs.slides.add_slide(inputs.layout(prs, "title_only"))
s.shapes.title.text = "Next steps"
inputs.save(prs, out / "deck.pptx")
