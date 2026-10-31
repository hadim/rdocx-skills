"""P05 inputs: a.pptx (theme Northwind, 4 slides) and b.pptx (theme Harbor: other colours and fonts, its own
master with its own logo, layouts renamed "Harbor ..."; 4 slides, slide 2 a native chart with notes, slide 3 a
picture with a caption)."""
import io
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "lib"))
import inputs  # noqa: E402
from inputs import Inches  # noqa: E402

out = Path(sys.argv[1])
out.mkdir(parents=True, exist_ok=True)
a = inputs.base_deck(inputs.NORTHWIND)
for key, title in [("title", "A1: Northwind overview"), ("content", "A2: Agenda"), ("title_only", "A3: Results"),
                   ("section", "A4: Close")]:
    s = a.slides.add_slide(inputs.layout(a, key))
    s.shapes.title.text = title
    if key == "content":
        inputs.bullets(s.placeholders[1].text_frame, [(0, ["Results"]), (0, ["Harbor partnership"]), (0, ["Close"])])
    if key == "title_only":
        inputs.textbox(s.shapes, 0.8, 1.8, 6, 1.5, ["Revenue ", ("up 12 %", {"bold": True})], size=24)
inputs.save(a, out / "a.pptx")

names = ["Harbor " + n for n in ("Title", "Content", "Section", "Two Content", "Comparison", "Title Only", "Blank",
                                 "Content Caption", "Picture Caption", "Vertical Text", "Vertical Title")]
b = inputs.base_deck(inputs.HARBOR, layout_names=names)
photo = inputs.png(800, 600, 1, "harbor")
(out / "harbor.png").write_bytes(photo)
s = b.slides.add_slide(inputs.layout(b, "title"))
s.shapes.title.text = "B1: Harbor partnership"
s = b.slides.add_slide(inputs.layout(b, "title_only"))
s.shapes.title.text = "B2: Harbor volumes"
inputs.chart(s.shapes, 0.8, 1.5, 8.5, 5.0, ["2023", "2024", "2025", "2026"],
             [("Containers", (410, 455, 520, 580)), ("Bulk", (220, 210, 240, 260))], name="Volumes chart",
             title="Volumes (k units)", colors=["0B6E4F", "F18F01"])
inputs.notes(s, "B2 notes: volumes from the port authority.")
s = b.slides.add_slide(inputs.layout(b, "picture"))
s.shapes.title.text = "B3: Harbor site"
s.placeholders[1].insert_picture(io.BytesIO(photo))
s.placeholders[2].text = "The new quay, September 2026"
s = b.slides.add_slide(inputs.layout(b, "section"))
s.shapes.title.text = "B4: Harbor close"
inputs.save(b, out / "b.pptx")
