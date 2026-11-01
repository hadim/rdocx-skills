"""P04 input: 8 slides titled Alpha..Hotel, each with notes; slide 3 (Charlie) and slide 4 (Delta) each show a
picture of their own, so deleting Delta must remove its slide, notes slide and image, while the copy of Charlie
keeps Charlie's image and notes. truth.json gives the expected final order."""
import io
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "lib"))
import inputs  # noqa: E402
from inputs import Inches  # noqa: E402

out = Path(sys.argv[1])
out.mkdir(parents=True, exist_ok=True)
names = ["Alpha", "Bravo", "Charlie", "Delta", "Echo", "Foxtrot", "Golf", "Hotel"]
pics = {"Charlie": inputs.png(640, 400, 0, "charlie"), "Delta": inputs.png(640, 400, 1, "delta")}
prs = inputs.base_deck()
for k, name in enumerate(names):
    s = prs.slides.add_slide(inputs.layout(prs, "title" if k == 0 else "content" if name not in pics else "title_only"))
    s.shapes.title.text = f"{name}: section {k + 1}"
    if name in pics:
        s.shapes.add_picture(io.BytesIO(pics[name]), Inches(0.8), Inches(1.6), width=Inches(6.0)).name = f"{name} figure"
        inputs.textbox(s.shapes, 7.2, 1.8, 5.0, 2.0, [f"{name} ", ("highlights", {"bold": True})], size=20)
    elif k:
        inputs.bullets(s.placeholders[1].text_frame, [(0, [f"{name} point one"]), (1, ["detail"])])
    inputs.notes(s, f"notes-{name.lower()}: speak to {name}.")
for name, data in pics.items():
    (out / f"{name.lower()}.png").write_bytes(data)
inputs.save(prs, out / "deck.pptx")
(out / "truth.json").write_text(json.dumps({
    "order": ["Alpha", "Hotel", "Bravo", "Charlie", "Echo", "Charlie", "Foxtrot", "Golf"],
    "hidden": ["Foxtrot"], "copy_position": 6, "deleted": "Delta"}))
