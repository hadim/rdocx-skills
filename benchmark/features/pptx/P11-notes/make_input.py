"""P11 input: a 6-slide deck with speaker notes on slides 1, 3 and 5 only."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "lib"))
import inputs  # noqa: E402

out = Path(sys.argv[1])
out.mkdir(parents=True, exist_ok=True)
prs = inputs.base_deck()
titles = ["Quarterly review", "Agenda", "Market view", "Regional results", "Delivery", "Next steps"]
for k, title in enumerate(titles):
    s = prs.slides.add_slide(inputs.layout(prs, "title" if k == 0 else "content"))
    s.shapes.title.text = title
    if k:
        inputs.bullets(s.placeholders[1].text_frame, [(0, [f"{title}: first point"]), (0, ["Second point"])])
inputs.notes(prs.slides[0], "Welcome everyone.")
inputs.notes(prs.slides[2], "Old notes: survey still open.")
inputs.notes(prs.slides[4], "Walk through the stages.")
inputs.save(prs, out / "deck.pptx")
