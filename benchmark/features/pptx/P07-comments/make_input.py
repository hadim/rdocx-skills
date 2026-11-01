"""P07 input: a 5-slide deck; slide 3 holds a text box named "Key figure" among other shapes; no comments yet."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "lib"))
import inputs  # noqa: E402

out = Path(sys.argv[1])
out.mkdir(parents=True, exist_ok=True)
prs = inputs.base_deck()
s = prs.slides.add_slide(inputs.layout(prs, "title"))
s.shapes.title.text = "Northwind quarterly review"
for k, title in enumerate(["Agenda", "Revenue", "Costs", "Next steps"], 2):
    s = prs.slides.add_slide(inputs.layout(prs, "title_only"))
    s.shapes.title.text = title
    inputs.textbox(s.shapes, 0.8, 1.8, 5.5, 1.2, [f"{title}: summary"], size=20, name="Summary")
    if k == 3:
        inputs.textbox(s.shapes, 7.0, 1.8, 5.0, 1.2, [("Revenue up 12 %", {"bold": True, "size": 32})],
                       name="Key figure")
        inputs.textbox(s.shapes, 7.0, 3.4, 5.0, 1.0, ["Source: finance, September close"], size=14, name="Source")
inputs.save(prs, out / "deck.pptx")
