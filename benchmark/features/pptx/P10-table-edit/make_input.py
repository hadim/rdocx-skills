"""P10 input: a 3-slide deck; slide 2 holds a 4x4 table (Medium Style 2 - Accent 1, header row, 14 pt runs, one
cell with its own green fill)."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "lib"))
import inputs  # noqa: E402

out = Path(sys.argv[1])
out.mkdir(parents=True, exist_ok=True)
prs = inputs.base_deck()
s = prs.slides.add_slide(inputs.layout(prs, "title"))
s.shapes.title.text = "Northwind quarterly review"
s = prs.slides.add_slide(inputs.layout(prs, "title_only"))
s.shapes.title.text = "Regional results"
inputs.table(s.shapes, 0.8, 1.7, 11.7, 2.4, [["Region", "Revenue", "Plan", "Delta"],
                                             ["North", "1,240", "1,150", "+8 %"],
                                             ["South", "980", "1,000", "-2 %"],
                                             ["West", "760", "700", "+9 %"]],
             name="Results table", fills={(1, 3): "C6E0B4"})
s = prs.slides.add_slide(inputs.layout(prs, "title_only"))
s.shapes.title.text = "Next steps"
inputs.save(prs, out / "deck.pptx")
