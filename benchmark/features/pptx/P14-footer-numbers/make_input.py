"""P14 input: a 6-slide deck (title slide, bullets, two columns, table, section header, bullets) whose layouts
carry date, footer and slide-number placeholders; no slide has any of them yet."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "lib"))
import inputs  # noqa: E402

out = Path(sys.argv[1])
out.mkdir(parents=True, exist_ok=True)
prs = inputs.base_deck()
s = prs.slides.add_slide(inputs.layout(prs, "title"))
s.shapes.title.text = "Northwind quarterly review"
s.placeholders[1].text = "Q3 2026"
for key, title in [("content", "Agenda"), ("two", "Market view"), ("title_only", "Results"),
                   ("section", "Delivery"), ("content", "Next steps")]:
    s = prs.slides.add_slide(inputs.layout(prs, key))
    s.shapes.title.text = title
    for ph in s.placeholders:
        if ph.placeholder_format.idx in (1, 2):
            inputs.bullets(ph.text_frame, [(0, [f"{title}: first point"]), (0, ["Second point"])])
    if key == "title_only":
        inputs.table(s.shapes, 0.8, 1.8, 11.7, 2.0, [["Region", "Revenue"], ["North", "1,240"], ["South", "980"]])
inputs.save(prs, out / "deck.pptx")
