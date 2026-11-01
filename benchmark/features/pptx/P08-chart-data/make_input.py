"""P08 input: a 4-slide deck; slide 3 holds a formatted clustered-column chart (2 series North/South over 4
quarters, series fills in theme-like colours, data labels with a number format, a title, a bottom legend) with its
embedded workbook."""
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "lib"))
import inputs  # noqa: E402

out = Path(sys.argv[1])
out.mkdir(parents=True, exist_ok=True)
prs = inputs.base_deck()
s = prs.slides.add_slide(inputs.layout(prs, "title"))
s.shapes.title.text = "Northwind quarterly review"
for title in ["Agenda", "Revenue by region", "Next steps"]:
    s = prs.slides.add_slide(inputs.layout(prs, "title_only"))
    s.shapes.title.text = title
    if title == "Revenue by region":
        inputs.chart(s.shapes, 0.8, 1.5, 8.8, 5.2, ["Q1 26", "Q2 26", "Q3 26", "Q4 26"],
                     [("North", (1240, 1310, 1405, 1520)), ("South", (980, 1010, 990, 1060))],
                     name="Revenue chart", title="Revenue by region (EUR k)", colors=["1E2761", "C8553D"])
        inputs.textbox(s.shapes, 9.9, 2.0, 2.8, 2.0, ["North keeps growing"], size=16, name="Comment")
inputs.save(prs, out / "deck.pptx")
(out / "new_data.json").write_text(json.dumps({
    "categories": ["Q1 26", "Q2 26", "Q3 26", "Q4 26", "Q1 27"],
    "series": {"North": [1250, 1330, 1410, 1535, 1600], "South": [990, 1005, 1000, 1070, 1110]}}))
