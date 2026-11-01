"""P01 inputs: content.json (the deck's theme and its 10 slides, with speaker notes) and photo.png (4:3)."""
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "lib"))
import inputs  # noqa: E402

out = Path(sys.argv[1])
out.mkdir(parents=True, exist_ok=True)
(out / "photo.png").write_bytes(inputs.png(800, 600, 1, "hq"))
theme = inputs.NORTHWIND
content = {
    "theme": {"name": theme["name"], "headFont": theme["major"], "bodyFont": theme["minor"], "colors": theme["colors"]},
    "slides": [
        {"kind": "title", "title": "Northwind 2026 strategy", "subtitle": "Board meeting, 12 November 2026"},
        {"kind": "bullets", "title": "Agenda", "bullets": ["Where we stand", "Results by region", "Revenue trend",
                                                          "Our new headquarters", "Plan for 2027"]},
        {"kind": "section", "title": "Where we stand", "subtitle": "Results to the end of September"},
        {"kind": "table", "title": "Results by region",
         "table": [["Region", "Revenue (EUR k)", "Plan", "Delta"], ["North", "1,240", "1,150", "+8 %"],
                   ["South", "980", "1,000", "-2 %"], ["West", "760", "700", "+9 %"]]},
        {"kind": "chart", "title": "Revenue trend",
         "chart": {"categories": ["Q1", "Q2", "Q3", "Q4"],
                   "series": [{"name": "North", "values": [1150, 1190, 1240, 1300]},
                              {"name": "South", "values": [940, 960, 980, 1010]}]}},
        {"kind": "picture", "title": "Our new headquarters", "image": "photo.png", "caption": "Opening March 2027"},
        {"kind": "two_columns", "title": "Strengths and risks",
         "left": ["Loyal customers", "Strong brand in the North"], "right": ["Price pressure in the South",
                                                                              "Hiring delays"]},
        {"kind": "bullets", "title": "Plan for 2027", "bullets": ["Open two regional offices", "Launch the loyalty app",
                                                                 "Cut delivery time to 48 hours"]},
        {"kind": "bullets", "title": "What we ask the board", "bullets": ["Approve the 2027 budget",
                                                                         "Approve the headquarters move"]},
        {"kind": "closing", "title": "Thank you", "subtitle": "Questions and discussion"},
    ],
}
for k, s in enumerate(content["slides"], 1):
    s["notes"] = f"Notes {k}: speak to '{s['title']}' for two minutes."
(out / "content.json").write_text(json.dumps(content, indent=1))
