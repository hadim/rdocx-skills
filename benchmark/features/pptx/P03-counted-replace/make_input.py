"""P03 input: a 5-slide deck where "Project Falcon" occurs 7 times (title, bold run, a match split over a bold and
an italic red run, notes, a table cell, a group child split mid-word over an underlined run, a match split over
three runs) and twice in other cases ("project falcon", "PROJECT FALCON") as decoys."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "lib"))
import inputs  # noqa: E402
from inputs import Inches  # noqa: E402

out = Path(sys.argv[1])
out.mkdir(parents=True, exist_ok=True)
prs = inputs.base_deck()
s = prs.slides.add_slide(inputs.layout(prs, "title"))
s.shapes.title.text = "Project Falcon update"
inputs.runs(s.placeholders[1].text_frame.paragraphs[0], ["Steering committee, ", ("October 2026", {"italic": True})])

s = prs.slides.add_slide(inputs.layout(prs, "content"))
s.shapes.title.text = "Status"
inputs.bullets(s.placeholders[1].text_frame, [
    (0, ["Status of ", ("Project Falcon", {"bold": True}), " is green"]),
    (0, ["Budget for ", ("Project ", {"bold": True}), ("Falcon", {"italic": True, "color": "C8553D"}), " approved"]),
    (1, ["The project falcon team grew to 12 people"]),
])
inputs.notes(s, "Remind them that Project Falcon started in May.")

s = prs.slides.add_slide(inputs.layout(prs, "title_only"))
s.shapes.title.text = "Phases"
inputs.table(s.shapes, 0.8, 1.8, 11.7, 2.4, [["Phase", "Scope", "Due"],
                                             ["1", "Pilot in two regions", "December"],
                                             ["2", "Project Falcon phase 2", "June"]], name="Phases table")

s = prs.slides.add_slide(inputs.layout(prs, "title_only"))
s.shapes.title.text = "Governance"
g = s.shapes.add_group_shape()
g.name = "Org"
inputs.textbox(g.shapes, 0.8, 2.0, 5.5, 1.0, ["Owner: ", ("Proj", {"underline": True}), "ect Falcon PMO"], size=20,
               name="Owner")
inputs.textbox(g.shapes, 6.8, 2.0, 5.5, 1.0, ["Sponsor: Chief Operating Officer"], size=20, name="Sponsor")

s = prs.slides.add_slide(inputs.layout(prs, "content"))
s.shapes.title.text = "Next steps"
inputs.bullets(s.placeholders[1].text_frame, [
    (0, ["Close ", ("Project Fal", {"bold": True}), ("con", {"bold": True, "italic": True}), " phase 1"]),
    (0, ["Plan the regional roll-out"]),
])
inputs.notes(s, "PROJECT FALCON is the internal codename; do not show it to customers.")
inputs.save(prs, out / "deck.pptx")
