# bench: tools=python-pptx; processes=1; raw_xml=no
import sys

from pptx import Presentation

prs = Presentation(sys.argv[1])
notes = {1: "Spend two minutes on the agenda.",
         2: "Updated: the survey closed on 30 September.\nSource: field team, 412 answers.",
         3: "Stress the regional split.", 5: "Close with the three asks."}
for k, text in notes.items():
    prs.slides[k].notes_slide.notes_text_frame.text = text
prs.save(sys.argv[2])
