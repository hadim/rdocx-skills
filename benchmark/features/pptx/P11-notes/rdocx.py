# bench: tools=rpptx (Python), pptx_ops.save_atomic; processes=1; raw_xml=no
import sys

import pptx_ops
import rpptx

prs = rpptx.Presentation(sys.argv[1])
prs.slides[1].notes_text = "Spend two minutes on the agenda."
prs.slides[2].notes_text = "Updated: the survey closed on 30 September.\nSource: field team, 412 answers."
prs.slides[3].notes_text = "Stress the regional split."
prs.slides[5].notes_text = "Close with the three asks."
pptx_ops.save_atomic(prs, sys.argv[2], sys.argv[1])
