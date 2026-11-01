# bench: tools=rpptx (Python), pptx_ops.save_atomic; processes=1; raw_xml=no
import sys

import pptx_ops
import rpptx

prs = rpptx.Presentation(sys.argv[1])
alpha, bravo, charlie, delta, echo, foxtrot, golf, hotel = list(prs.slides)  # handles follow their slides
copy = prs.slides.duplicate(charlie)  # content and notes, inserted after Charlie


def position(slide):
    return list(prs.slides).index(slide)


prs.slides.move(position(copy), position(echo))  # `to` is the final index: Echo moves up as the copy leaves
prs.slides.move(position(hotel), 1)
prs.slides.remove(delta)
foxtrot.hidden = True
pptx_ops.save_atomic(prs, sys.argv[2], sys.argv[1])
