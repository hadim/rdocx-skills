# bench: tools=rpptx (Python), pptx_ops.save_atomic; processes=1; raw_xml=no
import sys

import pptx_ops
import rpptx

prs = rpptx.Presentation(sys.argv[1])
pic = next(sh for sh in prs.slides[2].shapes if sh.name == "Site photo")
pic.replace_image(sys.argv[2])
pptx_ops.save_atomic(prs, sys.argv[3], sys.argv[1])
