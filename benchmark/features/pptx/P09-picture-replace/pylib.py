# bench: tools=python-pptx; processes=1; raw_xml=yes
import sys

from pptx import Presentation

prs = Presentation(sys.argv[1])
slide = prs.slides[2]
pic = next(sh for sh in slide.shapes if sh.name == "Site photo")
blip = pic._element.blipFill.blip  # python-pptx has no replace-image call: point the blip at a new image part
old_rid = blip.rEmbed
_, rid = slide.part.get_or_add_image_part(sys.argv[2])
blip.rEmbed = rid
slide.part.drop_rel(old_rid)
prs.save(sys.argv[3])
