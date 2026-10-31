# bench: tools=rpptx (Python), pptx_ops.save_atomic; processes=1; raw_xml=no
import sys

import pptx_ops
import rpptx

prs = rpptx.Presentation(sys.argv[1])
dana, sam = "{6F1B3A52-0000-4000-8000-000000000001}", "{6F1B3A52-0000-4000-8000-000000000002}"
prs.add_comment_author(id=dana, name="Dana Reviewer", user_id="Dana Reviewer", provider_id="None", initials="DR")
prs.add_comment_author(id=sam, name="Sam Author", user_id="Sam Author", provider_id="None", initials="SA")
slide = prs.slides[2]
key = next(sh for sh in slide.shapes if sh.name == "Key figure")
cid = "{6F1B3A52-0000-4000-8000-0000000000A1}"
slide.add_comment(id=cid, author_id=dana, created="2026-10-01T12:00:00Z", text="Please check this figure.",
                  shape_id=key.shape_id)
slide.reply_to_comment(cid, id="{6F1B3A52-0000-4000-8000-0000000000A2}", author_id=sam,
                       created="2026-10-01T12:05:00Z", text="Updated with Q3 numbers.")
pptx_ops.save_atomic(prs, sys.argv[2], sys.argv[1])
