# bench: {"tools": ["rdocx Python"], "processes": 1, "raw_xml": false, "notes": "reply_to threads the reply in commentsExtended.xml and resolves the parent, but writes no range or reference mark for the reply in document.xml (nor a commentsIds entry, nor initials: reply_to has no initials argument); LibreOffice drops the reply on import, and Word writes a reference for every reply (NOTES_review.md, finding 1)"}
import sys

import docx_ops
import rdocx

doc = rdocx.Document.open(sys.argv[1])
cid = next(c.id for c in doc.comments if c.author == "Alex Author" and c.parent_id is None)
doc.reply_to(cid, author="Dana Reviewer", text="Agreed, changed to 45 days.", date="2026-10-30T10:00:00Z")
doc.resolve_comment(cid)
docx_ops.save_atomic(doc, sys.argv[2], sys.argv[1])
