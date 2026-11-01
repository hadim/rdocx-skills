# bench: {"tools": ["rdocx Python (insert_document)"], "processes": 1, "raw_xml": false, "notes": "insert_document renames B's conflicting styles (Heading1Merge1), renumbers its list and copies its picture, but drops B's w:docDefaults: B's body text falls back to A's Calibri instead of Georgia (NOTES_review.md, finding 3)"}
import sys

import docx_ops
import rdocx

a, b = rdocx.Document.open(sys.argv[1]), rdocx.Document.open(sys.argv[2])
a.insert_document(b, at=len(a.paragraphs) + len(a.tables))  # the end of a's body
docx_ops.save_atomic(a, sys.argv[3], sys.argv[1])
