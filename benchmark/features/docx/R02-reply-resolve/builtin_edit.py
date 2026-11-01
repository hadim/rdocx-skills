"""Nest the reply's markers in the parent's (as comment.py prints), mark the parent done, set the date."""
import re
import sys
from pathlib import Path

u, parent, cid = Path(sys.argv[1]), sys.argv[2], sys.argv[3]
w = lambda name: u / "word" / name  # noqa: E731
doc = w("document.xml").read_text(encoding="utf-8")
ref = f'<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="{cid}"/></w:r>'
doc = doc.replace(f'<w:commentRangeStart w:id="{parent}"/>',
                  f'<w:commentRangeStart w:id="{parent}"/><w:commentRangeStart w:id="{cid}"/>')
doc = doc.replace(f'<w:commentRangeEnd w:id="{parent}"/>',
                  f'<w:commentRangeEnd w:id="{cid}"/><w:commentRangeEnd w:id="{parent}"/>')
doc = re.sub(rf'(<w:commentReference w:id="{parent}"/></w:r>)', lambda m: m.group(1) + ref, doc)
w("document.xml").write_text(doc, encoding="utf-8")
comments = w("comments.xml").read_text(encoding="utf-8")
# resolve: w15:done="1" on the parent's commentEx (keyed by the paraId of its last paragraph)
block = re.search(rf'<w:comment [^>]*w:id="{parent}"[^>]*>.*?</w:comment>', comments, re.S).group(0)
para_id = re.findall(r'w14:paraId="([0-9A-F]+)"', block)[-1]
ext = w("commentsExtended.xml").read_text(encoding="utf-8")
ext = re.sub(rf'(<w15:commentEx w15:paraId="{para_id}"[^>]*w15:done=")0"', r'\g<1>1"', ext)
w("commentsExtended.xml").write_text(ext, encoding="utf-8")
# comment.py dates the reply now: write the requested date where it put the timestamp
stamp = re.search(rf'<w:comment w:id="{cid}" [^>]*w:date="([^"]*)"', comments).group(1)
for name in ("comments.xml", "commentsExtensible.xml"):
    w(name).write_text(w(name).read_text(encoding="utf-8").replace(stamp, "2026-10-30T10:00:00Z"), encoding="utf-8")
