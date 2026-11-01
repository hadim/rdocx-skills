"""Place the comment markers around the phrase (comment.py prints them) and set the requested date."""
import re
import sys
from pathlib import Path

u, cid = Path(sys.argv[1]), sys.argv[2]
PHRASE = "thirty days after delivery"
doc = u / "word" / "document.xml"
xml = doc.read_text(encoding="utf-8")
# after merge_runs.py the phrase sits in one run: split it into before / phrase / after
m = re.search(r"<w:r>(<w:rPr>(?:(?!</w:r>).)*?</w:rPr>)?<w:t[^>]*>([^<]*" + re.escape(PHRASE) + r"[^<]*)</w:t></w:r>", xml)
rpr, (before, after) = m.group(1) or "", m.group(2).split(PHRASE, 1)
run = lambda s: f'<w:r>{rpr}<w:t xml:space="preserve">{s}</w:t></w:r>' if s else ""  # noqa: E731
new = (run(before) + f'<w:commentRangeStart w:id="{cid}"/>' + run(PHRASE) + f'<w:commentRangeEnd w:id="{cid}"/>'
       f'<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="{cid}"/></w:r>'
       + run(after))
doc.write_text(xml[:m.start()] + new + xml[m.end():], encoding="utf-8")
# comment.py dates the comment now: write the requested date where it put the timestamp
comments = u / "word" / "comments.xml"
stamp = re.search(rf'<w:comment w:id="{cid}" [^>]*w:date="([^"]*)"', comments.read_text(encoding="utf-8")).group(1)
for name in ("comments.xml", "commentsExtensible.xml"):
    part = u / "word" / name
    part.write_text(part.read_text(encoding="utf-8").replace(stamp, "2026-10-30T09:00:00Z"), encoding="utf-8")
