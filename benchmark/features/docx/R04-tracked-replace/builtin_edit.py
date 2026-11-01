"""Wrap the old word in w:del (w:delText) and the new one in w:ins, in the run that holds it."""
import re
import sys
from pathlib import Path

doc = Path(sys.argv[1])
xml = doc.read_text(encoding="utf-8")
attrs = 'w:author="Dana Reviewer" w:date="2026-10-30T11:00:00Z"'
m = re.search(r"<w:r>(<w:rPr>(?:(?!</w:r>).)*?</w:rPr>)?<w:t[^>]*>([^<]*\bthirty\b[^<]*)</w:t></w:r>", xml)
rpr, (before, after) = m.group(1) or "", re.split(r"\bthirty\b", m.group(2), maxsplit=1)
run = lambda s: f'<w:r>{rpr}<w:t xml:space="preserve">{s}</w:t></w:r>'  # noqa: E731
new = (run(before) + f'<w:del w:id="9001" {attrs}><w:r>{rpr}<w:delText>thirty</w:delText></w:r></w:del>'
       + f'<w:ins w:id="9002" {attrs}>{run("forty-five")}</w:ins>' + run(after))
doc.write_text(xml[:m.start()] + new + xml[m.end():], encoding="utf-8")
