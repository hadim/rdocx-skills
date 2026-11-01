"""Headings from pandoc's AST, each found as a line of its own in LibreOffice's PDF text, page by page."""
import json
import sys

ast, pages_txt, out = sys.argv[1], sys.argv[2], sys.argv[3]
pages = [{line.strip() for line in page.splitlines()} for page in open(pages_txt).read().split("\f")]
text = lambda inl: "".join(x["c"] if x["t"] == "Str" else " " if x["t"] in ("Space", "SoftBreak") else ""  # noqa: E731
                           for x in inl)
rows, page = [], 1
for block in json.load(open(ast))["blocks"]:
    if block["t"] == "Header":
        level, _, inlines = block["c"]
        title = text(inlines)
        page = next(i + 1 for i in range(page - 1, len(pages)) if title in pages[i])
        rows.append({"level": level, "text": title, "page": page})
open(out, "w").write(json.dumps({"headings": rows}, indent=1))
