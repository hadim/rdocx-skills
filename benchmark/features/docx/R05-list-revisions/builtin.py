# bench: {"tools": ["pandoc --track-changes=all -t json"], "processes": 2, "raw_xml": false, "notes": "pandoc --track-changes=all reports insertions and deletions (body, cells, notes) but not formatting changes (w:rPrChange): the format change is missing"}
"""Tracked changes from pandoc's AST: insertion and deletion spans carry the author and the date."""
import json
import subprocess
import sys

ast = json.loads(subprocess.run(["pandoc", "--track-changes=all", "-t", "json", sys.argv[1]],
                                capture_output=True, text=True, check=True).stdout)
rows = []


def text_of(node):
    if isinstance(node, list):
        return "".join(text_of(x) for x in node)
    if not isinstance(node, dict):
        return ""
    t, c = node.get("t"), node.get("c")
    if t == "Str":
        return c
    if t in ("Space", "SoftBreak"):
        return " "
    return text_of(c) if isinstance(c, list) else ""


def walk(node):
    if isinstance(node, list):
        for x in node:
            walk(x)
    elif isinstance(node, dict):
        if node.get("t") == "Span" and {"insertion", "deletion"} & set(node["c"][0][1]):
            attrs = dict(node["c"][0][2])
            kind = "insert" if "insertion" in node["c"][0][1] else "delete"
            rows.append({"type": kind, "author": attrs.get("author"), "date": attrs.get("date"),
                         "text": text_of(node["c"][1])})
        walk(node.get("c"))


walk(ast["blocks"])
open(sys.argv[2], "w").write(json.dumps({"revisions": rows}, indent=1))
