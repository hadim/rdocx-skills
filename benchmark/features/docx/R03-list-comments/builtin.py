# bench: {"tools": ["pandoc --track-changes=all -t json"], "processes": 2, "raw_xml": false, "notes": "pandoc's comment-start / comment-end spans carry id, author, date and text; the anchor is rebuilt by walking the AST (tables included)"}
"""Comments from pandoc's AST: each comment-start span carries id, author, date and the comment text; the
anchored text is what lies between comment-start and comment-end."""
import json
import subprocess
import sys

ast = json.loads(subprocess.run(["pandoc", "--track-changes=all", "-t", "json", sys.argv[1]],
                                capture_output=True, text=True, check=True).stdout)
comments, open_ids = {}, []


def text_of(inlines):
    out = []
    for x in inlines:
        t, c = x["t"], x.get("c")
        if t == "Str":
            out.append(c)
        elif t in ("Space", "SoftBreak"):
            out.append(" ")
        elif t == "LineBreak":
            out.append("\n")
        elif t in ("Emph", "Strong", "Underline", "SmallCaps", "Strikeout", "Superscript", "Subscript"):
            out.append(text_of(c))
        elif t in ("Span", "Link", "Quoted"):
            out.append(text_of(c[1] if t != "Quoted" else c[1]))
    return "".join(out)


def walk_inlines(inlines):
    for x in inlines:
        if x["t"] == "Span" and "comment-start" in x["c"][0][1]:
            attrs = dict(x["c"][0][2])
            cid = int(attrs["comment-id"])
            comments[cid] = {"id": cid, "author": attrs.get("author"), "date": attrs.get("date"),
                             "text": text_of(x["c"][1]), "anchor": ""}
            open_ids.append(cid)
        elif x["t"] == "Span" and "comment-end" in x["c"][0][1]:
            walk_inlines(x["c"][1])
            open_ids.remove(int(dict(x["c"][0][2])["comment-id"]))
        elif x["t"] in ("Span", "Emph", "Strong", "Underline", "SmallCaps", "Strikeout"):
            walk_inlines(x["c"][1] if x["t"] == "Span" else x["c"])
        else:
            piece = text_of([x])
            for cid in open_ids:
                comments[cid]["anchor"] += piece


def walk_blocks(blocks):
    for blk in blocks:
        t, c = blk["t"], blk.get("c")
        if t in ("Para", "Plain", "Header"):
            for cid in open_ids:
                if comments[cid]["anchor"]:
                    comments[cid]["anchor"] += "\n"
            walk_inlines(c[2] if t == "Header" else c)
        elif t == "Table":  # [attr, caption, colspecs, head, bodies, foot]
            head, bodies, foot = c[3], c[4], c[5]
            for row in head[1] + [r for body in bodies for r in body[2] + body[3]] + foot[1]:
                for cell in row[1]:
                    walk_blocks(cell[4])
        elif t in ("BlockQuote", "Div"):
            walk_blocks(c if t == "BlockQuote" else c[1])
        elif t in ("BulletList", "OrderedList"):
            for item in (c if t == "BulletList" else c[1]):
                walk_blocks(item)


walk_blocks(ast["blocks"])
open(sys.argv[2], "w").write(json.dumps({"comments": [comments[k] for k in sorted(comments)]}, indent=1))
