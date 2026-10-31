"""P07 oracle: slide 3 (and no other) has a comment part; a modern comment (PowerPoint 365 threads) whose author is
"Dana Reviewer", text "Please check this figure.", anchored by ac:spMk to the id (and creationId, when given) of
the shape "Key figure" on that slide (sldMk naming the slide's id; the slide's p188:commentRel extension naming
the relationship, as PowerPoint requires), with one reply by "Sam Author" reading
"Updated with Q3 numbers."; parts carry the comment and author content types. A legacy comment (p:cmLst) cannot
anchor to a shape and is reported as such."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "lib"))
import oracle  # noqa: E402
from oracle import NS  # noqa: E402

P188 = "http://schemas.microsoft.com/office/powerpoint/2018/8/main"
AC = "http://schemas.microsoft.com/office/drawing/2013/main/command"
PC = "http://schemas.microsoft.com/office/powerpoint/2013/main/command"
A16 = "http://schemas.microsoft.com/office/drawing/2014/main"
M = {"p188": P188, "ac": AC, "pc": PC, "a": NS["a"], "a16": A16}


def check(out, inputs):
    pkg = oracle.Pkg(out)
    errors = pkg.integrity()
    ids = pkg.sld_ids()
    for n, (sid, _, slide, _) in enumerate(ids, 1):
        modern = pkg.rels_of_type(slide, "comments")
        rels = [r for r in pkg.rels(slide) if r["type"].endswith("/comments")]
        if n != 3:
            if rels:
                errors.append(f"slide {n} has comments")
            continue
        if not rels:
            return errors + ["slide 3 has no comment part"]
        rel = rels[0]
        if not rel["type"].startswith("http://schemas.microsoft.com/office/2018/10/"):
            return errors + ["legacy comment (p:cmLst): it has a position, no anchor to a shape, and no reply thread"]
        part = rel["target"]
        crel = pkg.xml(slide).find(f"p:extLst/p:ext/{{{P188}}}commentRel", NS)
        if crel is None or crel.get(f"{{{NS['r']}}}id") != rel["id"]:
            errors.append("slide 3 lacks the p188:commentRel extension naming its comment relationship")
        if pkg.content_type(part) != "application/vnd.ms-powerpoint.comments+xml":
            errors.append(f"comment part content type {pkg.content_type(part)}")
        authors_part = (pkg.rels_of_type(pkg.main, "authors") or [None])[0]
        if authors_part is None:
            return errors + ["no authors part related to the presentation"]
        if pkg.content_type(authors_part) != "application/vnd.ms-powerpoint.authors+xml":
            errors.append(f"authors part content type {pkg.content_type(authors_part)}")
        authors = {a.get("id"): a.get("name") for a in pkg.xml(authors_part).iter(f"{{{P188}}}author")}
        cms = pkg.xml(part).findall("p188:cm", M)
        if len(cms) != 1:
            return errors + [f"{len(cms)} comments on slide 3, expected 1"]
        cm = cms[0]
        if authors.get(cm.get("authorId")) != "Dana Reviewer":
            errors.append(f"comment author {authors.get(cm.get('authorId'))!r}")
        if "".join(cm.find("p188:txBody", M).itertext()).strip() != "Please check this figure.":
            errors.append("comment text differs")
        key = oracle.find_one(pkg, slide, lambda sh: pkg.shape_name(sh) == "Key figure")
        mk = cm.find("ac:deMkLst", M)
        sp = mk.find("ac:spMk", M) if mk is not None else None
        sld = mk.find("pc:sldMk", M) if mk is not None else None
        if sp is None:
            errors.append("comment not anchored to a shape (no ac:spMk)")
        else:
            if sp.get("id") != pkg.shape_id(key):
                errors.append(f"anchored to shape id {sp.get('id')}, 'Key figure' is {pkg.shape_id(key)}")
            cid = key.find(".//a16:creationId", M)
            if sp.get("creationId") and (cid is None or cid.get("id") != sp.get("creationId")):
                errors.append("spMk creationId does not match the shape's a16:creationId")
        if sld is None or sld.get("sldId") != sid:
            errors.append(f"sldMk names slide id {None if sld is None else sld.get('sldId')}, slide 3 is {sid}")
        replies = cm.findall("p188:replyLst/p188:reply", M)
        if len(replies) != 1:
            errors.append(f"{len(replies)} replies, expected 1")
        else:
            r = replies[0]
            if authors.get(r.get("authorId")) != "Sam Author":
                errors.append(f"reply author {authors.get(r.get('authorId'))!r}")
            if "".join(r.find("p188:txBody", M).itertext()).strip() != "Updated with Q3 numbers.":
                errors.append("reply text differs")
    return errors


oracle.main(check)
