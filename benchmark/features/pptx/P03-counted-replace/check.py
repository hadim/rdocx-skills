"""P03 oracle: compares every paragraph of the slides and notes with the input. No "Project Falcon" left, 7
"Project Osprey", the decoys untouched, every other paragraph unchanged; character by character, text outside a
match keeps its formatting (bold, italic, underline, strike, size, colour) and the replacement takes the
formatting of one of the matched characters. (The refusal half is run by run.py: count 8, exit != 0, no output.)"""
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "lib"))
import oracle  # noqa: E402
from oracle import NS, q  # noqa: E402

OLD, NEW = "Project Falcon", "Project Osprey"


def fmt(r):
    rpr = r.find("a:rPr", NS)
    if rpr is None:
        return (False, False, False, None, None, None)
    color = rpr.find("a:solidFill/*", NS)
    return (rpr.get("b") == "1", rpr.get("i") == "1", rpr.get("u") not in (None, "none"), rpr.get("strike"),
            rpr.get("sz"), color.get("val") if color is not None else None)


def chars(p):
    out = []
    for child in p:
        if child.tag in (q("a:r"), q("a:fld")):
            f = fmt(child)
            out += [(c, f) for c in (child.findtext("a:t", default="", namespaces=NS))]
        elif child.tag == q("a:br"):
            out.append(("\n", None))
    return out


def paragraphs(pkg):
    out = []
    for s in pkg.slides():
        for part in (s, pkg.notes_of(s)):
            if part:
                tree = pkg.xml(part).find("p:cSld/p:spTree", NS)
                out += [(part, chars(p)) for p in tree.iter(q("a:p"))]
    return out


def check(out, inputs):
    before = paragraphs(oracle.Pkg(inputs / "deck.pptx"))
    pkg = oracle.Pkg(out)
    errors = pkg.integrity()
    after = paragraphs(pkg)
    alltext = "\n".join("".join(c for c, _ in p) for _, p in after)
    if alltext.count(OLD) or alltext.count(NEW) != 7:
        errors.append(f"{alltext.count(OLD)} '{OLD}' left, {alltext.count(NEW)} '{NEW}' (expected 0 and 7)")
    for decoy in ("project falcon team", "PROJECT FALCON is"):
        if decoy not in alltext:
            errors.append(f"decoy {decoy!r} changed")
    if len(before) != len(after):
        return errors + [f"{len(after)} paragraphs, expected {len(before)}"]
    kept = 0
    for (part, b), (_, a) in zip(before, after):
        btext = "".join(c for c, _ in b)
        want = btext.replace(OLD, NEW)
        atext = "".join(c for c, _ in a)
        if atext != want:
            errors.append(f"{part}: {atext!r}, expected {want!r}")
            continue
        if OLD not in btext:
            if [f for _, f in a] != [f for _, f in b]:
                errors.append(f"{part}: formatting of {btext!r} changed")
            continue
        i = j = 0
        lost = []
        for m in re.finditer(re.escape(OLD), btext):
            while i < m.start():
                if a[j][1] != b[i][1]:
                    lost.append(i)
                i, j = i + 1, j + 1
            allowed = {f for _, f in b[m.start():m.end()]}
            for k in range(len(NEW)):
                if a[j + k][1] not in allowed:
                    errors.append(f"{part}: {btext!r}: replacement char {k} has formatting {a[j + k][1]}, "
                                  f"not one of the matched runs'")
                    break
            i, j = m.end(), j + len(NEW)
        while i < len(b):
            if a[j][1] != b[i][1]:
                lost.append(i)
            i, j = i + 1, j + 1
        if lost:
            errors.append(f"{part}: {btext!r}: {len(lost)} characters outside the match lost their formatting")
        kept += 1
    return errors, {"paragraphs_with_matches": kept}


oracle.main(check)
