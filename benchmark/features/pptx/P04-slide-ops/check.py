"""P04 oracle: the final order by title marker is Alpha, Hotel, Bravo, Charlie, Echo, Charlie (the copy), Foxtrot,
Golf; only Foxtrot is hidden; the copy is its own slide part with Charlie's picture and its own notes slide holding
Charlie's notes (related back to the copy); Delta's slide, notes and picture are gone from the package; the
presentation's slide relationships and <p:sldIdLst> agree; relationships, content types and reachability are
clean (no orphan part)."""
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "lib"))
import oracle  # noqa: E402
from oracle import NS  # noqa: E402


def check(out, inputs):
    truth = json.loads((inputs / "truth.json").read_text())
    pkg = oracle.Pkg(out)
    errors = pkg.integrity()
    ids = pkg.sld_ids()
    order = [(pkg.title(p) or "?").split(":")[0] for _, _, p, _ in ids]
    if order != truth["order"]:
        errors.append(f"order {order}, expected {truth['order']}")
    hidden = [(pkg.title(p) or "?").split(":")[0] for _, _, p, h in ids if h]
    if hidden != truth["hidden"]:
        errors.append(f"hidden {hidden}, expected {truth['hidden']}")
    parts = [p for _, _, p, _ in ids]
    if len(set(parts)) != len(parts) or len({i for i, _, _, _ in ids}) != len(ids):
        errors.append("two sldId entries share a slide part or an id")
    related = set(pkg.rels_of_type(pkg.main, "slide"))
    if related != set(parts):
        errors.append(f"slides related but not listed or listed but not related: {sorted(related ^ set(parts))}")
    charlie = (inputs / "charlie.png").read_bytes()
    delta = (inputs / "delta.png").read_bytes()
    if len(parts) >= 6 and order[5] == "Charlie":
        copy, orig = parts[5], parts[3]
        imgs = [pkg.read(t) for t in pkg.rels_of_type(copy, "image")]
        if charlie not in imgs:
            errors.append("the copy does not show Charlie's picture")
        notes = pkg.notes_of(copy)
        if notes is None or notes == pkg.notes_of(orig):
            errors.append("the copy has no notes slide of its own")
        elif "notes-charlie" not in (pkg.notes_text(copy) or ""):
            errors.append(f"the copy's notes read {pkg.notes_text(copy)!r}")
        elif pkg.rels_of_type(notes, "slide") != [copy]:
            errors.append("the copy's notes slide is not related back to the copy")
    for n in pkg.names:
        data = pkg.read(n)
        if data == delta:
            errors.append(f"Delta's picture still in the package ({n})")
        if n.endswith(".xml") and (b"notes-delta" in data or b">Delta: section" in data):
            errors.append(f"Delta's slide or notes still in the package ({n})")
    notes_owner = {}
    for p in parts:
        nt = pkg.notes_of(p)
        if nt:
            notes_owner.setdefault(nt, []).append(p)
    if any(len(v) > 1 for v in notes_owner.values()):
        errors.append("two slides share a notes slide")
    orphans = pkg.orphans()
    if orphans:
        errors.append(f"orphan parts: {orphans}")
    return errors


oracle.main(check)
