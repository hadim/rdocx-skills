"""P11 oracle: each slide's notes (the notes slide's body placeholder) read as expected; every notes slide is
related to its slide and back, to the notes master, and has the notes-slide content type."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "lib"))
import oracle  # noqa: E402

WANT = {1: "Welcome everyone.", 2: "Spend two minutes on the agenda.",
        3: "Updated: the survey closed on 30 September.\nSource: field team, 412 answers.",
        4: "Stress the regional split.", 5: "Walk through the stages.", 6: "Close with the three asks."}
NOTES_CT = "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"


def check(out, inputs):
    pkg = oracle.Pkg(out)
    errors = pkg.integrity()
    slides = pkg.slides()
    if len(slides) != 6:
        return errors + [f"{len(slides)} slides"]
    for n, slide in enumerate(slides, 1):
        notes = pkg.notes_of(slide)
        got = pkg.notes_text(slide)
        if got != WANT[n]:
            errors.append(f"slide {n}: notes {got!r}, expected {WANT[n]!r}")
        if notes is None:
            continue
        if pkg.rels_of_type(notes, "slide") != [slide]:
            errors.append(f"slide {n}: notes slide not related back to its slide")
        if not pkg.rels_of_type(notes, "notesMaster"):
            errors.append(f"slide {n}: notes slide without notes master")
        if pkg.content_type(notes) != NOTES_CT:
            errors.append(f"slide {n}: notes slide content type {pkg.content_type(notes)}")
    return errors


oracle.main(check)
