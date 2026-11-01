"""P02 oracle: the JSON names 8 slides in order, each title exactly, each slide's marker strings in that slide's
text only, the notes markers in the notes, and no notes on slide 7."""
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "lib"))
import oracle  # noqa: E402


def check(out, inputs):
    truth = json.loads((inputs / "truth.json").read_text())
    data = json.loads(out.read_text())
    slides = data["slides"]
    errors = []
    if len(slides) != 8:
        return [f"{len(slides)} slides, expected 8"]
    for k, s in enumerate(slides):
        n = k + 1
        if s.get("number") != n:
            errors.append(f"slide {n}: number {s.get('number')}")
        if (s.get("title") or "").strip() != truth["titles"][k]:
            errors.append(f"slide {n}: title {s.get('title')!r}")
        text = " ".join(s.get("text") or [])
        for m in truth["markers"][str(n)]:
            if m not in text:
                errors.append(f"slide {n}: marker {m} missing from text")
        for other, marks in truth["markers"].items():
            if int(other) != n:
                for m in marks:
                    if m in text and not any(m in mine and m != mine for mine in truth["markers"][str(n)]):
                        errors.append(f"slide {n}: marker {m} of slide {other} found here")
        notes = s.get("notes") or ""
        want = truth["notes"].get(str(n))
        if want and notes.strip() != want:
            errors.append(f"slide {n}: notes {notes!r}")
        if not want and notes.strip():
            errors.append(f"slide {n}: notes {notes!r}, expected none")
    return errors


oracle.main(check, pptx_output=False)
