"""P06 oracle: reads the reported frames, matches them to the known truth by (slide, shape name) and scores
precision and recall; pass only when both are 1."""
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "lib"))
import oracle  # noqa: E402


def check(out, inputs):
    truth = json.loads((inputs / "truth.json").read_text())
    want = {(f["slide"], f["shape"]) for f in truth["overflowing"]}
    got = {(int(f["slide"]), f["shape"]) for f in json.loads(out.read_text())["overflowing"]}
    hit = got & want
    precision = len(hit) / len(got) if got else 0.0
    recall = len(hit) / len(want)
    errors = []
    if got - want:
        errors.append(f"false alarms: {sorted(got - want)}")
    if want - got:
        errors.append(f"missed: {sorted(want - got)}")
    return errors, {"precision": round(precision, 3), "recall": round(recall, 3), "reported": sorted(got)}


oracle.main(check, pptx_output=False)
