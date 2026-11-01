"""R14 oracle: the three changed pairs, the added and the removed paragraph are reported (a reported entry may
carry more than the paragraph, such as a whole table row), and nothing else."""
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
sys.path.insert(0, str(Path(__file__).resolve().parent))
import review_oracle as o  # noqa: E402
from make_input import EXPECTED  # noqa: E402


def main(out):
    got = json.loads(Path(out).read_text())
    checks = {}
    for a, b in EXPECTED["changed"]:
        checks[f"changed:{a[:30]}"] = any(a in (ga or "") and b in (gb or "") for ga, gb in got.get("changed", []))
    for kind in ("added", "removed"):
        for t in EXPECTED[kind]:
            checks[f"{kind}:{t[:30]}"] = any(t in (g or "") for g in got.get(kind, []))
    total = sum(len(got.get(k, [])) for k in ("changed", "added", "removed"))
    checks["nothing_else"] = total == sum(len(v) for v in EXPECTED.values())
    o.report(checks, {"got": got})


o.guard(lambda: main(sys.argv[1]))
