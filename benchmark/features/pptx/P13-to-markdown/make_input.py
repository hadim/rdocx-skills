"""P13 input: the standard 8-slide review deck (title, bullets with levels, two columns, table, group, chart,
picture placeholder, section header; notes on 7 of 8 slides), plus truth.json (titles, markers, notes)."""
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "lib"))
import inputs  # noqa: E402

out = Path(sys.argv[1])
out.mkdir(parents=True, exist_ok=True)
inputs.save(inputs.review_deck(inputs.png(800, 600, 1)), out / "deck.pptx")
(out / "truth.json").write_text(json.dumps(inputs.REVIEW, indent=1))
