"""D01 input: the PNG picture the new report embeds (400 x 300 px)."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from authoring_inputs import png_bytes  # noqa: E402

out = Path(sys.argv[1])
(out / "picture.png").write_bytes(png_bytes(400, 300, (31, 78, 121)))
