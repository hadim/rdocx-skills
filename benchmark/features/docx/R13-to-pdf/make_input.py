"""R13 input: the four-page report of R11 (header, footer page number, tracked changes on pages 1 and 2, a
picture, a table, a footnote)."""
import importlib.util
import sys
from pathlib import Path

spec = importlib.util.spec_from_file_location("r11", Path(__file__).resolve().parents[1] / "R11-render-png" / "make_input.py")
r11 = importlib.util.module_from_spec(spec)
spec.loader.exec_module(r11)
PAGES = r11.PAGES

if __name__ == "__main__":
    print(r11.build(sys.argv[1]))
