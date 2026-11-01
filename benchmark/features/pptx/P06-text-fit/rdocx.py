# bench: tools=rpptx (Python) fit_report; processes=1; raw_xml=no
import json
import os
import sys

import rpptx

report = rpptx.Presentation(sys.argv[1]).fit_report(font_dir=os.environ["FONT_DIR"])
found = {(f.update.slide_index + 1, f.update.name) for f in report.frames
         if f.overflow and f.update.owner_kind == "slide"}  # table cells come under their table's name
with open(sys.argv[2], "w") as fh:
    json.dump({"overflowing": [{"slide": s, "shape": n} for s, n in sorted(found)]}, fh, indent=1)
