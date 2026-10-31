# bench: tools=slide_text.py; processes=2; raw_xml=no
import json
import os
import re
import subprocess
import sys

SKILL = os.environ["PPTX_SKILL"]  # the built-in pptx skill's folder
md = subprocess.run([sys.executable, f"{SKILL}/scripts/slide_text.py", sys.argv[1]],
                    capture_output=True, text=True, check=True).stdout
slides = []
parts = re.split(r"<!-- Slide number: (\d+) -->", md)
for number, block in zip(parts[1::2], parts[2::2]):
    body, _, notes = block.partition("### Notes:")
    title, text = None, []
    for line in body.strip().splitlines():
        if line.startswith("# ") and title is None:
            title = line[2:].strip()
        elif line.startswith("|") and not re.match(r"^\|( --- \|)+$", line):
            text += [c.strip() for c in line.strip("|").split(" | ")]
        elif line.strip() and not line.startswith("|"):
            text.append(line)
    slides.append({"number": int(number), "title": title, "text": text, "notes": notes.strip() or None})
with open(sys.argv[2], "w") as f:
    json.dump({"slides": slides}, f, indent=1)
