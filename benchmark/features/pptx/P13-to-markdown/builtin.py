# bench: tools=slide_text.py; processes=2; raw_xml=no
import os
import re
import subprocess
import sys

SKILL = os.environ["PPTX_SKILL"]  # the built-in pptx skill's folder
md = subprocess.run([sys.executable, f"{SKILL}/scripts/slide_text.py", sys.argv[1]],
                    capture_output=True, text=True, check=True).stdout
out, in_notes = [], False
for line in md.splitlines():
    if line.startswith("<!-- Slide number"):
        in_notes = False
        out.append("")
    elif line.startswith("### Notes:"):
        in_notes = True
        out += ["", "**Notes:**", ""]
    elif in_notes or not line.strip() or re.match(r"^(#|\||!\[)", line):
        out.append(line)
    else:
        out.append("- " + line)
with open(sys.argv[2], "w") as f:
    f.write(re.sub(r"\n{3,}", "\n\n", "\n".join(out)).strip() + "\n")
