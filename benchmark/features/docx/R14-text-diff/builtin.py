# bench: {"tools": ["pandoc -t plain", "difflib"], "processes": 3, "raw_xml": false, "notes": "pandoc's plain text puts a table row on one line, so the changed cell is reported with its row"}
import difflib
import json
import subprocess
import sys


def lines(path):
    text = subprocess.run(["pandoc", "-t", "plain", "--wrap=none", path], capture_output=True, text=True,
                          check=True).stdout
    return [line.strip() for line in text.splitlines() if line.strip()]


a, b = lines(sys.argv[1]), lines(sys.argv[2])
out = {"changed": [], "added": [], "removed": []}
for op, i1, i2, j1, j2 in difflib.SequenceMatcher(None, a, b, autojunk=False).get_opcodes():
    if op == "replace":
        out["changed"] += [list(pair) for pair in zip(a[i1:i2], b[j1:j2])]
        out["removed"] += a[i1 + (j2 - j1):i2]
        out["added"] += b[j1 + (i2 - i1):j2]
    elif op == "delete":
        out["removed"] += a[i1:i2]
    elif op == "insert":
        out["added"] += b[j1:j2]
open(sys.argv[3], "w").write(json.dumps(out, indent=1))
