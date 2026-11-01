# bench: {"tools": ["rdocx diff --json"], "processes": 2, "raw_xml": false}
import json
import subprocess
import sys

res = json.loads(subprocess.run(["rdocx", "diff", "--json", sys.argv[1], sys.argv[2]], capture_output=True,
                                text=True, check=True).stdout)
out = {"changed": [], "added": [], "removed": []}
for d in res["differences"]:
    out[d["change"]].append([d["text_a"], d["text_b"]] if d["change"] == "changed" else d["text_a"] or d["text_b"])
open(sys.argv[3], "w").write(json.dumps(out, indent=1))
