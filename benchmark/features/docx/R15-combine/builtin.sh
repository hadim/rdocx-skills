#!/usr/bin/env bash
# bench: {"tools": ["unzip", "python + lxml (merge by hand)", "zip"], "processes": 4, "raw_xml": true, "files": ["builtin.sh", "builtin_merge.py"]}
set -euo pipefail
W=$(mktemp -d)
unzip -q "$1" -d "$W/a"
unzip -q "$2" -d "$W/b"
find "$W" -type l -delete
python "$(dirname "$0")/builtin_merge.py" "$W/a" "$W/b"
(cd "$W/a" && zip -qXr "$W/out.docx" .)
mv "$W/out.docx" "$3"
