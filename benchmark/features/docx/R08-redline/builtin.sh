#!/usr/bin/env bash
# bench: {"tools": ["unzip", "python + lxml + difflib (redline by hand)", "zip"], "processes": 4, "raw_xml": true, "files": ["builtin.sh", "builtin_edit.py"]}
set -euo pipefail
W=$(mktemp -d)
unzip -q "$1" -d "$W/v1"
unzip -q "$2" -d "$W/v2"
find "$W" -type l -delete
python "$(dirname "$0")/builtin_edit.py" "$W/v1/word/document.xml" "$W/v2/word/document.xml"
(cd "$W/v1" && zip -qXr "$W/out.docx" .)
mv "$W/out.docx" "$3"
