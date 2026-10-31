#!/usr/bin/env bash
# bench: {"tools": ["unzip", "python + lxml (XML edit)", "zip"], "processes": 3, "raw_xml": true, "files": ["builtin.sh", "builtin_edit.py"]}
set -euo pipefail
W=$(mktemp -d)
unzip -q "$1" -d "$W/u"
find "$W/u" -type l -delete
python "$(dirname "$0")/builtin_edit.py" "$W/u/word/document.xml" "Dana Reviewer"
(cd "$W/u" && zip -qXr "$W/out.docx" .)
mv "$W/out.docx" "$2"
