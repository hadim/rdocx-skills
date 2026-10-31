#!/usr/bin/env bash
# bench: {"tools": ["unzip", "comment.py", "python (XML edit)", "zip"], "processes": 4, "raw_xml": true, "files": ["builtin.sh", "builtin_edit.py"]}
set -euo pipefail
S="$BUILTIN_DOCX/scripts" W=$(mktemp -d)
unzip -q "$1" -d "$W/u"
find "$W/u" -type l -delete
parent=$(grep -o '<w:comment [^>]*>' "$W/u/word/comments.xml" | grep 'w:author="Alex Author"' \
  | sed -n 's/.*w:id="\([0-9]*\)".*/\1/p')
cid=$(python "$S/comment.py" "$W/u" "Agreed, changed to 45 days." --author "Dana Reviewer" --initials DR \
  --parent "$parent" | sed -n 's/^Added reply id=\([0-9]*\).*/\1/p')
python "$(dirname "$0")/builtin_edit.py" "$W/u" "$parent" "$cid"
(cd "$W/u" && zip -qXr "$W/out.docx" .)
mv "$W/out.docx" "$2"
