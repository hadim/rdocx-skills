#!/usr/bin/env bash
# bench: {"tools": ["unzip", "merge_runs.py", "comment.py", "python (XML edit)", "zip"], "processes": 5, "raw_xml": true, "files": ["builtin.sh", "builtin_edit.py"]}
set -euo pipefail
S="$BUILTIN_DOCX/scripts" W=$(mktemp -d)
unzip -q "$1" -d "$W/u"
find "$W/u" -type l -delete
python "$S/merge_runs.py" "$W/u" >/dev/null
cid=$(python "$S/comment.py" "$W/u" "Should this be 45 days?" --author "Dana Reviewer" --initials DR \
  | sed -n 's/^Added comment id=\([0-9]*\).*/\1/p')
python "$(dirname "$0")/builtin_edit.py" "$W/u" "$cid"
(cd "$W/u" && zip -qXr "$W/out.docx" .)
mv "$W/out.docx" "$2"
