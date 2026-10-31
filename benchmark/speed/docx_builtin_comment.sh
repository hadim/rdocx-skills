#!/usr/bin/env bash
# Built-in docx skill: a comment anchored on a phrase. unzip, merge_runs.py, comment.py (directory mode, as the
# skill prefers when markers are placed too), place the markers it prints around the phrase, zip; with --validate,
# then validate.py --original.
#   docx_builtin_comment.sh IN OUT PHRASE [--validate]
set -euo pipefail
S="$BUILTIN_DOCX/scripts" W=$(mktemp -d)
unzip -q "$1" -d "$W/u"
find "$W/u" -type l -delete
python "$S/merge_runs.py" "$W/u" >/dev/null
cid=$(python "$S/comment.py" "$W/u" "Speed benchmark comment" --author "Speed Bench" --initials SB \
  | sed -n 's/^Added comment id=\([0-9]*\).*/\1/p')
python "$(dirname "$0")/docx_builtin_mark.py" "$W/u" "$cid" "$3"
(cd "$W/u" && zip -qXr "$W/out.docx" .)
mv "$W/out.docx" "$2"
if [ "${4:-}" = "--validate" ]; then
  python "$S/office/validate.py" "$2" --original "$1" >/dev/null
fi
