#!/usr/bin/env bash
# Built-in docx skill, editing workflow: unzip, strip symlinks, merge_runs.py, edit document.xml (and the headers
# and footers) in place, zip; with --validate, then validate.py --original as the skill's workflow says.
#   docx_builtin_replace.sh IN OUT OLD NEW COUNT [--validate]
set -euo pipefail
S="$BUILTIN_DOCX/scripts" W=$(mktemp -d)
unzip -q "$1" -d "$W/u"
find "$W/u" -type l -delete
python "$S/merge_runs.py" "$W/u" >/dev/null
python "$(dirname "$0")/docx_builtin_replace.py" "$W/u" "$3" "$4" "$5"
(cd "$W/u" && zip -qXr "$W/out.docx" .)
mv "$W/out.docx" "$2"
if [ "${6:-}" = "--validate" ]; then
  python "$S/office/validate.py" "$2" --original "$1" >/dev/null
fi
