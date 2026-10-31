#!/usr/bin/env bash
# Built-in pptx skill, editing workflow: unpack with zipfile (as the skill shows), edit the slide and notes XML,
# zip from inside the folder; with --validate, then validate.py --original.
#   pptx_builtin_replace.sh IN OUT OLD NEW COUNT [--validate]
set -euo pipefail
S="$BUILTIN_PPTX/scripts" W=$(mktemp -d)
python -c "import sys,zipfile; zipfile.ZipFile(sys.argv[1]).extractall(sys.argv[2])" "$1" "$W/u"
python "$(dirname "$0")/pptx_builtin_replace.py" "$W/u" "$3" "$4" "$5"
(cd "$W/u" && zip -qXr "$W/out.pptx" .)
mv "$W/out.pptx" "$2"
if [ "${6:-}" = "--validate" ]; then
  python "$S/office/validate.py" "$2" --original "$1" >/dev/null
fi
