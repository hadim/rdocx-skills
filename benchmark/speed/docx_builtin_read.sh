#!/usr/bin/env bash
# Built-in docx skill, complete text through its own XML read (as features/docx D02 did): unzip, then read every
# story with lxml; prints JSON.
set -euo pipefail
W=$(mktemp -d)
unzip -q "$1" -d "$W/u"
find "$W/u" -type l -delete
python "$(dirname "$0")/docx_builtin_read.py" "$W/u"
