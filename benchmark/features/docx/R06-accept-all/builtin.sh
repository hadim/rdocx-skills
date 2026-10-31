#!/usr/bin/env bash
# bench: {"tools": ["accept_changes.py (LibreOffice)"], "processes": 2, "raw_xml": false, "libreoffice": true, "lo_profile": true, "notes": "the case the skill warns about (a deleted paragraph before an empty spacer paragraph, also as the last item of a bulleted list) came out right with LibreOffice 26.2"}
set -euo pipefail
python "$BUILTIN_DOCX/scripts/accept_changes.py" "$1" "$2" >/dev/null
