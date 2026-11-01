#!/usr/bin/env bash
set -euo pipefail
W=$(mktemp -d)
unzip -q "$1" -d "$W/unpacked"
find "$W/unpacked" -type l -delete
python "$(dirname "$0")/builtin_pictures.py" "$W/unpacked" "$2" "$3"
(cd "$W/unpacked" && zip -qXr "$4" .)
