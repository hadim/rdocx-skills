#!/usr/bin/env bash
# bench: {"tools": ["rdocx revision accept"], "processes": 1, "raw_xml": false}
set -euo pipefail
rdocx revision accept "$1" -o "$2" >/dev/null
