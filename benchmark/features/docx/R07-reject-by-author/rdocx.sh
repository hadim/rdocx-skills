#!/usr/bin/env bash
# bench: {"tools": ["rdocx revision reject --author"], "processes": 1, "raw_xml": false}
set -euo pipefail
rdocx revision reject "$1" --author "Dana Reviewer" -o "$2" >/dev/null
