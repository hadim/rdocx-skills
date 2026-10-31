#!/usr/bin/env bash
# bench: {"tools": ["rdocx toc rebuild"], "processes": 1, "raw_xml": false}
set -euo pipefail
rdocx toc rebuild "$1" -o "$2" >/dev/null
