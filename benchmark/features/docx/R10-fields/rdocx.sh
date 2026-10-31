#!/usr/bin/env bash
# bench: {"tools": ["rdocx fields update --now"], "processes": 1, "raw_xml": false}
set -euo pipefail
rdocx fields update "$1" -o "$2" --now "$(date +%Y-%m-%dT%H:%M:%S)" >/dev/null
