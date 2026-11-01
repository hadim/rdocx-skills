#!/usr/bin/env bash
set -euo pipefail
node "$(dirname "$0")/builtin.js" "$1" "$2"
