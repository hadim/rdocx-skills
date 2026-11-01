#!/usr/bin/env bash
set -euo pipefail
pandoc -t markdown "$1" -o "$2"
