#!/usr/bin/env bash
set -euo pipefail
rdocx convert "$1" --to md -o "$2"
