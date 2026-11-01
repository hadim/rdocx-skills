#!/usr/bin/env bash
set -euo pipefail
pandoc "$1" -o "$2"
