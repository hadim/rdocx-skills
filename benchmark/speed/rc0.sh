#!/usr/bin/env bash
# Run a checker whose exit code is its answer (validate, fit: 1 means "invalid" or "overflows", not an error):
# its output goes to OUT, its exit code to OUT.rc, and this script exits 0 unless the checker could not run.
#   rc0.sh OUT command args...
out=$1
shift
"$@" >"$out" 2>&1
rc=$?
echo "$rc" >"$out.rc"
[ "$rc" -le 1 ]
