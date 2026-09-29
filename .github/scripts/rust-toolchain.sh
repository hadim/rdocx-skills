#!/bin/sh
# Install the Rust toolchain that rdocx's rust-toolchain.toml names at commit $1 of the upstream the lock names,
# with its components and targets, so that no rustup auto-install happens inside the build. rdocx_env.py never
# installs a toolchain.
set -eu
if ! command -v rustup > /dev/null; then  # not every runner image ships rustup
  curl -fsSL https://sh.rustup.rs | sh -s -- -y --profile minimal --default-toolchain none
  PATH="$HOME/.cargo/bin:$PATH"
  echo "$HOME/.cargo/bin" >> "${GITHUB_PATH:-/dev/null}"
fi
commit="$1"
file="${RUNNER_TEMP:-/tmp}/rust-toolchain.toml"
repo=$(python -c 'import json; print(json.load(open("rdocx.lock.json"))["upstream"].removeprefix("https://github.com/").removesuffix(".git"))')
curl -fsSL "https://raw.githubusercontent.com/${repo}/${commit}/rust-toolchain.toml" -o "$file"
args=$(python - "$file" <<'PY'
import sys, tomllib
t = tomllib.load(open(sys.argv[1], "rb"))["toolchain"]
print(" ".join([t["channel"]] + [f"-c {c}" for c in t.get("components", [])] + [f"-t {x}" for x in t.get("targets", [])]))
PY
)
# shellcheck disable=SC2086
rustup toolchain install --profile minimal $args
rustup show
