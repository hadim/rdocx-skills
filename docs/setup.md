# Installing the pinned rdocx build

The skills never use whatever rdocx happens to be installed. They use the build named in `rdocx.lock.json`:

- `upstream` and `commit`: the rdocx source, by full 40-character git hash;
- `artifacts.<platform>`: the SHA-256 of the two CLIs and the two wheels built from that commit, per
  platform (`linux-x86_64`, `linux-aarch64`, `macos-arm64`);
- `installed_files.<platform>`: the SHA-256 of every file the two wheels install (modules, stubs, native
  libraries), checked again by `status` on the installed copy.

`scripts/rdocx_env.py` (standard library only, Python 3.9 or later, Linux and macOS; Windows is not
supported) enforces it.

## Commands

```bash
python3 scripts/rdocx_env.py status [--allow-local]  # exit 0 when the install matches the lock (or is an accepted local build)
python3 scripts/rdocx_env.py install                 # from a dist folder whose files match the lock; exit 2 if there is none
python3 scripts/rdocx_env.py install --build         # compile the pinned commit when no verified dist exists
python3 scripts/rdocx_env.py install --from DIR      # from a folder of prebuilt files, checked against the lock
python3 scripts/rdocx_env.py build [--target linux-aarch64]
python3 scripts/rdocx_env.py lock [--write] [--platform P]
python3 scripts/rdocx_env.py test [pytest args]      # the acceptance suite on the installed build
python3 scripts/rdocx_env.py paths
```

The install goes to `~/.local/share/rdocx-skills/<first 12 characters of the commit>/` (override the root
with `RDOCX_HOME`), with `current/bin/{rdocx,rpptx,python}`: two CLIs and a Python wrapper whose
environment has both modules. Nothing is installed system-wide. `install` from a dist folder uses no
network; `build` fetches the source and the pinned build tools, `test` the pinned test dependencies.

## Trust model

- **Prebuilt files**: each artifact is copied once into a private temporary folder, hashed there, compared
  with the lock, and installed from that copy only, so a file changed after the check is never used. The
  wheels are installed with `pip --no-index --no-deps --no-compile`. `status` re-hashes the CLIs, the
  `bin/python` wrapper and every file of the two packages against the lock, refuses any other file in them
  (compiled caches included: the wrapper sets `PYTHONDONTWRITEBYTECODE`), checks the interpreter link,
  `pyvenv.cfg` and the start-up files of site-packages (`.pth`, `sitecustomize`) against their state at
  install, and checks that `current` points at the pinned install; any mismatch exits 1, and `install`
  repairs it. Someone who can rewrite both the install and its `installed.json` can still defeat these
  checks: they detect corruption and stray changes, not an attacker who owns the home folder.
- **Test environment**: `test` installs the hash-pinned test dependencies into `$RDOCX_HOME/testenv`, which
  sees the installed packages through a `.pth` file, and runs the recipe blocks with the runtime
  `bin/python`: running the suite adds nothing to what the skills use.
- **Source**: the pinned commit is read object by object from git (`cat-file --batch`) with replace refs
  disabled, and the SHA-1 of every commit, tree and blob is recomputed and compared with its id before it is
  written to the build folder. A replace ref, `export-ignore` attributes or a substituted object in a local
  clone cannot change what is built. Paths that could escape the build folder are refused; submodules are
  skipped.
- **Build tools**: `cargo --locked` (checksums from rdocx's `Cargo.lock`), maturin and ziglang installed by
  pip with `--require-hashes --only-binary :all:` from `scripts/build-requirements.txt` and
  `scripts/cross-requirements.txt`, in environments recreated when those files change; cargo-zigbuild at a
  fixed version, checked before a cross-build.
- **Local builds**: files built in the same run are installed as a local build (`installed.json` records
  their hashes and every installed file); `status`, `install` and `test` accept one only where the lock has
  no hashes for the platform, or with `--allow-local`. Builds are not
  bit-reproducible across machines: record new hashes with `lock --write` only after checking where the
  files came from, because the lock is the only thing another machine trusts.

## Where prebuilt files are looked for

`$RDOCX_DIST/<commit>/<platform>/`, then `dist/<commit>/<platform>/` in this repository, then
`$RDOCX_HOME/dist/<commit>/<platform>/` (where `build` writes when the repository is read-only). A CLI may be
stored gzipped (`rdocx.gz`, for transfers that limit file sizes): its hash is checked after decompression.
`dist/` is not committed to git.

## Where the source is looked for (build)

`$RDOCX_SRC`, then a folder named `rdocx` next to this repository: the first git clone that contains the
pinned commit. Without one, the pinned commit alone is fetched (`--depth 1`) from `upstream` into
`$RDOCX_HOME/src.git`. The clone's working tree and branches are never touched.

The build needs the Rust toolchain named in rdocx's `rust-toolchain.toml` (install rustup from rustup.rs or
your package manager; the script never installs a toolchain). It takes 10 to 30 minutes on two cores; set
`CARGO_BUILD_JOBS` on small machines. Cross-building the other Linux architecture uses zig:
`cargo install --locked cargo-zigbuild --version 0.20.1`, then `build --target linux-aarch64`.

## Environments

- **A developer machine** (Claude Code): `install` finds `dist/<commit>/<platform>/` when the platform is
  recorded; otherwise `install --build`, once.
- **Small or short-lived sandboxes** (agent VMs with little memory or short command timeouts, where the
  home folder may be reset between sessions): never build there. Build or cross-build elsewhere, record the
  hashes, bring the dist folder in (a synced folder, or `install --from DIR` on staged files) and install
  once per session; it takes seconds.
- **No verified build available**: the skills tell the user and fall back to the built-in skills for the
  task at hand, and offer `install --build` in the background.

## Moving the pin

1. Pick the new upstream commit; update `commit` and `versions` in `rdocx.lock.json`, empty `artifacts` and
   `installed_files`.
2. `build` on each platform (or `build --target` from Linux), then `lock --write --platform P` for each.
3. `install`, then `test`. Every strict xfail that now passes means a gap is closed: remove its marker and
   its entry in `tests/gaps.py`, and update both skills' `references/gaps.md` and tables. Every new failure
   is a regression: report it upstream before moving the pin.
4. Commit the lock, the tests and the skill changes together.

## Later: releases

When rdocx publishes wheels on PyPI and binaries on GitHub releases, the lock gains a `"source": "release"`
form: the same per-platform SHA-256, the wheels installed with `pip install --require-hashes --no-deps`, the
CLIs downloaded from the release and checked before being made executable. The skills do not change.
