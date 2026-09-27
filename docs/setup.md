# Installing the pinned rdocx build

The skills never use whatever rdocx happens to be installed. They use the build named in `rdocx.lock.json`:

- `upstream` and `commit`: the rdocx source, by full 40-character git hash; `ref`: the tag or branch it was
  pinned from (for the reader);
- `release`: where this repository's build workflow published the files of that commit;
- `artifacts.<platform>`: the SHA-256 of the two CLIs and the two wheels built from that commit, per
  platform (`linux-x86_64`, `linux-aarch64`, `macos-arm64`);
- `installed_files.<platform>`: the SHA-256 of every file the two wheels install (modules, stubs, native
  libraries), checked again by `status` on the installed copy.

`scripts/rdocx_env.py` (standard library only, Python 3.9 or later, Linux and macOS; Windows is not
supported) enforces it.

## Commands

```bash
python3 scripts/rdocx_env.py status [--allow-local]  # exit 0 when the install matches the lock (or is an accepted local build)
python3 scripts/rdocx_env.py install                 # from a dist folder, else the release; exit 2 if neither has the files
python3 scripts/rdocx_env.py install --build         # compile the pinned commit when neither has them
python3 scripts/rdocx_env.py install --from DIR      # from a folder of prebuilt files, checked against the lock
python3 scripts/rdocx_env.py build [--target linux-aarch64]
python3 scripts/rdocx_env.py lock [--write] [--platform P]  # compare (or record) a local dist folder
python3 scripts/rdocx_env.py lock --write --release  # record every platform of the release named in the lock
python3 scripts/rdocx_env.py bump REF                # pin an upstream tag, branch or full commit hash
python3 scripts/rdocx_env.py test [pytest args]      # the acceptance suite on the installed build
python3 scripts/rdocx_env.py paths
```

The install goes to `~/.local/share/rdocx-skills/<first 12 characters of the commit>/` (override the root
with `RDOCX_HOME`), with `current/bin/{rdocx,rpptx,python}`: two CLIs and a Python wrapper whose
environment has both modules. Nothing is installed system-wide. `install` from a dist folder uses no
network, from the release it downloads four files; `build` fetches the source and the pinned build tools,
`test` the pinned test dependencies.

## Trust model

- **The lock is the only trust root.** Whatever the files come from (a dist folder, a download, a build),
  they are installed only if they match it, or, for files built in the same run, as a local build.
- **Prebuilt files**: each artifact is copied once into a private temporary folder, hashed there, compared
  with the lock, and installed from that copy only, so a file changed after the check is never used. The
  wheels are installed with `pip --no-index --no-deps --no-compile`. `status` re-hashes the CLIs, the
  `bin/python` wrapper and every file of the two packages against the lock, refuses any other file in them
  (compiled caches included: the wrapper sets `PYTHONDONTWRITEBYTECODE`), checks the interpreter link,
  `pyvenv.cfg` and the start-up files of site-packages (`.pth`, `sitecustomize`) against their state at
  install, and checks that `current` points at the pinned install; any mismatch exits 1, and `install`
  repairs it. Someone who can rewrite both the install and its `installed.json` can still defeat these
  checks: they detect corruption and stray changes, not an attacker who owns the home folder.
- **Releases**: a downloaded file lands in `$RDOCX_HOME/dist/<commit>/<platform>/` and goes through the
  same staging and hash check as any dist folder; one that does not match the lock stops the install (no
  fallback to a build). `lock --write --release` hashes the downloaded files on the reviewer's machine,
  checks them against the release's `SHA256SUMS`, and records them; each asset also carries a build
  provenance attestation (`gh attestation verify FILE -R hadim/rdocx-skills`).
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
  no hashes for the platform, or with `--allow-local`. Builds are not bit-reproducible across machines: the
  lock records the release's files, built once by the build workflow, rather than each machine's own.

## Where prebuilt files are looked for

`$RDOCX_DIST/<commit>/<platform>/`, then `dist/<commit>/<platform>/` in this repository, then
`$RDOCX_HOME/dist/<commit>/<platform>/` (where `build` writes when the repository is read-only, and where
downloads land), then the release named by `release` in the lock (or `$RDOCX_RELEASE_URL`, for a mirror):
assets named `<platform>.<file>`, listed with their SHA-256 in `SHA256SUMS`. A CLI in a dist folder may be
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

- **Any machine with network access** (Claude Code, agent VMs): `install` downloads the pinned files from the
  release and verifies them; it takes seconds.
- **No network**: bring a dist folder in (a synced folder, or `install --from DIR` on staged files).
- **Small or short-lived sandboxes** (little memory, short command timeouts, a home folder reset between
  sessions): never build there; `install` once per session.
- **No verified build available**: the skills tell the user and fall back to the built-in skills for the
  task at hand, and offer `install --build` in the background.

The Linux files are built on Ubuntu 22.04 runners: they need glibc 2.35 or later.

## Continuous integration

- `.github/workflows/ci.yml`, on every push to `main` and every pull request: `install --build` on
  linux-x86_64 (the release download, or a build right after a bump), then the acceptance suite.
- `.github/workflows/build.yml`: builds one upstream commit on three native runners (ubuntu-22.04,
  ubuntu-22.04-arm, macos-14), runs the suite on each, and publishes the release
  `rdocx-<first 12 characters of the commit>` with `SHA256SUMS`, provenance attestations, and the suite
  summaries in its notes (gaps closed, other failures). It runs on a bump (the lock changed on `main`), by
  hand on any upstream tag, branch or commit (Actions, build, Run workflow, `ref`), and weekly on upstream
  `main` HEAD. An existing release is never rebuilt or overwritten; a failing suite does not block it.

## Moving the pin

Each commit of `main` is a version of the plugin (it has no version number), so `main` always holds a lock
whose release exists and whose hashes are recorded, together with the tests and skills that match it:

1. Build the candidate: Actions, build, Run workflow, with `ref` set to an upstream tag, branch or full commit
   hash. It publishes the release `rdocx-<commit12>`, with the suite's results on three platforms in its
   notes (gaps closed, other failures). The weekly run does the same for upstream `main` HEAD.
2. On a branch: `bump REF` (the commit, the release name, empty hashes), then `lock --write --release`,
   `install` and `test`. Every strict xfail that now passes means a gap is closed: remove its marker and its
   entry in `tests/gaps.py`, and update both skills' `references/gaps.md` and tables. Every new failure is a
   regression: keep the previous pin, or narrow what the skills claim.
3. Commit the lock, the tests and the skill changes together, and merge into `main`.

A bump pushed to `main` before its release exists still works: the build workflow runs on the change of the
lock and publishes the release; `lock --write --release` then fills the hashes in a second commit. Until
then, `install` has nothing to download and the skills fall back to the built-in ones.

## Later: upstream releases

When rdocx publishes binaries on its GitHub releases and wheels on PyPI, the lock can name them instead of
this repository's releases (the download step then follows their file names); the per-platform SHA-256 in the
lock and the checks stay the same, and the build workflow is no longer needed.
