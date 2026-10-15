# Installing the pinned rdocx build

The skills never use whatever rdocx happens to be installed. They use the build named in `rdocx.lock.json`:

- `upstream` and `commit`: the rdocx source, by full 40-character git hash; `ref`: the tag or branch it was
  pinned from (for the reader);
- `release`: the URL of this repository's build of that commit, where its files are published (empty after a
  bump, until `lock --write --release` finds the build and records it). Every pin is such a build, an upstream
  release tag as much as an integration branch;
- `versions`: the rdocx and rpptx versions the wheels carry, which name the builds and the plugin's releases;
- `artifacts.<platform>`: the SHA-256 of the two CLIs and the two wheels built from that commit, per
  platform (`linux-x86_64`, `linux-aarch64`, `macos-arm64`, `macos-x86_64`, `windows-x86_64`);
- `installed_files.<platform>`: the SHA-256 of every file the two wheels install (modules, stubs, native
  libraries), checked again by `status` on the installed copy.

`scripts/rdocx_env.py` (standard library only, Python 3.9 or later, Linux, macOS and Windows) enforces it.

## Commands

```bash
python3 scripts/rdocx_env.py status [--allow-local]  # exit 0 when the install matches the lock (or is an accepted local build)
python3 scripts/rdocx_env.py install                 # from a dist folder, else the release; exit 2 if neither has the files
python3 scripts/rdocx_env.py install --build         # compile the pinned commit when neither has them
python3 scripts/rdocx_env.py install --from DIR      # from a folder of prebuilt files, checked against the lock
python3 scripts/rdocx_env.py build [--target linux-aarch64]
python3 scripts/rdocx_env.py lock [--write] [--platform P]  # compare (or record) a local dist folder
python3 scripts/rdocx_env.py lock --write --release  # record every platform of the release named in the lock (found first if none)
python3 scripts/rdocx_env.py bump REF                # pin an upstream tag, branch or full commit hash (release and hashes emptied)
python3 scripts/rdocx_env.py test [pytest args]      # the acceptance suite on the installed build
python3 scripts/rdocx_env.py paths
```

The install goes to `~/.local/share/rdocx-skills/<first 12 characters of the commit>/` (override the root
with `RDOCX_HOME`), with `current/bin/{rdocx,rpptx,python}`: two CLIs and a Python wrapper whose
environment has both modules. The wrapper is a shell script that sets `PYTHONDONTWRITEBYTECODE` and
`PYTHONUTF8` (UTF-8 for files and pipes whatever the locale) and runs the venv's interpreter. Nothing is
installed system-wide. `install` from a dist folder uses no
network, from the releases it downloads four files; `build` fetches the source and the pinned build tools,
`test` the pinned test dependencies. Once a pin is installed, `install` removes the installs of older pins
beyond the most recent one, with their downloads and source exports (about 100 MB each; `RDOCX_KEEP_INSTALLS`
sets how many installs stay, 0 keeps them all); build caches and the test environment stay.

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
- **Releases**: the files are downloaded into `$RDOCX_HOME/dist/<commit>/<platform>/` and go through the same
  staging and hash check as any dist folder; a file that does not match the lock stops the install (no
  fallback to a build). `lock --write --release` hashes the downloaded files on the reviewer's machine, checks
  them against the release's `SHA256SUMS`, then records them. Each build carries a provenance attestation:
  `gh attestation verify FILE -R hadim/rdocx-skills` shows the workflow run that built it.
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
downloads land), then the release named by `release` in the lock: this repository's build of the commit, whose
assets are named `<platform>.<file>` and listed with their SHA-256 in `SHA256SUMS`.
`$RDOCX_RELEASE_URL` names a mirror: it stands for `https://github.com/<owner>/<repo>/releases/download` and
holds one folder per release tag. A CLI in a dist folder may be
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
  sessions): never build there; `install` once per session. Where the home folder is wiped between sessions
  or its disk is full, install under `/tmp` (seconds, hash-checked):
  `git clone --depth 1 https://github.com/hadim/rdocx-skills /tmp/rdocx-skills-src && RDOCX_HOME=/tmp/rdocx python3 /tmp/rdocx-skills-src/scripts/rdocx_env.py install`,
  and keep `RDOCX_HOME=/tmp/rdocx` set in each command, so that the skills' `R=${RDOCX_HOME:-...}/current/bin`
  finds it.
- **No verified build available**: the skills tell the user and fall back to the built-in skills for the
  task at hand, and offer `install --build` in the background.
- **Windows** (x86_64): the commands run in Git Bash, the shell Claude Code uses there, with `python` (or
  `py -3`) instead of `python3`, which is often absent or the Microsoft Store stub. The install sits in
  `%USERPROFILE%\.local\share\rdocx-skills`, so `~/.local/share/rdocx-skills` in Git Bash names the same
  folder; the CLIs are `rdocx.exe` and `rpptx.exe` (Git Bash finds them as `$R/rdocx` and `$R/rpptx`),
  `current` is a junction, which needs no privilege where a symlink would, and `bin/python` is the same shell
  wrapper, which Git Bash runs and `cmd.exe` or PowerShell cannot. The CLIs link the Microsoft Visual C++
  runtime (`vcruntime140.dll`), as Rust programs built with MSVC do. A build needs the MSVC build tools
  (Visual Studio Build Tools, C++ workload) besides rustup.

The Linux CLIs need glibc 2.35 or later (the wheels 2.28). The Windows files are built with MSVC.
Intel Macs (`macos-x86_64`) are neither built nor tested here: `install` exits 2 there and the skills hand over to
the built-in ones, or `install --build` compiles the pinned commit on the machine.

## Continuous integration

- `.github/workflows/ci.yml`, on every push to `main` and every pull request: `install --build` on
  linux-x86_64, linux-aarch64 (ubuntu-22.04-arm, the CLIs' glibc 2.35 floor), macos-arm64 and windows-x86_64 (the
  release download, or a build when the release lacks the platform's files, as right after a bump; its summary then
  says so), then the acceptance suite, summarised per platform by `.github/scripts/suite-summary.sh`. On Windows it
  runs from Git Bash, as an agent does, on a checkout made with `core.autocrlf false` (the repository has no
  `.gitattributes`: the plugin directory refuses one that rewrites files at checkout). On a push to `main`, a last
  job (the only one with `contents: write`) writes the four summaries into the notes of the build the lock pins
  (`rdocx-<YYYYMMDDHHMM>-<commit12>`, or `rdocx-<YYYYMMDD>-<commit12>` for the releases built before the time was
  added, of this repository and of the pinned commit), as the section `Acceptance suite (skills main <sha12>, <run
  URL>)` right after the notes' preamble: the results of the skills and tests that ship with the build. It runs
  after a failing leg too, and running it again replaces that section (`.github/scripts/release_notes.py`). The
  build stays a pre-release: the plugin's releases are the repository's latest (below).
- `.github/workflows/build.yml`, for every pin, by hand (Actions, build, Run workflow, `ref`: any upstream tag,
  branch or commit): builds it on four native runners (ubuntu-22.04, ubuntu-22.04-arm, macos-14, windows-2025),
  checks that each build installs, verifies and runs, and publishes the pre-release `rdocx-<YYYYMMDDHHMM>-<first
  12 characters of the commit>` (the UTC build date and time to the minute first: GitHub lists releases by tag
  name, so the time keeps two builds of the same day in build order) with `SHA256SUMS` and provenance
  attestations. Its title names the versions the wheels carry and what was built: `rdocx 0.16.1 · rpptx 0.14.1
  (build <date>, <commit12>)` for an upstream release tag, `rdocx 0.16.0 · rpptx 0.14.0 +
  integration/open-prs-2026-10-14 from hadim/rdocx (build ...)` for anything else, the repository named when it is
  not tensorbee/rdocx (`.github/scripts/plugin_release.py label`). Integration branches live on the maintainer's
  fork, hadim/rdocx: an upstream release plus the maintainer's pull requests still open upstream. It
  runs no suite: `ci.yml` runs it on the pull request that pins the release, and writes it into the notes once
  `main` pins it (above); releases built before carry a `Build-time suite (preliminary, ...)` section too. Inputs:
  `lto` (`thin`, the default: ThinLTO and 16 codegen units, about half the build time of `fat`, upstream's release
  profile, for slightly larger and slower binaries with the same behaviour; each has its own build cache) and `dry_run` (every platform
  built from scratch, nothing published: to time a profile). A commit has at most one release: when a release tag already ends with `-<commit12>`
  (in either format), its files are never rebuilt or overwritten, but a platform added to the workflow later is
  built and added to it (its files, its lines of `SHA256SUMS`; then `lock --write --release` records it). The
  releases built before the time was added keep their tag `rdocx-<YYYYMMDD>-<commit12>`, still valid wherever a
  tag is read (`-` sorts before a digit, so such a tag still lists below the later builds of its day).
- `lock --write --release` finds the build of the pinned commit: this repository's release whose tag ends with
  `-<commit12>` (through `git ls-remote --tags` of this repository, with no API call and no token; the tag cannot
  be derived from the commit, since the date and time are the build's). It records its URL: the lock's `release`
  is the only source of truth.
- `.github/workflows/release.yml`, by hand: publishes a release of the plugin (next section).

## Moving the pin

`main` always holds a lock whose build exists and whose hashes are recorded, together with the tests and skills
that match it, so that any commit of `main` can be released:

1. Build the commit first: Actions, build, Run workflow, with `ref` set to the upstream tag (`v<version>`), branch
   or full commit hash. It publishes the pre-release `rdocx-<YYYYMMDDHHMM>-<commit12>` without running the suite.
   Then, on a branch: `bump REF`, `lock --write --release` (it finds that build, checks its files and records its
   URL and hashes), `install` and `test`. Every strict xfail that now passes means a gap is
   closed: remove its marker and its entry in `tests/gaps.py`, and update both skills' `references/gaps.md`
   and tables. Every new failure is a regression: keep the previous pin, or narrow what the skills claim.
2. Bring the skills along on the same branch, without waiting to be asked. `tests/test_skill_docs.py` already fails
   on a cited Python name, keyword, CLI command or flag that the build lacks, on a public class or CLI command of the
   build that the skills neither cite nor exclude (give it a pointer, or add it to `NOT_DOCUMENTED` with a reason), on
   a gap page out of step with `tests/gaps.py`, and on a SKILL.md over 16 KB or a description over 800 characters.
   What it cannot judge:
   - what the new build adds: a new method or flag fails nothing, so document what an agent would use (a SKILL.md
     table row or a references entry, a recipe when it helps), or agents only find it by searching the stubs;
   - whether each sentence still holds: go through both SKILL.md files, their references (`cli.md`,
     `python-api.md`, `gaps.md`, `interop.md`) and `skills/*/scripts` with the new build installed, looking for a
     workaround a closed gap made useless, advice about a behaviour that changed, and a limitation stated without
     its gap key.
   When the new pin is an upstream release, also run the behaviour tests (`test_docx_*.py`, `test_pptx*.py`)
   against it and delete those that pass: upstream has that behaviour now.
3. Commit the lock, the tests and the skill changes together, and merge into `main`. CI runs the suite on the
   pinned files on four platforms; when the pin is one of this repository's
   releases, the push to `main` writes those results into the release notes as `Acceptance suite (skills main
   <sha12>, <run URL>)` (above the preliminary build-time suite of the releases built while `build.yml` still ran
   one). Every later push to `main` that keeps the pin replaces it with its own run.

On the pull request, failures of `tests/test_skill_docs.py` are the docs pass still to do, not regressions.
Merging into `main` publishes nothing: users get the new pin with the next release of the plugin.

## Releasing the plugin

Pushes to `main` reach nobody. `release` is the repository's default branch, so whoever adds the marketplace without
a ref (`/plugin marketplace add hadim/rdocx-skills`, the only form Claude Desktop takes) and the Claude plugin directory
follow it, and get a new version only from `.github/workflows/release.yml`: Actions, release, Run workflow. It:

1. refuses a `main` whose `ci.yml` run did not succeed, or that is released already;
2. picks the CalVer version of the day, `YYYY.M.D` (`YYYY.M.D.2` and on for another release the same day);
3. moves the `release` branch to `main`'s tree with that `version` written into `.claude-plugin/plugin.json`, as a
   merge of `main` so that `release` only moves forward, and tags it `v<version>`;
4. publishes the GitHub release `v<version>`, titled with the rdocx and rpptx versions of the pinned build
   (`rdocx 0.16.1 · rpptx 0.14.1`, or `... + <integration branch> from hadim/rdocx`), marked latest, whose notes name the pinned
   build and list the pull requests merged since the previous release.

`main` carries no `version`: whoever adds the marketplace as `hadim/rdocx-skills#main` gets every push, the
development channel. Pull requests target `main`, not the default branch: pick it as the base on GitHub, and run
`build.yml` from `main` (Use workflow from). The version, the title and the notes come from
`.github/scripts/plugin_release.py`, tested in `tests/test_ci_scripts.py`. The workflow pushes to `release` with the
workflow token: a protection rule on that branch must let it through.
