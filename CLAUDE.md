# rdocx-skills: working notes for agents

A Claude plugin (`rdocx`) with two auto-loading skills, `docx` and `pptx`, that make agents use the rdocx /
rpptx CLIs and Python bindings for every Office file task, plus the acceptance suite that decides whether a
new rdocx build can be pinned. Public repository: nothing specific to a company or a project goes in here
(no document names, paths, people, or private file content; fixtures are synthetic).

## Map

- `skills/<name>/SKILL.md`: what an agent reads first; keep it under 16 KB and its description under 800 characters,
  details go to `references/` (`tests/test_skill_docs.py` checks the sizes, every Python name, keyword, CLI command and
  flag the skills cite, and the other way round that every public class and CLI command of the build is cited or
  listed in its `NOT_DOCUMENTED` with a reason; methods and flags an agent finds in the stubs and `--help`).
- `skills/<name>/references/`: `cli.md`, `python-api.md`, `recipes.md` (every block is executed by
  `tests/test_docs_snippets.py`), `gaps.md`, and `interop.md` for docx.
- `skills/<name>/scripts/`: helpers importable and runnable as commands; tested in `tests/test_skill_scripts.py`.
- `scripts/rdocx_env.py`: install (from a dist folder, else the releases), build, verify, bump, lock, test;
  standard library only, Python >= 3.9, Linux, macOS and Windows (Git Bash: `.exe` CLIs, `current` a junction).
- `rdocx.lock.json`: the pin (upstream commit and ref, release URLs, SHA-256 per platform). `dist/`
  (gitignored) holds local builds; downloads land in `$RDOCX_HOME/dist/`.
- `.github/workflows/`: `ci.yml` (the suite on four platforms, every push and PR, Intel Macs left out; on a push to `main` whose lock
  pins one of this repository's releases, its summaries become the "Acceptance suite" section of that release's
  notes, through `.github/scripts/release_notes.py`), `build.yml` (only for a commit between two upstream releases, by
  hand: builds it on four runners (no macos-x86_64), checks that it installs and runs, and publishes the release
  `rdocx-<YYYYMMDDHHMM>-<commit12>`, UTC build date and time first since GitHub lists releases by tag name (older
  releases: `rdocx-<YYYYMMDD>-<commit12>`, still valid); no suite there; inputs `lto` (`thin` by default, about half
  the build time of `fat`, upstream's profile) and `dry_run` (build from scratch, publish nothing, to time a profile)).
  `.github/scripts/suite-summary.sh` groups a suite log (gaps closed, docs not yet updated, other failures); both
  scripts are tested in `tests/test_ci_scripts.py`.
- `tests/`: acceptance suite; `tests/gaps.py` is the registry of known gaps (strict xfails).
- `findings/`: gaps met in real use, waiting for triage into tests (see `findings/README.md`); empty now.
- `docs/setup.md`: install, trust model, CI, how to move the pin.

## Rules

- English everywhere, ASCII in code. Findings as `findings/README.md` describes them: what fails, a neutral
  reproduction that builds its own input, the real output, an acceptance criterion.
- Tests cover the skills' own contract: every recipe block runs (`test_docs_snippets.py`), the cited names exist and
  every class and command has a pointer (`test_skill_docs.py`), the helper scripts, the installer, the CI scripts.
  rdocx's behaviour is upstream's to test: the behaviour files (`test_docx_*.py`, `test_pptx*.py`) only hold what the
  pinned build has and upstream `main` does not yet (a fix or a feature carried on an integration branch, or a gap
  as a strict xfail in `tests/gaps.py`, with `# upstream: owner/repo#N` when reported). When the pin moves to an
  upstream release, run them against it and delete those that pass. A regression that hits the skills gets a test
  then, and an upstream issue.
- A behaviour change of rdocx is either a fix (remove the gap marker) or a regression (keep the test red and the
  previous pin, or narrow the skills); never weaken a test to make it pass.
- The lock records the files of the upstream releases (`lock --write --release`, which checks their
  `SHA256SUMS` and their build provenance with `gh attestation verify`), after review; a commit between two
  upstream releases is pinned through a release published by `build.yml` instead.
- The plugin has no version number: each commit of `main` is a version, so `main` always pairs a lock whose
  release exists and is recorded with the tests and skills that match it (docs/setup.md, Moving the pin).
- Moving the pin, a new integration snapshot included, always brings the skills along in the same PR, without being
  asked: fix every cited name that moved, keep the recipes running, make the behaviour tests pass, and document what
  the new build adds that an agent would use (a SKILL.md table row or a references entry, a recipe when it helps). The
  coverage check only asks a pointer per public class and CLI command, so a new method or flag gets no line unless
  this pass writes it. Onto an upstream release, also delete the behaviour tests that pass on it. Stop before the
  merge: a short brief, and the merge asked as a question.
- No release is asked of upstream: the pin takes what upstream publishes, and a commit needed between two
  upstream releases goes through `build.yml` (above).
- Git on a clone reached through a shell on another machine (a remote or sandboxed shell over a synced folder):
  run git with `GIT_OPTIONAL_LOCKS=0`; switching branches needs the permission to delete files there. The
  session commits on a branch; the maintainer pushes and files any upstream issue.
- Compare file contents in asserts through `conftest.digest`: pytest diffs byte strings in full under
  `CI=true`, which took 20 minutes per failing comparison on the runners.
- Commits: Conventional Commits.

## State on 12/10/2026

- Repository: github.com/hadim/rdocx-skills, public; `main` protected (no force push, no deletion).
- Pinned: the integration branch `integration/open-prs-2026-10-12-4` of hadim/rdocx (commit `52e7f0b8`) through this
  repository's release `rdocx-202610121958-52e7f0b8e400` (ThinLTO, four platforms, no macos-x86_64): tensorbee/rdocx
  `v0.16.0` (`733ce6f7`) plus every open PR of the user (the previous pin `integration/open-prs-2026-10-12-3`,
  `12b1d6fc`) and tensorbee/rdocx#363, which lists range paragraphs reading each story source once. v0.16.0 had made
  `rebuild_toc`, the page field updates and the comment listing quadratic or worse (tensorbee/rdocx#362): on the
  13-page fixture 17 s, 7 s and 4.2 s, now 1.2 s, 0.2 s and 0.2 s (26 pages: TOC 134 s, now 2.6 s).
- Open gaps: none.
- Suite against `52e7f0b8`: 356 passed, 1 skipped (Poppler absent) on macos-arm64 from the release files, in 1.5
  minutes. The behaviour files hold the tests that fail on upstream v0.16.0, that is the PRs carried on the
  integration branch: when the pin moves to an upstream release that carries them, run them against it and delete
  those that pass.
- The skills load ahead of the built-in docx and pptx skills: 8 of 8 headless sessions (4 Word, 4 PowerPoint tasks)
  called `rdocx:docx` / `rdocx:pptx` first with the shortened descriptions, the built-in skills' descriptions loaded
  alongside (2026-10-12).
- A skills `main` pinned on the fork steps back from upstream releases: go back to tensorbee/rdocx with `bump
  v<version>` once a release carries these PRs. Older pins: `git log -p rdocx.lock.json`.

## Next steps

1. When rdocx tags a release: `bump v<version>`, `lock --write --release`, `install`, `test`, as
   docs/setup.md (Moving the pin) describes.
