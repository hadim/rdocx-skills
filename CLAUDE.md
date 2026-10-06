# rdocx-skills: working notes for agents

A Claude plugin (`rdocx`) with two auto-loading skills, `docx` and `pptx`, that make agents use the rdocx /
rpptx CLIs and Python bindings for every Office file task, plus the acceptance suite that decides whether a
new rdocx build can be pinned. Public repository: nothing specific to a company or a project goes in here
(no document names, paths, people, or private file content; fixtures are synthetic).

## Map

- `skills/<name>/SKILL.md`: what an agent reads first; keep it under 200 lines, details go to `references/`
  (`tests/test_skill_docs.py` checks the size, and every Python name, keyword, CLI command and flag the skills cite).
- `skills/<name>/references/`: `cli.md`, `python-api.md`, `recipes.md` (every block is executed by
  `tests/test_docs_snippets.py`), `gaps.md`, and `interop.md` for docx.
- `skills/<name>/scripts/`: helpers importable and runnable as commands; tested in `tests/test_skill_scripts.py`.
- `scripts/rdocx_env.py`: install (from a dist folder, else the releases), build, verify, bump, lock, test;
  standard library only, Python >= 3.9, Linux, macOS and Windows (Git Bash: `.exe` CLIs, `current` a junction).
- `rdocx.lock.json`: the pin (upstream commit and ref, release URLs, SHA-256 per platform). `dist/`
  (gitignored) holds local builds; downloads land in `$RDOCX_HOME/dist/`.
- `.github/workflows/`: `ci.yml` (suite on linux-x86_64, macos-arm64, macos-x86_64 and windows-x86_64, every
  push and PR), `build.yml` (only for a commit between two upstream releases: builds it on five runners and
  publishes the release `rdocx-<YYYYMMDD>-<commit12>`, UTC build date first; by hand, and weekly on upstream
  `main` HEAD as a candidate).
- `tests/`: acceptance suite; `tests/gaps.py` is the registry of known gaps (strict xfails).
- `findings/`: gaps met in real use, waiting for triage into tests (see `findings/README.md`); empty now.
- `docs/setup.md`: install, trust model, CI, how to move the pin.

## Rules

- English everywhere, ASCII in code. Findings as `findings/README.md` describes them: what fails, a neutral
  reproduction that builds its own input, the real output, an acceptance criterion.
- Gaps live here as strict xfails (`tests/gaps.py`) and narrow what the skills claim; the suite run on each
  new build says which ones closed. When a gap is reported upstream, a comment on its entry in `tests/gaps.py`
  names the issue (`# upstream: owner/repo#N`); the suite, not the issue's state, says when it closed.
- Never weaken a test to make it pass. A behaviour change of rdocx is either a fix (remove the gap marker) or
  a regression (keep the test red and the previous pin, or narrow the skills).
- Every claim in the skills is backed by a test or a recipe block that runs. Changing a claim means changing
  or adding the test in the same commit.
- The lock records the files of the upstream releases (`lock --write --release`, which checks their
  `SHA256SUMS` and their build provenance with `gh attestation verify`), after review; a commit between two
  upstream releases is pinned through a release published by `build.yml` instead.
- The plugin has no version number: each commit of `main` is a version, so `main` always pairs a lock whose
  release exists and is recorded with the tests and skills that match it (docs/setup.md, Moving the pin).
- Compare file contents in asserts through `conftest.digest`: pytest diffs byte strings in full under
  `CI=true`, which took 20 minutes per failing comparison on the runners.
- Commits: Conventional Commits.

## State on 06/10/2026

- Repository: github.com/hadim/rdocx-skills, public; `main` protected (no force push, no deletion).
- Pinned, until the next upstream release: `integration/open-prs-2026-10-06` of hadim/rdocx (`e22641a8`,
  upstream main `20888b7a` plus tensorbee/rdocx#271, the fix of gap comment-several-paragraphs), through this
  repository's release `rdocx-20261006-e22641a8f20a` (five platforms). Next upstream release that carries
  #271: `bump v<version>` with `upstream` set back to tensorbee/rdocx.
- Before it, pinned: the upstream releases tensorbee/rdocx `v0.15.0` (rdocx 0.15.0) and `rpptx-v0.13.1` (rpptx 0.13.1),
  both at commit `9d019472` (sprint S88), the first ones that ship the CLIs and the wheels together under one
  tag per family, with `SHA256SUMS` and build provenance attestations (tensorbee/rdocx#266). The lock records
  five platforms: linux-x86_64, linux-aarch64 (CLIs need glibc 2.35, wheels manylinux_2_28), macos-arm64,
  macos-x86_64 (new: upstream builds it) and windows-x86_64. This repository's releases and the integration
  branches of hadim/rdocx are no longer used for the pin; `build.yml` stays for a commit between two
  upstream releases.
- Suite on macos-arm64: no gap closed or opened, no other failure against 0.15.0 / 0.13.1.
- The skills load ahead of Anthropic's docx and pptx skills: 8 of 8 headless sessions (summary, replace,
  new memo; add slide, outline, new deck) called `rdocx:docx` / `rdocx:pptx` first, with the
  document-skills plugin loaded alongside.

## Next steps

1. When rdocx tags a release: `bump v<version>`, `lock --write --release`, `install`, `test`, as
   docs/setup.md (Moving the pin) describes. The weekly `build.yml` run on upstream `main` says beforehand
   which gaps the next release closes.
