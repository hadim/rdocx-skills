# rdocx-skills: working notes for agents

A Claude plugin (`rdocx`) with two auto-loading skills, `docx` and `pptx`, that make agents use the rdocx /
rpptx CLIs and Python bindings for every Office file task, plus the acceptance suite that decides whether a
new rdocx build can be pinned. Public repository: nothing specific to a company or a project goes in here
(no document names, paths, people, or private file content; fixtures are synthetic).

## Map

- `skills/<name>/SKILL.md`: what an agent reads first; keep it under 200 lines, details go to `references/`.
- `skills/<name>/references/`: `cli.md`, `python-api.md`, `recipes.md` (every block is executed by
  `tests/test_docs_snippets.py`), `gaps.md`, and `interop.md` for docx.
- `skills/<name>/scripts/`: helpers importable and runnable as commands; tested in `tests/test_skill_scripts.py`.
- `scripts/rdocx_env.py`: install (from a dist folder, else the release), build, verify, bump, lock, test;
  standard library only, Python >= 3.9.
- `rdocx.lock.json`: the pin (upstream commit and ref, release URL, SHA-256 per platform). `dist/`
  (gitignored) holds local builds; downloads land in `$RDOCX_HOME/dist/`.
- `.github/workflows/`: `ci.yml` (suite on every push and PR), `build.yml` (builds one upstream commit on
  three runners and publishes the release `rdocx-<commit12>`).
- `tests/`: acceptance suite; `tests/gaps.py` is the registry of known gaps (strict xfails).
- `findings/`: gaps met in real use, waiting for triage into tests (see `findings/README.md`); empty now.
- `docs/setup.md`: install, trust model, CI, how to move the pin.

## Rules

- English everywhere, ASCII in code. Findings as `findings/README.md` describes them: what fails, a neutral
  reproduction that builds its own input, the real output, an acceptance criterion.
- Gaps live here as strict xfails (`tests/gaps.py`) and narrow what the skills claim; the suite run on each
  new build says which ones closed. Which upstream issue or pull request fixes a gap is not tracked here.
- Never weaken a test to make it pass. A behaviour change of rdocx is either a fix (remove the gap marker) or
  a regression (keep the test red and the previous pin, or narrow the skills).
- Every claim in the skills is backed by a test or a recipe block that runs. Changing a claim means changing
  or adding the test in the same commit.
- The lock records the files of a release published by `build.yml` (`lock --write --release`), after review.
- The plugin has no version number: each commit of `main` is a version, so `main` always pairs a lock whose
  release exists and is recorded with the tests and skills that match it (docs/setup.md, Moving the pin).
- Compare file contents in asserts through `conftest.digest`: pytest diffs byte strings in full under
  `CI=true`, which took 20 minutes per failing comparison on the runners.
- Commits: Conventional Commits.

## State on 29/09/2026

- Repository: github.com/hadim/rdocx-skills, public; `main` protected (no force push, no deletion).
- Pinned: hadim/rdocx `f2fa36d1` (`integration/open-prs-2026-09-29-2`: tensorbee/rdocx `main` at `9a7ed714`,
  sprint S75, plus the 40 open pull requests #173 to #212, not yet reviewed upstream), rdocx 0.14.0, rpptx
  0.12.1. That branch is never rewritten nor deleted: `install --build` fetches the pinned commit from it, and a
  later snapshot gets a new dated branch (the first one, `integration/open-prs-2026-09-29` at `f3df95bf`, held
  #173 to #195). Release `rdocx-f2fa36d1d18a` built by `build.yml` for linux-x86_64, linux-aarch64
  (ubuntu-22.04 runners, glibc 2.35) and macos-arm64 (macos-14), with provenance attestations. The lock records
  its hashes. Back to tensorbee/rdocx once those pull requests land there.
- Suite: 462 passed, 10 strict xfails (8 gap keys), 12 skipped (installer and release tests without hashes in
  the lock) on linux-x86_64 in `ci.yml` and on macos-arm64 with the local build (29/09/2026, 21 more gaps
  closed by the pin, 53 of 61 in all). 474 passed and 10 xfails on macos-arm64 with the release installed.
- The skills load ahead of Anthropic's docx and pptx skills: 8 of 8 headless sessions (summary, replace,
  new memo; add slide, outline, new deck) called `rdocx:docx` / `rdocx:pptx` first, with the
  document-skills plugin loaded alongside.

## Next steps

1. When rdocx tags a release or `main` moves: run `build.yml` by hand on the ref, read the suite summary in
   the release notes, then bump as docs/setup.md describes.
2. When rdocx publishes binaries on GitHub releases and wheels on PyPI: point the lock at them and retire
   `build.yml` (docs/setup.md, Later: upstream releases).
