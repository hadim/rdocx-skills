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
- `scripts/rdocx_env.py`: install, build, verify, test; standard library only, Python >= 3.9.
- `rdocx.lock.json`: the pin (upstream commit, SHA-256 per platform). `dist/` (gitignored) holds the built files.
- `tests/`: acceptance suite; `tests/gaps.py` is the registry of known gaps (strict xfails).
- `findings/`: gaps met in real use, waiting for triage (see `findings/README.md`).
- `docs/setup.md`: install, environments, how to move the pin.

## Rules

- English everywhere, ASCII in code. Reports and tickets in the style of `findings/*.md`: what fails, a
  neutral reproduction that builds its own input, the real output, an acceptance criterion.
- Never weaken a test to make it pass. A behaviour change of rdocx is either a fix (remove the gap marker) or
  a regression (keep the test red, report upstream, do not move the pin).
- Every claim in the skills is backed by a test or a recipe block that runs. Changing a claim means changing
  or adding the test in the same commit.
- The lock only records files built from the pinned commit by `rdocx_env.py build`, after review.
- Commits: Conventional Commits.

## State on 27/09/2026

- Pinned: tensorbee/rdocx `9a7ed714` (sprint S75), rdocx 0.14.0, rpptx 0.12.1.
- Recorded in the lock, with `installed_files`: `linux-x86_64` (native build) and `linux-aarch64`
  (cross-built with zig). Not yet: `macos-arm64`.
- Suite on linux-x86_64: 367 passed, 85 strict xfails (55 gap keys); with `--runxfail`, every gap test fails
  for its stated reason. `test` runs from its own environment (`$RDOCX_HOME/testenv`), so the runtime install
  stays verifiable.
- Gap keys by ticket state (`tests/gaps.py`): 38 `lot6-NN` (filed by hand, numbers to add), 3 `filed 27/09`
  (numbers to add), 14 `not filed yet` (found by the acceptance reviews; their tests are the reproductions).
- `findings/`: the three reproductions filed on 27/09, kept until their numbers are in `tests/gaps.py`.

## Next steps

1. Create the GitHub repository (public), push, protect `main`.
2. Build `macos-arm64` natively: `python3 scripts/rdocx_env.py build`, check, `lock --write`, `install`,
   `test`. Commit the lock.
3. Replace the `lot6-NN` and `filed 27/09` placeholders in `tests/gaps.py` with the upstream issue numbers
   (`gh issue list -R tensorbee/rdocx -s all --limit 100`), then delete the matching `findings/` files.
4. Draft one upstream ticket per `not filed yet` gap, in the style of `findings/*.md`, from its test in
   `tests/`; the maintainer files them by hand.
5. CI (GitHub Actions): on push and PR, `install --build` on linux-x86_64 (cache `$RDOCX_HOME/target` and the
   cargo registry) then `test`; a scheduled job that builds upstream `main` HEAD in a scratch lock and runs
   the suite, to see gaps closing and regressions arriving before a pin move. Pin third-party actions by
   commit SHA.
6. Install the plugin where it is used and check that the skills load on their own for a .docx and a .pptx
   task, ahead of the built-in `docx` and `pptx` skills.
7. When rdocx publishes releases: add the `release` source to the lock and `rdocx_env.py` (docs/setup.md).
