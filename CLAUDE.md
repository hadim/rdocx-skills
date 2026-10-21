# rdocx-skills: working notes for agents

A Claude plugin (`rdocx`) with two auto-loading skills, `docx` and `pptx`, that make agents use the rdocx /
rpptx CLIs and Python bindings for every Office file task, plus the acceptance suite that decides whether a
new rdocx build can be pinned. Public repository: nothing specific to a company or a project goes in here
(no document names, paths, people, or private file content; fixtures are synthetic).

## Map

The plugin is the `plugin/` folder (what users install and the plugin directory scans: manifest, skills, scripts, the
lock, the icon and README pictures); the marketplace entry (`.claude-plugin/marketplace.json`), the tests, the
workflows, the docs and this file sit beside it. Keep anything the plugin does not ship out of `plugin/`, and keep
`README.md` and `plugin/README.md` in step (`tests/test_readme.py`).

- `plugin/skills/<name>/SKILL.md`: what an agent reads first; keep it under 16 KB and its description under 800 characters,
  details go to `references/` (`tests/test_skill_docs.py` checks the sizes, every Python name, keyword, CLI command and
  flag the skills cite, and the other way round that every public class and CLI command of the build is cited or
  listed in its `NOT_DOCUMENTED` with a reason; methods and flags an agent finds in the stubs and `--help`).
- `plugin/skills/<name>/references/`: `cli.md`, `python-api.md`, `recipes.md` (every block is executed by
  `tests/test_docs_snippets.py`), `gaps.md`, and `interop.md` for docx.
- `plugin/skills/<name>/scripts/`: helpers importable and runnable as commands; tested in `tests/test_skill_scripts.py`.
- `plugin/scripts/rdocx_env.py`: install (from a dist folder, else the releases), build, verify, bump, lock, test;
  standard library only, Python >= 3.9, Linux, macOS and Windows (Git Bash: `.exe` CLIs, `current` a junction).
- `plugin/rdocx.lock.json`: the pin. On `main`, the upstream commit and ref only (no release, no hashes: CI builds
  it); on `release`, also the plugin release that carries its files and their SHA-256 per platform. `dist/`
  (gitignored) holds local builds; downloads land in `$RDOCX_HOME/dist/`.
- `.github/workflows/`: `ci.yml` (every push and PR: builds the pinned commit on four runners, Intel Macs left out,
  with ThinLTO and a build cache per platform and commit, records the build into the run's copy of the lock, and
  runs the suite on it), `release.yml` (by hand: a release of the plugin, below, which builds the pinned commit again
  on the four platforms, runs the suite on those files and attaches them to the release; older releases stay;
  input `dry_run`). The plugin's releases are the only published builds.
  `.github/scripts/suite-summary.sh` groups a suite log (gaps closed, docs not yet updated, other failures);
  `plugin_release.py` names the plugin's releases, writes their version and the head of their notes; the scripts are tested in
  `tests/test_ci_scripts.py`.
- `tests/`: acceptance suite; `tests/gaps.py` is the registry of known gaps (strict xfails).
- `findings/`: gaps met in real use, waiting for triage into tests (see `findings/README.md`); empty now.
- `docs/setup.md`: install, trust model, CI, how to move the pin, how to release the plugin.
- `plugin/assets/`: the plugin icon (an SVG source and the PNG that plugin.json names for the directory listing) and
  the README's two renders.

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
- Only the plugin's releases publish builds: `main`'s lock names an upstream commit (`bump REF`), `ci.yml` builds it
  on every push and PR, and `release.yml` builds it again (even when it is unchanged), tests those files, attaches
  them to the `v<version>` release and points the lock of `release` at it. Older releases stay, each with its own
  files; delete them only when the maintainer asks.
- `main` always pairs a lock with the tests and skills that match it (docs/setup.md, Moving the pin), but a push
  to `main` reaches nobody: `release` is the default branch, which users
  (Claude Code, Claude Desktop) and the Claude plugin directory follow, and only `release.yml` moves it (CalVer
  `vYYYY.M.D`, `version` in plugin.json there only; docs/setup.md, Releasing the plugin). Pull requests target `main`
  (`gh pr create --base main`). Releasing is the maintainer's call: never run it unasked.
- Moving the pin, a new integration snapshot included, always brings the skills along in the same PR, without being
  asked: fix every cited name that moved, keep the recipes running, make the behaviour tests pass, and document what
  the new build adds that an agent would use (a SKILL.md table row or a references entry, a recipe when it helps). The
  coverage check only asks a pointer per public class and CLI command, so a new method or flag gets no line unless
  this pass writes it. Onto an upstream release, also delete the behaviour tests that pass on it. Stop before the
  merge: a short brief, and the merge asked as a question.
- No release is asked of upstream: the pin takes any upstream commit or tag, or a commit carried on an integration
  branch, and this repository builds it (above).
- Git on a clone reached through a shell on another machine (a remote or sandboxed shell over a synced folder):
  run git with `GIT_OPTIONAL_LOCKS=0`; switching branches needs the permission to delete files there. The
  session commits on a branch; the maintainer pushes and files any upstream issue.
- Compare file contents in asserts through `conftest.digest`: pytest diffs byte strings in full under
  `CI=true`, which took 20 minutes per failing comparison on the runners.
- No `.gitattributes`: the plugin directory refuses one that rewrites files at checkout (`text`, `eol`, `filter`).
  Files are committed with LF, and CI sets `core.autocrlf false` before checking out on Windows.
- Commits: Conventional Commits.

## State on 20/10/2026

- Repository: github.com/hadim/rdocx-skills, public; `main` protected (no force push, no deletion). Being
  submitted to the Claude plugin directory, following the default branch `release`; the rdocx maintainer agreed to
  the plugin name `rdocx` (hadim/rdocx-skills#40). First release `v2026.10.15` (pin `d17e1da9`); latest
  `v2026.10.16.2` (pin `d44a69af`).
- Pinned: hadim/rdocx `integration/open-prs-2026-10-20` at `6e286888`: tensorbee/rdocx `main` `b072b985` plus the
  open pull requests tensorbee/rdocx#393 to #397 and #399 (the fix of #398, Python handles that rescanned the whole
  document). Upstream does not merge pull requests from here: it adapts them in a sprint and closes them, so pin its
  `main` or its tags again once a sprint carries them.
- Open gaps: 3 (`tests/gaps.py`, both `references/gaps.md`): docx-custom-properties-not-listed (reported as
  tensorbee/rdocx#387); docx-toc-outline-switch-ignores-depth, a regression of this pin (2026-10-21: tensorbee/rdocx
  2a88a8f5, F-283, lets Heading 2 into `TOC \o "1-1" \u`; `7496a81c` passes it), and
  docx-toc-outline-level-beyond-depth, the same missing range check on a paragraph's own outline level, `7496a81c`
  included; `recipes.md` carries the workaround (take `\u` out for the rebuild). Every gap of `d44a69af` closed
  with `7496a81c` (tensorbee/rdocx#384, #385, #386); the python-docx and python-pptx names it does not take raise
  an error naming the rdocx or rpptx way, the contract #379 asked for.
- Suite against `6e286888`: 160 passed, 4 xfailed, 18 skipped (the install and release tests, without recorded
  hashes) on macos-arm64 from a local build, in under a minute. Walking every body and cell paragraph handle is
  linear with #399 (600 paragraphs and 965 cells: 0.2 s, nested-table cells about 15 ms each).
  `tests/test_docx_carried.py` and `tests/test_pptx_carried.py` hold what the carried pull requests fix: delete each
  test once upstream `main` has its fix.
- The skills load ahead of the built-in docx and pptx skills: 8 of 8 headless sessions (4 Word, 4 PowerPoint tasks)
  called `rdocx:docx` / `rdocx:pptx` first with the shortened descriptions, the built-in skills' descriptions loaded
  alongside (2026-10-12).
- Older pins: `git log -p --follow plugin/rdocx.lock.json`.

## Next steps

1. When upstream `main` moves (S96 brings fillable PDF forms, tensorbee/rdocx#381) or tags a release: `bump <ref>` on
   a branch, `install --build` and `test` (or let ci.yml build it on the pull request), as docs/setup.md (Moving the
   pin) describes; delete the gap tests that pass; release the plugin when the maintainer asks.
