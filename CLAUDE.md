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
  with ThinLTO and a build cache per platform and commit, records the build into the run's copy of the lock, and runs
  the suite on it; a green run marks its tree, and a push to `main` whose tree passed already, a green pull request
  merged while `main` stood still, runs nothing again), `release.yml` (by hand: a release of the plugin, below, which
  takes the files `ci.yml` built, attested and tested on that tree (else builds them), runs the suite on them and
  attaches them to the release; older releases stay; input `dry_run`). The plugin's releases are the only published
  builds.
  `.github/scripts/suite-summary.sh` groups a suite log (gaps closed, docs not yet updated, other failures);
  `plugin_release.py` names the plugin's releases, writes their version and the head of their notes;
  `tested-tree.sh` tells whether `ci.yml` passed on a tree (`release.yml` takes that); the scripts are tested in
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
  upstream release, run them against it and delete those that pass, but `test_pptx_scaling.py` and
  `test_docx_scaling.py`: they time every public mutating call of rpptx and rdocx, and their passing calls guard
  against a regression. A regression that hits the skills gets a test then, and an upstream issue.
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
  this pass writes it. Onto an upstream release, also delete the behaviour tests that pass on it (not the scaling
  tests'). Stop before the merge: a short brief, and the merge asked as a question.
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

## State on 24/10/2026

- Repository: github.com/hadim/rdocx-skills, public; `main` protected (no force push, no deletion). Being
  submitted to the Claude plugin directory, following the default branch `release`; the rdocx maintainer agreed to
  the plugin name `rdocx` (hadim/rdocx-skills#40). First release `v2026.10.15` (pin `d17e1da9`); latest
  `v2026.10.23` (pin `a57be365`, the upstream release v0.16.3 / rpptx-v0.14.3).
- Pinned: `integration/open-prs-2026-10-23-8` on hadim/rdocx (`c0feead6`): the upstream release v0.16.3 /
  rpptx-v0.14.3 (tensorbee/rdocx `a57be365`) plus the open pull requests tensorbee/rdocx#409 to #422, fixes of the
  findings #402 to #408 and of the edit costs of #403 (rpptx: #413, then #416 stacked on it) and #407 (rdocx: #415,
  then the series #417 to #422 stacked on it). Upstream does not merge pull requests from here: it adapts them in a
  sprint and closes them, so pin its `main` or its tags again once a sprint carries them, and delete
  `tests/test_docx_carried.py` then (its tests name their pull requests).
- Open gaps: none (`tests/gaps.py` is empty, both `references/gaps.md` say so). Closed by this pin:
  docx-edit-cost-grows-with-body (#415, #417 to #422: 136 of the 145 rdocx calls that grew with the body; the
  maintainer accepted the linear cost of the nine others, called once or a few times per document, and of six of
  the 136 whose linear part shows on macOS runners, `ACCEPTED` in `test_docx_scaling.py`), pptx-edit-cost-grows-with-deck (#413, #416: none of the 248 rpptx calls grows with the
  deck), docx-default-line-spacing-multiple (#414), docx-part-xml-without-relationships (#409),
  docx-insert-document-duplicate-style-ids and docx-insert-document-new-document-refused (#410),
  docx-insert-document-reused-picture-swapped (#411), docx-comments-range-across-cells (#412, which adds
  `Comment.anchor_issue`). An edit still costs a little more in a larger body, accepted: at 1000 paragraphs
  `run.text` about 70 ms, a replacement 0.11 to 0.14 s (the docx gap page gives the costs).
- Suite against `c0feead6` on the four CI platforms (hadim/rdocx-skills#63): no failure, the scaling tests included.
  The calls that left GROWING of `test_docx_scaling.py` sit near FIXED for 35 of them, so that a noisy runner made
  them XPASS on some platforms and runs only: they left all the same, being far under RATIO. The gap marker is
  strict. The docx SKILL.md is 4 bytes under its 16,000-byte budget.
- `rdocx_env.py build` exports the source with the commit's time only over a build folder whose last build was that
  commit (`target/<platform>/built-commit`): a folder restored from another commit's cache, built after this commit was
  made, linked that commit's crates on Windows (2026-10-23).
- The skills load ahead of the built-in docx and pptx skills: 8 of 8 headless sessions (4 Word, 4 PowerPoint tasks)
  called `rdocx:docx` / `rdocx:pptx` first with the shortened descriptions, the built-in skills' descriptions loaded
  alongside (2026-10-12).
- Older pins: `git log -p --follow plugin/rdocx.lock.json`.

## Next steps

1. When upstream `main` moves (S96 brings fillable PDF forms, tensorbee/rdocx#381) or tags a release: `bump <ref>` on
   a branch, `install --build` and `test` (or let ci.yml build it on the pull request), as docs/setup.md (Moving the
   pin) describes; delete the gap tests that pass; release the plugin when the maintainer asks.
