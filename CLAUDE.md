# rdocx-skills: working notes for agents

A Claude plugin (`rdocx`) with two auto-loading skills, `docx` and `pptx`, that make agents use the rdocx /
rpptx CLIs and Python bindings for every Office file task, plus the acceptance suite that decides whether a
new rdocx build can be pinned. Public repository: nothing specific to a company or a project goes in here
(no document names, paths, people, or private file content; fixtures are synthetic).

## Map

- `skills/<name>/SKILL.md`: what an agent reads first; keep it under 16 KB and its description under 800 characters, details go to `references/`
  (`tests/test_skill_docs.py` checks the size, every Python name, keyword, CLI command and flag the skills cite, and
  the other way round that every public API of the build is cited or listed in its `NOT_DOCUMENTED` with a reason).
- `skills/<name>/references/`: `cli.md`, `python-api.md`, `recipes.md` (every block is executed by
  `tests/test_docs_snippets.py`), `gaps.md`, and `interop.md` for docx.
- `skills/<name>/scripts/`: helpers importable and runnable as commands; tested in `tests/test_skill_scripts.py`.
- `scripts/rdocx_env.py`: install (from a dist folder, else the releases), build, verify, bump, lock, test;
  standard library only, Python >= 3.9, Linux, macOS and Windows (Git Bash: `.exe` CLIs, `current` a junction).
- `rdocx.lock.json`: the pin (upstream commit and ref, release URLs, SHA-256 per platform). `dist/`
  (gitignored) holds local builds; downloads land in `$RDOCX_HOME/dist/`.
- `.github/workflows/`: `ci.yml` (the whole suite on the five platforms for a push to `main`, a run by hand and a PR
  that changes the lock; any other PR runs it on linux-x86_64 and the platform smoke subset, `-m smoke` from `SMOKE`
  in `tests/conftest.py`, on the four others; on a push to `main` whose lock pins one of this repository's releases,
  its summaries become the "Acceptance suite" section of that release's notes, through
  `.github/scripts/release_notes.py`), `build.yml` (only for a commit between two upstream releases, by hand: builds
  it on five runners, checks that it installs and runs, and publishes the release `rdocx-<YYYYMMDDHHMM>-<commit12>`,
  UTC build date and time first since GitHub lists releases by tag name (older releases: `rdocx-<YYYYMMDD>-<commit12>`,
  still valid); no suite there; inputs `lto` (`fat` upstream's profile, `thin` a faster build) and `dry_run` (build
  from scratch, publish nothing, to time a profile)). `.github/scripts/suite-summary.sh` groups a suite log (gaps
  closed, docs not yet updated, other failures); both scripts are tested in `tests/test_ci_scripts.py`.
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
- Pinned: the integration branch `integration/open-prs-2026-10-12-3` of hadim/rdocx (commit `12b1d6fc`) through
  this repository's release `rdocx-20261012-12b1d6fc85e3`: the previous pin `integration/open-prs-2026-10-12-2`
  (below) plus the fixes of its two gaps on the carried PRs: the HTML and Markdown importers collapse spaces on
  both sides of a span edge into one, kept where the run of spaces starts (`a <s>b </s> c` gives `a `, `b ` struck,
  `c`), also in links and `insert_html`, and no longer collapse `&nbsp;` (#359); a caller font file named after a
  typeface wins over the embedded face of that typeface even when its own family differs, and font directories load
  in file name order (#361). One integration fix adds a test only: in rpptx, a caller font named after a typeface
  wins over its embedded face.
- Open gaps: none. Closed by this pin: `html-space-doubled-at-span-edge`, `font-dir-alias-below-embedded-face`.
  `<p>&nbsp;</p>` is kept as a paragraph holding one non-breaking space on this build (tested).
- Suite against `12b1d6fc`: 814 passed, 9 skipped (Poppler absent), 0 xfailed on macos-arm64 from the release files.
- Before it, pinned: the integration branch `integration/open-prs-2026-10-12-2` of hadim/rdocx (commit `1ffda276`) through
  this repository's release `rdocx-20261012-1ffda27631c6`: the previous pin `integration/open-prs-2026-10-12`
  (below) plus the fixes of its three gaps on the carried PRs: the HTML importers of rdocx and rpptx keep a space
  before or after an emphasised span outside it and a space written inside it in it (#359); rpptx's own renders
  draw the faces a deck embeds, except MicroType Express ones, and a caller font wins over an embedded face of the
  same family in both bindings (#361); `import_slide(keep_source_formatting=True)` and `import_slides` read the
  source and the destination as they are in memory (#357). Two integration fixes add tests only: a caller font wins
  over an embedded face in rpptx renders; slide import keeps source formatting edited in memory.
- Before it, pinned: the integration branch `integration/open-prs-2026-10-12` of hadim/rdocx (commit `ffb8d3d6`)
  through this repository's release `rdocx-20261012-ffb8d3d67f42`: tensorbee/rdocx `v0.16.0` (`733ce6f7`) plus every open PR of
  the user, those of the previous pin (tensorbee/rdocx#315, #317 to #320, #322, #324 to #330, #333, #335, #336, #338,
  #341 to #343) and the newly carried #354 to #359 and #361, with two integration fixes (a group pushing nesting past
  the limit, #326 with #355; slide handles retired after font embedding edits). What changed since
  `integration/open-prs-2026-10-11`: new Word documents name `compatibilityMode` 15 and the Markdown export is
  CommonMark-safe (#354); a deck nested deeper than 16 levels raises `XmlError` at open and `add_group_shape` refuses a
  17th level (#355); `add_chart` from Python in both bindings, `Document.charts`, `shape.chart` and
  `Chart.replace_data` (#356); `import_slide` carries charts, embedded objects, comments and jumps,
  `keep_source_formatting`, `import_slides` and `rpptx slide import` (#357); `Presentation.to_markdown` / `to_html` and
  `rpptx convert --to md|html` (#358); `Document.from_html`, `open_html`, `insert_html`, `from_markdown`,
  `open_markdown` and `rdocx convert` from `.md` / `.html` with `--template` (#359); font embedding with an explicit
  `FontEmbeddingLicense` in both bindings and `rdocx font` / `rpptx font list|embed|remove|settings` (#361). The
  recipes and tests embed Aileron (CC0-1.0), the subset Pillow bundles (`conftest.free_font`). A skills `main` pinned
  on the fork steps back from upstream releases: go back to tensorbee/rdocx with `bump v<version>` once a release
  carries these PRs.
- Gaps found on `1ffda276`, closed by `12b1d6fc`: docx `html-space-doubled-at-span-edge` (two spaces kept where a
  browser draws one at a span edge, a regression of the #359 fix) and pptx `font-dir-alias-below-embedded-face` (a
  `font_dir` file only named after the typeface lost to the embedded face). Closed by `1ffda276`:
  `html-space-inside-emphasis`, `render-embedded-font`, `import-stale-source-formatting`, the three gaps of
  `ffb8d3d6`.
- Suite against `1ffda276`: 810 passed, 9 skipped (Poppler absent), 2 xfailed (the two open gaps) on macos-arm64
  from the release files.
- Suite against `ffb8d3d6`: 799 passed, 9 skipped (Poppler absent), 3 xfailed on macos-arm64 from the release files;
  the build notes give 750 to 751 passed on the five platforms (18 to 19 skipped), with 5 failures there, all in
  `test_skill_docs.py` (the new APIs not yet documented), fixed on this branch.
- Before it, pinned: the integration branch `integration/open-prs-2026-10-11` (commit `37c02616`, release
  `rdocx-20261011-37c0261661ad`), no gap open.
- Before it, pinned: the integration branch `integration/open-prs-2026-10-10-4` (commit `154dc6bd`, release
  `rdocx-20261010-154dc6bd1190`), with three docx gaps open.
- Before it, pinned: the upstream releases tensorbee/rdocx `v0.16.0` (rdocx 0.16.0) and `rpptx-v0.14.0` (rpptx
  0.14.0), both at commit `733ce6f7` (sprint S90).
- Earlier: the integration branch `integration/open-prs-2026-10-08-3`, and before it the upstream releases tensorbee/rdocx `v0.15.0` (rdocx 0.15.0) and `rpptx-v0.13.1` (rpptx 0.13.1),
  both at commit `9d019472` (sprint S88), the first ones that ship the CLIs and the wheels together under one
  tag per family, with `SHA256SUMS` and build provenance attestations (tensorbee/rdocx#266). The lock records
  five platforms: linux-x86_64, linux-aarch64 (CLIs need glibc 2.35, wheels manylinux_2_28), macos-arm64,
  macos-x86_64 (new: upstream builds it) and windows-x86_64. This repository's releases and the integration
  branches of hadim/rdocx are no longer used for the pin; `build.yml` stays for a commit between two
  upstream releases.
- The skills load ahead of Anthropic's docx and pptx skills: 8 of 8 headless sessions (summary, replace,
  new memo; add slide, outline, new deck) called `rdocx:docx` / `rdocx:pptx` first, with the
  document-skills plugin loaded alongside.

## Next steps

1. When rdocx tags a release: `bump v<version>`, `lock --write --release`, `install`, `test`, as
   docs/setup.md (Moving the pin) describes.
