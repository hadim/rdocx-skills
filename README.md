<div align="center">

![rdocx-skills icon](assets/icon.svg)

# rdocx-skills

**Word and PowerPoint skills for AI agents: read, write, edit, redline and render `.docx` and `.pptx` files
with native, hash-pinned tools.**

[![CI](https://github.com/hadim/rdocx-skills/actions/workflows/ci.yml/badge.svg)](https://github.com/hadim/rdocx-skills/actions/workflows/ci.yml)
[![License: MIT](https://img.shields.io/badge/license-MIT-22c55e)](LICENSE)
[![Platforms](https://img.shields.io/badge/platforms-Linux%20%7C%20macOS%20%7C%20Windows-4f46e5)](docs/setup.md)
[![Built on rdocx](https://img.shields.io/badge/built%20on-rdocx-0ea5a4)](https://github.com/tensorbee/rdocx)
[![Skills](https://img.shields.io/badge/skills-docx%20%C2%B7%20pptx-f97316)](skills/)

</div>

Give your agent real Office skills. This plugin teaches it to work on Word and PowerPoint files through
[rdocx](https://github.com/tensorbee/rdocx) and rpptx, native Rust CLIs with Python bindings, instead of
LibreOffice round trips, python-docx, python-pptx or hand-written XML.

## ✨ Why rdocx-skills

- ⚡ **Native and fast**: one CLI call for whole-file work, the Python binding for anything finer. No office
  suite to install, no headless LibreOffice.
- 🎯 **Loads by itself**: the skills step in as soon as a `.docx` or `.pptx` shows up, even when you don't
  name a tool.
- ✍️ **Safe edits**: counted find and replace across runs, text edited without losing its formatting, and
  the input never written over.
- 🔍 **Real tracked changes**: lists, accepts and rejects revisions (all, or by id, author or date), and
  builds a redline of two versions.
- 🔒 **Verified supply chain**: every binary is pinned by SHA-256 and checked before it runs.
- 🧪 **Tested on four platforms**: an acceptance suite runs every recipe of the skills on each change.
- 🩹 **Honest about gaps**: known limits are listed with a workaround, or a hand-off to the built-in skills.

## 🚀 Quick start

In Claude Code:

```
/plugin marketplace add hadim/rdocx-skills
/plugin install rdocx@rdocx-skills
```

Then ask in plain words, for example:

- 💬 *"Summarise the tracked changes in `contract-v2.docx` and accept the typo fixes."*
- 💬 *"Redline `report-v1.docx` against `report-v2.docx` and save the result as a PDF."*
- 💬 *"Fill `template.docx` with the rows of `data.csv`."*
- 💬 *"Turn these notes into a ten-slide deck and check that no text overflows its box."*
- 💬 *"Replace the old product name on every slide and render the slides to PNG."*

The first time a skill runs, it installs the pinned build for your machine in a few seconds (see
[What the plugin downloads](#-what-the-plugin-downloads-and-runs)).

## 🧰 What the skills cover

| Skill | What your agent can do |
|---|---|
| 📝 **`rdocx:docx`** (Word) | read text, tables, comments and tracked changes; structure and JSON views; create from scratch, Markdown or HTML; fill templates; find and replace; formatting, styles, tables, pictures, hyperlinks, charts; comments and threads; accept or reject revisions; redline two versions; tables of contents and fields; layout; convert to PDF, PNG, Markdown or HTML |
| 📊 **`rdocx:pptx`** (PowerPoint) | read slide text, outline, notes and comments; create decks and fill templates; edit text without losing formatting; counted find and replace; shapes and groups, pictures, tables, charts; add, reorder, hide, import or remove slides; merge decks; check that text fits; render to PDF or PNG; export to Markdown or HTML; embed fonts |

Each skill teaches the CLI for whole-file operations and the Python binding for finer work, lists the tool's
known gaps with their workaround, and asks the agent to report any new gap as a neutral reproduction.

## 🔒 What the plugin downloads and runs

The plugin itself holds no binaries. Here is everything it fetches and runs:

- **Downloads**: the first time a skill runs, `scripts/rdocx_env.py install` downloads the pinned rdocx and
  rpptx CLIs and wheels for your machine (about 45 MB) from the GitHub release recorded in `rdocx.lock.json`:
  rdocx's own releases, or this repository's releases for a commit between two upstream releases. Nothing
  comes from PyPI.
- **Verifies**: it installs only the files whose SHA-256 matches the lock, and `status` checks them again. If
  no verified build is available, the skills say so and hand the task to the built-in skills.
- **Installs**: under `~/.local/share/rdocx-skills/` (`RDOCX_HOME`), for your user only, nothing system-wide.
- **Sends**: nothing. The tools run locally, and your documents never leave your machine.

Two maintainer commands go further, and the skills never run them on their own: `install --build` compiles the
pinned commit from source (the commit from GitHub, crates from crates.io with `cargo --locked`, build tools
from PyPI pinned by hash), and `lock --write --release` lists this repository's releases through the GitHub
API, with `GH_TOKEN` or `GITHUB_TOKEN` when one is set.

### Supply chain

`rdocx.lock.json` pins an upstream commit and, per platform, the SHA-256 of each CLI, each wheel and every
file the wheels install. The files come from the upstream releases of rdocx and rpptx (one tag per family,
CLI archives and wheels built by rdocx's own workflow), recorded in the lock after their checksums and their
build provenance (`gh attestation verify`: built from the pinned commit by a GitHub-hosted runner of
tensorbee/rdocx) were checked. A commit between two upstream releases can still be pinned: this repository's
build workflow then compiles it on four platforms (not Intel Macs) and publishes it as the release
`rdocx-<YYYYMMDDHHMM>-<commit12>` (UTC build time: GitHub lists releases by tag name). A local build verifies
every git object of the commit and uses `cargo --locked`. See [docs/setup.md](docs/setup.md).

```bash
python3 scripts/rdocx_env.py install      # download the pinned files and verify them (or: install --build)
python3 scripts/rdocx_env.py test         # the acceptance suite on the installed build
python3 scripts/rdocx_env.py bump v0.15.0 # move the pin to an upstream tag or commit (docs/setup.md)
```

**Platforms**: Linux x86_64 and arm64 (glibc 2.35 or later), macOS arm64, Windows x86_64 from Git Bash.
Intel Macs get ready-made files when the pin is an upstream release, and `install --build` otherwise.

## 🧪 Acceptance suite

`tests/` checks every vertical on synthetic fixtures generated from scratch (a survey report of about 13 pages
with the serialisation traits of a Google Docs export edited in Word, a 7-slide deck), the two acceptance
matrices (identity attributes and producer traits by operation), a document with a token in every story
(cells, content controls, tracked changes, text boxes, headers and footers, notes), the helper scripts and
every recipe of the skills. Each known gap is a strict expected failure (`tests/gaps.py`): when a new pin
fixes it, the suite fails on purpose so that the skills are updated. CI runs it on Linux x86_64 and arm64,
macOS arm64 and Windows x86_64 for every push and pull request.

## 🔄 Updates

The plugin has no version number: each commit of `main` is a version, with the skills, the tests and the
pinned rdocx build moving together. Third-party marketplaces do not update by themselves: turn on
auto-update for `rdocx-skills` in `/plugin` (Marketplaces), or run `/plugin marketplace update rdocx-skills`.

## 🗂️ Repository layout

| Path | Content |
|---|---|
| `.claude-plugin/` | plugin manifest and marketplace entry |
| `skills/docx/`, `skills/pptx/` | `SKILL.md`, `references/` (CLI, Python API, recipes, gaps, interoperability), `scripts/` (helpers usable as commands) |
| `scripts/` | `rdocx_env.py` (install, build, verify, test, bump) and the hash-pinned build requirements |
| `.github/` | CI on every push, and the build workflow for a commit between two upstream releases |
| `rdocx.lock.json` | the pinned build |
| `tests/` | acceptance suite, fixture generators, gap registry |
| `findings/` | gaps met in real use, as neutral reproductions, before they become tests and gap entries |
| `assets/` | the plugin icon |
| `docs/` | setup and maintenance |

## 📜 License

MIT. rdocx itself is MIT or Apache-2.0. This plugin is an independent project built on rdocx.
