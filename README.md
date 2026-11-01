<div align="center">

![rdocx-skills icon](plugin/assets/icon.svg)

# rdocx-skills

**Word and PowerPoint skills for AI agents: read, write, edit, review and render `.docx` and `.pptx` files
with native, hash-pinned tools.**

[![CI](https://github.com/hadim/rdocx-skills/actions/workflows/ci.yml/badge.svg)](https://github.com/hadim/rdocx-skills/actions/workflows/ci.yml)
[![Release](https://img.shields.io/github/v/release/hadim/rdocx-skills?label=release&color=4f46e5)](https://github.com/hadim/rdocx-skills/releases/latest)
[![License: MIT](https://img.shields.io/badge/license-MIT-22c55e)](https://github.com/hadim/rdocx-skills/blob/main/LICENSE)
[![Platforms](https://img.shields.io/badge/platforms-Linux%20%7C%20macOS%20%7C%20Windows-0ea5a4)](https://github.com/hadim/rdocx-skills/blob/main/docs/setup.md)
[![Built on rdocx](https://img.shields.io/badge/built%20on-rdocx-f97316)](https://github.com/tensorbee/rdocx)

</div>

| ![A redline of two versions of a lab report, rendered by rdocx](plugin/assets/screenshot-docx.png) | ![A slide with a table, a process and three figures, rendered by rpptx](plugin/assets/screenshot-pptx.png) |
|:---:|:---:|
| 📝 A redline of two versions of a lab report, rendered by rdocx in reviewer colours | 📊 A slide built and rendered by rpptx |

<p align="center"><sub>Both pictures come straight out of rdocx and rpptx: no Word, no PowerPoint, no LibreOffice.</sub></p>

Give your agent real Office skills. This plugin teaches it to work on Word and PowerPoint files through
[rdocx](https://github.com/tensorbee/rdocx) and rpptx, native Rust CLIs with Python bindings, instead of
LibreOffice round trips, python-docx, python-pptx or hand-written XML.

## ✨ Why rdocx-skills

Claude's built-in Word and PowerPoint skills create files with docx-js and pptxgenjs and change them by unzipping
the package and editing its XML by hand; python-docx and python-pptx stop where their API stops; and all of them
call LibreOffice to see a page. rdocx-skills gives your agent one native tool that reads, edits, reviews, renders
and checks Office files, so an everyday task is a call, not an XML project:

<!-- bench:value:start -->
| Your agent needs to... | Claude + rdocx-skills | Claude + built-in skill | Python libraries alone |
| --- | :--- | :--- | :--- |
| 🔀 Redline two versions of a Word file | `rdocx compare --granularity word`<br><sub>2 lines</sub> | a diff written by hand, emitted as `w:ins` / `w:del`<br><sub>86 lines · 🧬 raw XML</sub> | a diff written by hand, emitted as `w:ins` / `w:del`<br><sub>80 lines · 🧬 raw XML</sub> |
| ✅ Reject one reviewer's tracked changes | `rdocx revision reject --author`<br><sub>2 lines</sub> | the XML edited by hand<br><sub>42 lines · 🧬 raw XML</sub> | the XML edited by hand<br><sub>35 lines · 🧬 raw XML</sub> |
| 📑 Rebuild a stale table of contents | `rdocx toc rebuild`<br><sub>2 lines</sub> | LibreOffice for the page numbers, then the TOC rewritten in XML<br><sub>56 lines · 🧬 raw XML</sub> | ➖ no pagination |
| 💬 Comment on an exact phrase | `docx_ops.py comment`<br><sub>3 lines</sub> | `merge_runs.py`, `comment.py`, markers placed in the XML<br><sub>29 lines · 🧬 raw XML</sub> | `add_comment` on whole runs, runs split in XML<br><sub>30 lines · 🧬 raw XML</sub> |
| ✏️ Counted find and replace that keeps the formatting | `docx_ops.replace_batch`<br><sub>4 lines</sub> | `merge_runs.py`, an XML edit, zip<br><sub>41 lines · 🧬 raw XML</sub> | runs walked by hand<br><sub>38 lines</sub> |
| 📊 Insert a native chart in a Word file | `doc.add_chart(...)`<br><sub>13 lines</sub> | chart XML and its workbook written by hand<br><sub>83 lines · 🧬 raw XML</sub> | python-pptx's chart XML grafted with lxml<br><sub>36 lines · 🧬 raw XML</sub> |
| 🗂️ Bring slides over from another deck | `rpptx slide import`<br><sub>2 lines</sub> | every part, layout and relationship copied in XML<br><sub>113 lines · 🧬 raw XML</sub> | OPC parts and lxml<br><sub>48 lines · 🧬 raw XML</sub> |
| 🔁 Duplicate, move, delete and hide slides | `rpptx slide duplicate / move / remove / hide`<br><sub>13 lines</sub> | `add_slide.py`, `sldIdLst` edited, `clean.py`<br><sub>67 lines · 🧬 raw XML</sub> | deep copies and lxml (no API)<br><sub>30 lines · 🧬 raw XML</sub> |
| 🔢 Slide numbers and a footer on every slide | `rpptx footer`<br><sub>2 lines</sub> | placeholders copied in XML<br><sub>46 lines · 🧬 raw XML</sub> | placeholders cloned, the field in lxml<br><sub>22 lines · 🧬 raw XML</sub> |
| 📏 Find text that overflows its box | `rpptx fit`<br><sub>9 lines</sub> | render the deck and look<br><sub>80 lines · 🧬 raw XML · ⚠️ partly</sub> | an estimator written by hand<br><sub>71 lines · 🧬 raw XML · ⚠️ partly</sub> |
| 🖼️ See a page as a picture | `rdocx render --pages 1`<br><sub>3 lines</sub> | LibreOffice to PDF, then pdftoppm<br><sub>8 lines</sub> | ➖ no renderer |
<!-- bench:value:end -->

<sub>From the [benchmark](https://github.com/hadim/rdocx-skills/tree/main/benchmark): the code an agent writes for
each task with each stack, checked by the same neutral oracle. 🧬: OOXML written or edited by hand.</sub>

- 🛡️ **Hard to break a file**: every edit declares how many places it changes and applies all of them or none,
  most known mistakes are refused before anything is written, and a file is saved through a temporary copy,
  renamed only once complete, never over its input.
- ⚡ **Native and fast**: one CLI call for whole-file work, the Python binding for anything finer, no office suite
  to install.
- 🖨️ **Its own renderer**: PDF and PNG from rdocx's own layout engine, several times faster than a LibreOffice
  round trip; its fidelity to what Word draws is still catching up (see [How it compares](#-how-it-compares)).
- 🎯 **Loads by itself**: the skills step in as soon as a `.docx` or `.pptx` shows up, even when you don't
  name a tool.
- 🔒 **Verified supply chain**: every binary is pinned by SHA-256 and checked before it runs.
- 🧪 **Tested on four platforms**: an acceptance suite runs every recipe of the skills on each change.

## 🚀 Quick start

In Claude Code:

```
/plugin marketplace add hadim/rdocx-skills
/plugin install rdocx@rdocx-skills
```

Then ask in plain words, for example:

- 💬 *"Summarise the tracked changes in `paper-v2.docx` and accept the typo fixes."*
- 💬 *"Redline `report-v1.docx` against `report-v2.docx` and save the result as a PDF."*
- 💬 *"Fill `template.docx` with the rows of `data.csv`."*
- 💬 *"Turn these notes into a ten-slide deck and check that no text overflows its box."*
- 💬 *"Reply to every open comment in `review.pptx` and export the deck to Markdown."*

The first time a skill runs, it installs the pinned build for your machine in a few seconds (see
[What the plugin downloads and runs](#-what-the-plugin-downloads-and-runs)).

## 🧰 What your agent can do

rdocx and rpptx cover most of what Word and PowerPoint files hold. The skills point the agent at the right
call for each task.

### 📖 Read and extract

- 📄 Text of the body, tables, headers, footers, footnotes, text boxes and comments, also as JSON
- 🔎 Where a string occurs, the page each block lands on, the styles used, the metadata
- 🗂️ Slide text with its structure and formatting, titles and outline, speaker notes and comments

### ✍️ Create and edit

- 🆕 New documents from scratch, from Markdown or HTML, or from a `.dotx` or `.potx` template
- 🧩 Templates and forms filled from data: `{{ fields }}`, loops, content controls and legacy form fields
- 🔁 Counted find and replace, regular expressions included, across runs and without losing the formatting
- 🎨 Paragraphs, headings, lists, styles and numbering; tables with merges, borders and shading; page setup
- 🖼️ Pictures, hyperlinks, charts, equations from a subset of LaTeX, headers, footers and page numbers
- 🔷 Shapes, connectors, fills, lines, shadows, groups and z-order; themes, masters and layouts
- 🗃️ Slides added, reordered, hidden, imported or removed; decks merged; documents combined

### 💬 Review and collaborate

- 🗨️ Comments with replies and resolution, anchored on exact text or on a shape
- ✅ Tracked changes listed, accepted or rejected: all of them, or by id, author or date
- 🔀 A redline of two versions, word by word
- 📑 Tables of contents and page fields rebuilt; `rpptx diff` to see what changed between two decks
- 🔏 Documents and decks signed, and their signatures checked (certificate trust stays yours to judge)

### 🖨️ Render, convert and export

- 📄 PDF and PNG pages, with tracked changes shown in reviewer colours when asked; PDF/A, notes and handouts
- 🔄 Word to Markdown, HTML, ODT, RTF and EPUB; PowerPoint to Markdown and HTML
- 📏 Text-fit checks before a slide overflows, and validity checks before a file goes out
- 🔤 Fonts embedded, when their licence allows it

### 🛠️ When the API runs out

- 🧬 **Raw XML, safely**: paragraphs, runs, tables, shapes, text frames and slides expose their XML (`xml`,
  `replace_xml`), so an agent can still reach an Office feature that has no API yet, inside the same tool. The
  replacement is checked first, and a bad one changes nothing.
- 🩹 **Honest about gaps**: each skill lists the tool's known gaps with a workaround, or a hand-off to the
  built-in skills, and asks the agent to report any new one as a neutral reproduction.

## 📊 How it compares

<!-- bench:date:start -->
Run of 2026-11-01: rdocx 0.16.4 and rpptx 0.14.4 (pin `3707323312e9`) against the built-in skills (docx-js 9.15.0, pptxgenjs 4.0.1, LibreOffice 26.2.6.3, pandoc 3.13) and python-docx 1.2.0 / python-pptx 1.0.2.
<!-- bench:date:end -->

<!-- bench:scoreboard:start -->
|  | Claude + rdocx-skills | Claude + built-in skill | Python libraries alone |
| --- | :---: | :---: | :---: |
| 🧰 Tasks done, out of 44 | **43** | 41 | 37 |
| ✍️ Lines of code an agent writes (34 tasks all three do) | **388** | 1600 | 1190 |
| 🧬 Tasks that need hand-written OOXML | **1** | 29 | 29 |
| 🛑 Known mistakes refused before a file is written | **80 %** | 3 % | 39 % |
| 🔎 Corrupt or invalid docx caught by the stack's checks | 87 % | **89 %** | 36 % |
| 🔎 Corrupt or invalid pptx caught by the stack's checks | **97 %** | 73 % | 39 % |
| ⚡ Every page of a 52-page Word report to PNG | **2.13 s** | 7.38 s | ➖ |
| ⚡ Every slide of a 60-slide deck to PNG | **1.34 s** | 5.81 s | ➖ |
| ✏️ Counted replacement in a 52-page Word report | 3.07 s | 0.44 s | **0.19 s** |
| 🏗️ Build a 52-page Word report from scratch | 13 s | **0.29 s** | 0.48 s |
| 🖨️ Word files with the same page count as Word's own PDF | 4 / 8 | **7 / 8** | ➖ |
<!-- bench:scoreboard:end -->

<sub>Claude + rdocx-skills: this plugin. Claude + built-in skill: what Claude uses without it (docx-js and pptxgenjs
to create, the XML unzipped and edited to change a file, pandoc to read, LibreOffice to render). Python libraries
alone: python-docx and python-pptx, what an agent writes with no skill and, as far as third-party reviews describe it,
what ChatGPT runs in its code interpreter; fast, as everything runs in one Python process, but with no renderer, no
validator and no tracked changes. Best value of each row in bold; ➖: the stack has no way to do it. Where rdocx-skills
is behind, the cause is reported upstream and the benchmark is rerun on each new pin.</sub>

👉 [The full benchmark](https://github.com/hadim/rdocx-skills/tree/main/benchmark): the method, all 44 tasks
with the code of each stack, broken-file detection, speed on realistic files and pages compared with Office's own
PDFs.

## 🔒 What the plugin downloads and runs

The plugin itself holds no binaries. Here is everything it fetches and runs:

| | |
|---|---|
| ⬇️ **Downloads** | Once, the pinned rdocx and rpptx CLIs and wheels for your machine (about 45 MB), from this repository's GitHub release recorded in `rdocx.lock.json`. Nothing comes from PyPI. |
| ✅ **Verifies** | The SHA-256 of every file against the lock, before anything is installed; `status` checks them again. Without a verified build, the skills say so and hand the task to the built-in skills. |
| 📁 **Installs** | Under `~/.local/share/rdocx-skills/` (`RDOCX_HOME`), for your user only, nothing system-wide. |
| 🚫 **Sends** | Nothing. The tools run locally, and your documents never leave your machine. |

<details>
<summary>Two maintainer commands go further</summary>

The skills never run them on their own:

- `install --build` compiles the pinned commit from source: the commit from GitHub, crates from crates.io with
  `cargo --locked`, build tools from PyPI pinned by hash.
- `lock --write --release` downloads the files of the release the lock names and records their hashes.

</details>

### 🔐 Supply chain

- 📌 **The pin**: `rdocx.lock.json` names an upstream commit and, per platform, the SHA-256 of each CLI, each
  wheel and every file the wheels install.
- 📦 **The builds**: the plugin's releases are the only published builds. Each one builds the pinned commit, an
  upstream release as much as an integration branch, on four platforms, runs the acceptance suite on those files,
  attaches them with their `SHA256SUMS` and build provenance attestations, and its lock points at them. Its
  title says what it holds: `rdocx 0.16.1 · rpptx 0.14.1` for an upstream release, or `... + <ref>` for anything
  else. Older releases stay, each with its own files.
- 🍴 **Integration branches** live on the maintainer's fork, [hadim/rdocx](https://github.com/hadim/rdocx): an
  upstream release plus the maintainer's pull requests still open on tensorbee/rdocx, pinned when the skills need
  a fix before upstream releases it. The lock's `upstream` and `ref` name the fork and the branch.
- 🧾 **Local builds**: every git object of the commit is verified, and `cargo --locked` builds it.

```bash
python3 plugin/scripts/rdocx_env.py install      # download the pinned files and verify them (or: install --build)
python3 plugin/scripts/rdocx_env.py test         # the acceptance suite on the installed build
python3 plugin/scripts/rdocx_env.py bump v0.16.1 # move the pin to an upstream tag or commit (docs/setup.md)
```

### 💻 Platforms

- 🐧 Linux x86_64 and arm64, glibc 2.28 or later
- 🍎 macOS on Apple silicon (Intel Macs: `install --build`)
- 🪟 Windows x86_64, from Git Bash

See [docs/setup.md](https://github.com/hadim/rdocx-skills/blob/main/docs/setup.md) for the details.

## 📦 Also as a conda package

Each release of the plugin also publishes its build, the very files the suite tested, as the conda package
`rdocx-skills` on the prefix.dev channel [`hadim/rdocx-skills`](https://prefix.dev/channels/hadim/rdocx-skills):
both CLIs and both Python modules (`import rdocx`, `import rpptx`, Python 3.10 or later), for the platforms above.
Its version is the plugin's (`2026.10.25`), and the channel keeps the 10 latest. Use it with pixi, conda or mamba,
inside or outside an agent:

```bash
# the CLIs on your PATH, in an environment of their own (~/.pixi/envs/rdocx-skills)
pixi global install -c https://prefix.dev/hadim/rdocx-skills -c conda-forge rdocx-skills
# one command, nothing installed
pixi exec -c https://prefix.dev/hadim/rdocx-skills -c conda-forge -s rdocx-skills rdocx text report.docx
# in a pixi workspace
pixi workspace channel add https://prefix.dev/hadim/rdocx-skills
pixi add rdocx-skills
# conda or mamba
conda install -c https://prefix.dev/hadim/rdocx-skills -c conda-forge rdocx-skills
```

To have the skills use it rather than their own install (on Linux and macOS), set `R` to the environment's `bin`
folder where their setup sets it, and `RDOCX_BIN_DIR` for their helper scripts; `$R/rdocx`, `$R/rpptx` and
`$R/python` then work as after `rdocx_env.py install`:

```bash
R=~/.pixi/envs/rdocx-skills/bin   # pixi global; in a workspace: .pixi/envs/default/bin; conda: $CONDA_PREFIX/bin
export RDOCX_BIN_DIR=$R
```

On Windows the CLIs sit in the environment's `Library/bin` and Python at its root. conda-forge's `rdocx`,
`rdocx-cli`, `rpptx` and `rpptx-cli` package upstream's releases instead; `rdocx-skills` follows the plugin's pin,
which may carry fixes that no release has yet, and is the build the skills are tested with. Both hold the same
modules and commands, so an environment takes one or the other (the solver refuses both).

## 🧪 Acceptance suite

`tests/` checks, on every push and pull request, on Linux x86_64 and arm64, macOS arm64 and Windows x86_64:

- 📄 every vertical on synthetic fixtures generated from scratch: a survey report of about 13 pages with the
  traits of a Google Docs export edited in Word, and a 7-slide deck;
- 🧮 the two acceptance matrices, identity attributes and producer traits by operation;
- 🧷 a document with a token in every story: cells, content controls, tracked changes, text boxes, headers and
  footers, notes;
- 📚 every recipe of the skills, run as written, and the helper scripts;
- 🚧 each known gap as a strict expected failure (`tests/gaps.py`): when a new pin fixes it, the suite fails on
  purpose so that the skills are updated.

## 🔄 Releases and updates

- 🏷️ **Versions**: each release of the plugin is a CalVer version (`2026.10.15`) and a GitHub release titled with
  the rdocx and rpptx versions it pins. The skills, the tests and the pinned build move together.
- 🌿 **The `release` branch** is the repository's default branch and carries the released version: Claude Code,
  Claude Desktop and the Claude plugin directory follow it. Development happens on `main`, which moves ahead between
  releases; to try it, add the marketplace as `hadim/rdocx-skills#main`.
- 🔔 **Updates**: third-party marketplaces do not update by themselves. Turn on auto-update for `rdocx-skills` in
  `/plugin` (Marketplaces), or run `/plugin marketplace update rdocx-skills`.

## 🗂️ Repository layout

<details>
<summary>What lives where</summary>

| Path | Content |
|---|---|
| `.claude-plugin/` | the marketplace entry, pointing at `plugin/` |
| `plugin/` | the plugin, and only what it ships: |
| `plugin/.claude-plugin/` | the plugin manifest |
| `plugin/skills/docx/`, `plugin/skills/pptx/` | `SKILL.md`, `references/` (CLI, Python API, recipes, gaps, interoperability), `scripts/` (helpers usable as commands) |
| `plugin/scripts/` | `rdocx_env.py` (install, build, verify, test, bump) and the hash-pinned build requirements |
| `plugin/rdocx.lock.json` | the pinned build |
| `plugin/assets/` | the plugin icon and the pictures of this page |
| `.github/` | CI on every push (it builds the pinned commit), the release workflow of the plugin |
| `tests/` | acceptance suite, fixture generators, gap registry |
| `findings/` | gaps met in real use, as neutral reproductions, before they become tests and gap entries |
| `benchmark/` | the comparison with Claude's built-in skills and python-docx / python-pptx, rerun by hand |
| `docs/` | setup, releases and maintenance |

</details>

## 🛡️ Privacy

The plugin collects nothing and sends none of your data anywhere: the tools run locally, and its only network
request is the one-time download of the pinned build from GitHub. See the
[privacy policy](https://github.com/hadim/rdocx-skills/blob/main/PRIVACY.md).

## 📜 License

MIT. rdocx itself is MIT or Apache-2.0. This plugin is an independent project built on rdocx.
