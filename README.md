# rdocx-skills

Agent skills for Word (.docx) and PowerPoint (.pptx) files built on [rdocx](https://github.com/tensorbee/rdocx):
native Rust CLIs and Python bindings that read, write, edit, compare and render Office files without
LibreOffice, python-docx, python-pptx or XML surgery.

- **`rdocx:docx`**: reading, structure and JSON views, formatting, tables, pictures, hyperlinks, comments
  and threads, tracked changes, redlines, tables of contents and fields, layout, PDF and PNG rendering.
- **`rdocx:pptx`**: slide text and outline, shapes and groups, pictures, tables, notes, comments, counted
  replacement, slide order and visibility, text-fit checks, PDF and PNG rendering.

Each skill loads by itself when an agent meets a .docx or .pptx task, teaches the CLI for whole-file
operations and the Python binding for anything finer, lists the tool's known gaps with their workaround or
their fallback to the default skills, and asks the agent to report every new gap as a neutral reproduction.

## Supply chain

The skills never run an rdocx build they cannot verify. `rdocx.lock.json` pins an upstream commit and, per
platform, the SHA-256 of each CLI, each wheel and every file the wheels install. The files come from this
repository's releases: a bump of the lock makes the build workflow compile the commit on Linux x86_64, Linux
arm64 and macOS arm64, run the acceptance suite on each, and publish them. `scripts/rdocx_env.py` downloads
them and installs only files that match the lock, re-checks them on `status`, or builds the pinned commit
itself (every git object of the commit verified, `cargo --locked`, build tools pinned by hash). If no
verified build is available, the skills say so and hand the task to the built-in skills. See
[docs/setup.md](docs/setup.md).

```bash
python3 scripts/rdocx_env.py install      # download the pinned files and verify them (or: install --build)
python3 scripts/rdocx_env.py test         # the acceptance suite on the installed build
python3 scripts/rdocx_env.py bump v0.15.0 # move the pin to an upstream tag or commit (docs/setup.md)
```

## Acceptance suite

`tests/` checks every vertical on synthetic fixtures generated from scratch (a survey report of about 13 pages with
the serialisation traits of a Google Docs export edited in Word, a 7-slide deck), the two acceptance
matrices (identity attributes and producer traits by operation), a document with a token in every story
(cells, content controls, tracked changes, text boxes, headers and footers, notes), the helper scripts and
every recipe of the skills. Each known gap is a strict expected failure tied to its upstream ticket (`tests/gaps.py`): when a
new pin fixes it, the suite fails on purpose so that the skills are updated. Run it before moving the pin.

## Layout

| Path | Content |
|---|---|
| `.claude-plugin/` | plugin manifest and marketplace entry |
| `skills/docx/`, `skills/pptx/` | `SKILL.md`, `references/` (CLI, Python API, recipes, gaps, interoperability), `scripts/` (helpers usable as commands) |
| `scripts/` | `rdocx_env.py` (install, build, verify, test, bump) and the hash-pinned build requirements |
| `.github/` | CI on every push, and the build workflow that publishes the releases |
| `rdocx.lock.json` | the pinned build |
| `tests/` | acceptance suite, fixture generators, gap registry |
| `findings/` | gaps met in real use, as neutral reproductions, before they become tests and upstream tickets |
| `docs/` | setup and maintenance |

## Install as a plugin

```
/plugin marketplace add hadim/rdocx-skills
/plugin install rdocx@rdocx-skills
```

License: MIT. rdocx itself is MIT or Apache-2.0.
