# Privacy

rdocx-skills is a Claude plugin: two skills and the scripts they call. It has no server, no account and no
telemetry.

## What the plugin collects

Nothing. It reads the Word and PowerPoint files you point it at, on your machine, and writes the files you ask
for there too. It keeps no copy of your documents and records nothing about you or your use of it.

## What leaves your machine

- **Nothing from your documents.** The rdocx and rpptx tools run locally and make no network request.
- **One download, once per pinned version.** The first time a skill runs, `scripts/rdocx_env.py install` downloads
  the pinned rdocx and rpptx files for your platform from this repository's GitHub releases and checks their
  SHA-256. That request reaches GitHub like any download from github.com and carries none of your data. GitHub's
  privacy statement applies to it.
- **Your conversation with Claude.** When Claude reads a document through these skills, the text it reads becomes
  part of your conversation, which the Claude product you use handles under its own terms and privacy policy.
  The plugin adds no other destination.

## What stays on your machine

The installed tools live under `~/.local/share/rdocx-skills/` (or `RDOCX_HOME`). Deleting that folder removes them.

## Contact

Questions or concerns: open an issue at https://github.com/hadim/rdocx-skills/issues.
