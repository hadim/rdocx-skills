#!/usr/bin/env python3
"""The final acceptance suite in the notes of the release the lock pins, for ci.yml on a push to main.

`build.yml` runs the suite with the skills tests of the commit it builds from, before the skills and tests catch
up with the new build, so its results are preliminary. Once `main` pins the release, ci.yml runs the suite of
`main` on it and writes those results at the head of the release notes, as the section
`## Acceptance suite (skills main <sha12>, <run url>)`, right after the notes' preamble. The build-time results
stay below it, under `## Build-time suite (preliminary, skills <sha12>)` (an older release titled them
`## Acceptance suite`: renamed). Running it again replaces the final section, it never adds a second one.

    release_notes.py tag [--repo OWNER/REPO] [LOCK]
        the tag of the lock's release when it is one of this repository's own builds
        (`rdocx-<YYYYMMDD>-<commit12>`), else nothing (an upstream release is never touched)
    release_notes.py rewrite --notes NOTES --sha SHA --run-url URL SUMMARY...
        NOTES with the final section made of the SUMMARY files (suite-summary.sh output), on stdout

Standard library only, Python >= 3.9.
"""
import argparse
import json
import re
import sys
from pathlib import Path

FINAL_PREFIX = "## Acceptance suite (skills main "
LEGACY_HEADING = "## Acceptance suite"
BUILD_PREFIX = "## Build-time suite"


def own_release_tag(release, repo, commit):
    """The tag of `release` (the lock's `release` field) when it is a build of this repository `repo` of the
    pinned `commit`, else None: upstream releases (a dict of family URLs), another repository's, another tag."""
    if not isinstance(release, str):
        return None
    m = re.fullmatch(rf"https://github\.com/{re.escape(repo)}/releases/download/(rdocx-\d{{8}}-([0-9a-f]{{12}}))/?",
                     release)
    if m is None or m.group(2) != commit[:12]:
        return None
    return m.group(1)


def build_heading(sha):
    return f"{BUILD_PREFIX} (preliminary, skills {sha[:12]})" if sha else f"{BUILD_PREFIX} (preliminary)"


def sections(body):
    """The preamble, then one list of lines per `## ` section (the platform summaries are `###`)."""
    out = [[]]
    for line in body.replace("\r\n", "\n").split("\n"):
        if line.startswith("## "):
            out.append([])
        out[-1].append(line)
    return out


def rewrite(body, summaries, sha, run_url):
    """`body` with the final section built from `summaries` (Markdown, one per platform) right after its preamble,
    and the build-time results kept below it under their own heading."""
    final = "\n\n".join(s.strip("\n") for s in summaries)
    if any(line.startswith("## ") for line in final.split("\n")):
        raise ValueError("a suite summary must not hold a `## ` heading")
    preamble, *rest = sections(body)
    built = re.search(r" from \S+@([0-9a-f]{40})\.", "\n".join(preamble))
    kept = []
    for sec in rest:
        if sec[0].startswith(FINAL_PREFIX):
            continue  # the previous final section, replaced
        if sec[0].rstrip() == LEGACY_HEADING:
            sec = [build_heading(built.group(1) if built else "")] + sec[1:]
        kept.append(sec)
    head = f"{FINAL_PREFIX}{sha[:12]}, {run_url})"
    parts = ["\n".join(preamble).strip("\n"), f"{head}\n\n{final}" if final else head]
    parts += ["\n".join(sec).strip("\n") for sec in kept]
    return "\n\n".join(p for p in parts if p) + "\n"


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)
    p = sub.add_parser("tag")
    p.add_argument("--repo", default="hadim/rdocx-skills")
    p.add_argument("lock", nargs="?", default="rdocx.lock.json")
    p = sub.add_parser("rewrite")
    p.add_argument("--notes", required=True)
    p.add_argument("--sha", required=True)
    p.add_argument("--run-url", required=True)
    p.add_argument("summaries", nargs="*")
    args = ap.parse_args(argv)
    if args.cmd == "tag":
        lock = json.loads(Path(args.lock).read_text(encoding="utf-8"))
        tag = own_release_tag(lock.get("release"), args.repo, lock.get("commit", ""))
        if tag is None:
            print(f"the lock's release is not a build of {args.repo}: {lock.get('release')!r}", file=sys.stderr)
        else:
            print(tag)
        return 0
    body = Path(args.notes).read_text(encoding="utf-8")
    summaries = [Path(s).read_text(encoding="utf-8") for s in sorted(args.summaries)]
    sys.stdout.write(rewrite(body, summaries, args.sha, args.run_url))
    return 0


if __name__ == "__main__":
    sys.exit(main())
