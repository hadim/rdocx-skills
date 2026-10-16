#!/usr/bin/env python3
"""The plugin's releases (release.yml) and the titles of the builds (build.yml). Standard library only.

  plugin_release.py label --ref REF --rdocx V --rpptx V [--upstream OWNER/REPO]
      `rdocx V` and `rpptx V` joined by a middle dot, followed by ` + REF` unless REF is an upstream
      release tag, and by ` from OWNER/REPO` when that is not tensorbee/rdocx (a fork's integration branch)
  plugin_release.py version --date YYYY-MM-DD [TAG...]
      the CalVer version of a release made that day: `YYYY.M.D`, or `YYYY.M.D.N` (N from 2) when the tags
      (`v<version>`) already hold that day
  plugin_release.py title LOCK
      the label of the build the lock pins
  plugin_release.py notes --repo OWNER/REPO --version V --run-url URL LOCK
      the head of the release notes: what the release's own files hold, where they come from, the run that built
      and tested them, how to install (LOCK: the lock the release branch gets, which points at this release)
  plugin_release.py set-version PLUGIN_JSON V
      write `"version": "V"` into plugin.json, right after `name`
"""
import argparse
import datetime
import json
import re
import sys
from pathlib import Path

FAMILIES = ("rdocx", "rpptx")
UPSTREAM_TAG = re.compile(r"(?:rpptx-)?v\d+\.\d+\.\d+")  # upstream's release tags: v<version>, rpptx-v<version>
UPSTREAM = "tensorbee/rdocx"


def label(versions, ref, repo=UPSTREAM):
    """`rdocx 0.16.1` and `rpptx 0.14.1` joined by a middle dot for a build of an upstream release tag, followed
    by ` + <ref>` for anything else, such as an integration branch (a full commit hash shortened to 12 characters),
    and by ` from <repo>` when it comes from another repository than upstream's, such as the maintainer's fork."""
    missing = [f for f in FAMILIES if not versions.get(f)]
    if missing:
        sys.exit(f"no {' nor '.join(missing)} version")
    base = " \u00b7 ".join(f"{f} {versions[f]}" for f in FAMILIES)
    if not ref or UPSTREAM_TAG.fullmatch(ref):
        return base
    ref = ref[:12] if re.fullmatch(r"[0-9a-f]{40}", ref) else ref
    return f"{base} + {ref}" + (f" from {repo}" if repo and repo != UPSTREAM else "")


def repo_of(lock):
    """`owner/repo` of the lock's upstream URL."""
    return lock["upstream"].removeprefix("https://github.com/").rstrip("/")


def next_version(day, tags):
    """CalVer of a release made on `day`: `YYYY.M.D`, then `YYYY.M.D.2`, `.3`... for the next ones that day."""
    base = f"{day.year}.{day.month}.{day.day}"
    taken = {t[1:] for t in tags if t.startswith("v")}
    if base not in taken:
        return base
    n = 2
    while f"{base}.{n}" in taken:
        n += 1
    return f"{base}.{n}"


def notes(lock, repo, version, run_url):
    """The head of a release's notes; the acceptance suite of its files and GitHub's list of merged pull requests
    follow it. The release carries its own files: `lock` is the one written on the release branch, which points
    at them."""
    upstream = repo_of(lock)
    commit, ref = lock["commit"], lock.get("ref", "")
    platforms = sorted(lock.get("artifacts", {}))
    if not platforms:
        sys.exit("the lock records no files")
    return "\n".join([
        f"Plugin version `{version}`: the skills, the tests and the rdocx build move together.",
        "",
        "- **rdocx and rpptx**: " + " and ".join(f"{f} {lock['versions'][f]}" for f in FAMILIES)
        + f", built from [{upstream}@{commit[:12]}](https://github.com/{upstream}/tree/{commit})"
        + (f" (`{ref}`)" if ref and ref != commit else "") + ".",
        f"- **Files**: the CLIs and wheels for {', '.join(platforms[:-1])} and {platforms[-1]}" if len(platforms) > 1
        else f"- **Files**: the CLIs and wheels for {platforms[0]}",
        "  are attached to this release with their `SHA256SUMS`, and the plugin of this version installs them from",
        f"  here, each checked against its `rdocx.lock.json`. [This run]({run_url}) built them from the commit above",
        "  and ran the acceptance suite on them (below). Provenance: "
        f"`gh attestation verify FILE -R {repo}`.",
        f"- **Install**: `/plugin marketplace add {repo}`, then `/plugin install rdocx@rdocx-skills`.",
        "",
    ])


def set_version(text, version):
    """plugin.json's text with `"version": "<version>"` right after the `name` line (replacing any version)."""
    text = re.sub(r'\n[ \t]*"version":[^\n]*', "", text)
    out, n = re.subn(r'(\n([ \t]*)"name":[^\n]*,)', rf'\1\n\2"version": {json.dumps(version)},', text, count=1)
    if n != 1:
        sys.exit('no "name" line in plugin.json')
    json.loads(out)
    return out


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)
    p = sub.add_parser("label")
    p.add_argument("--ref", default="")
    p.add_argument("--rdocx", required=True)
    p.add_argument("--rpptx", required=True)
    p.add_argument("--upstream", default=UPSTREAM)
    p = sub.add_parser("version")
    p.add_argument("--date", required=True, type=datetime.date.fromisoformat)
    p.add_argument("tags", nargs="*")
    p = sub.add_parser("title")
    p.add_argument("lock", type=Path)
    p = sub.add_parser("notes")
    p.add_argument("--repo", required=True)
    p.add_argument("--version", required=True)
    p.add_argument("--run-url", required=True)
    p.add_argument("lock", type=Path)
    p = sub.add_parser("set-version")
    p.add_argument("manifest", type=Path)
    p.add_argument("version")
    args = ap.parse_args(argv)
    if args.cmd == "label":
        print(label({"rdocx": args.rdocx, "rpptx": args.rpptx}, args.ref, args.upstream))
    elif args.cmd == "version":
        print(next_version(args.date, args.tags))
    elif args.cmd == "title":
        lock = json.loads(args.lock.read_text())
        print(label(lock.get("versions", {}), lock.get("ref", ""), repo_of(lock)))
    elif args.cmd == "notes":
        sys.stdout.write(notes(json.loads(args.lock.read_text()), args.repo, args.version, args.run_url))
    else:
        args.manifest.write_bytes(set_version(args.manifest.read_text(), args.version).encode())


if __name__ == "__main__":
    main()
