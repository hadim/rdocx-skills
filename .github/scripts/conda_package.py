#!/usr/bin/env python3
"""The conda package `rdocx-skills` of a plugin release (release.yml, job `conda`). Standard library only.

The package holds a release's own files for one platform, the two CLIs and the two wheels that the suite tested,
repackaged by conda/recipe.yaml (`pixi publish`), never built again; the prefix.dev channel keeps the latest ones.

  conda_package.py stage --release vVERSION --platform PLATFORM --out DIR [--repo OWNER/REPO]
      the release's files for PLATFORM into DIR, an empty or new folder (conda/dist, where the recipe reads them):
      their SHA-256 checked against the lock of the release tag, the `<platform>.` prefix of the asset names
      dropped; upstream's license files at the pinned commit into DIR/licenses; then the environment
      conda/recipe.yaml reads, as KEY=VALUE lines on stdout
  conda_package.py prune --channel URL --keep N [--dry-run]
      delete from the prefix.dev channel URL every package file of a version older than its N latest, with a
      token minted from the GitHub Actions OIDC token (the channel's trusted publisher), or PREFIX_DEV_TOKEN
"""
import argparse
import hashlib
import json
import os
import re
import sys
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path

REPO = "hadim/rdocx-skills"
# the plugin's platforms and their conda subdirs (no macos-x86_64: not built)
SUBDIRS = {"linux-x86_64": "linux-64", "linux-aarch64": "linux-aarch64", "macos-arm64": "osx-arm64",
           "windows-x86_64": "win-64"}
# upstream's license files, as its conda-forge feedstocks ship them: the code, the bundled fonts, the ICC profile
LICENSE_FILES = ("LICENSE",)
LICENSE_DIRS = {"crates/oxml-layout/fonts": re.compile(r"(LICENSE|NOTICE)-.+"),
                "crates/oxml-pdf/assets": re.compile(r"LICENSE-.+")}
PREFIX_DEV = "https://prefix.dev"


def fetch(url, headers=None, data=None, method=None):
    req = urllib.request.Request(url, data=data, method=method, headers={"User-Agent": "rdocx-skills", **(headers or {})})
    with urllib.request.urlopen(req, timeout=120) as res:
        return res.read()


def github_headers():
    token = os.environ.get("GH_TOKEN") or os.environ.get("GITHUB_TOKEN")
    return {"Authorization": f"Bearer {token}"} if token else {}


def release_lock(repo, tag):
    """The lock of the release branch at the release's tag: its files' hashes, the build's commit and versions."""
    return json.loads(fetch(f"https://raw.githubusercontent.com/{repo}/{tag}/plugin/rdocx.lock.json"))


def recipe_env(lock, tag):
    """What conda/recipe.yaml reads: the package version is the plugin's, the description names the build."""
    versions = lock.get("versions") or {}
    if not versions.get("rdocx") or not versions.get("rpptx"):
        sys.exit(f"the lock of {tag} records no rdocx or rpptx version")
    return {"RDOCX_SKILLS_VERSION": tag.removeprefix("v"),
            "RDOCX_VERSION": versions["rdocx"], "RPPTX_VERSION": versions["rpptx"],
            "RDOCX_COMMIT": lock["commit"], "RDOCX_REF": lock.get("ref") or lock["commit"],
            "RDOCX_UPSTREAM": lock["upstream"]}


def stage(repo, tag, platform, out):
    lock = release_lock(repo, tag)
    files = (lock.get("artifacts") or {}).get(platform)
    if not files or not lock.get("release"):
        sys.exit(f"{tag} has no files for {platform}")
    if out.exists() and any(out.iterdir()):
        sys.exit(f"{out} is not empty")
    out.mkdir(parents=True, exist_ok=True)
    for name, digest in sorted(files.items()):
        data = fetch(f"{lock['release']}/{platform}.{name}")
        if hashlib.sha256(data).hexdigest() != digest:
            sys.exit(f"{platform}.{name}: SHA-256 differs from the lock of {tag}")
        (out / name).write_bytes(data)
        print(f"{name}: {len(data)} bytes, SHA-256 checked", file=sys.stderr)
    upstream = lock["upstream"].removeprefix("https://github.com/").rstrip("/")
    licenses = out / "licenses"
    licenses.mkdir(exist_ok=True)
    paths = list(LICENSE_FILES)
    for folder, pattern in LICENSE_DIRS.items():
        listing = json.loads(fetch(f"https://api.github.com/repos/{upstream}/contents/{folder}?ref={lock['commit']}",
                                   github_headers()))
        paths += [f"{folder}/{e['name']}" for e in listing if pattern.fullmatch(e["name"])]
    for path in paths:
        (licenses / Path(path).name).write_bytes(
            fetch(f"https://raw.githubusercontent.com/{upstream}/{lock['commit']}/{path}"))
    for key, value in recipe_env(lock, tag).items():
        print(f"{key}={value}")


def version_key(version):
    """CalVer `YYYY.M.D[.N]` as a sortable tuple."""
    return tuple(int(p) for p in version.split("."))


def to_delete(packages, keep):
    """The package files (subdir, file name) of every version older than the `keep` latest, across subdirs."""
    versions = sorted({v for _, _, v in packages}, key=version_key, reverse=True)
    old = set(versions[keep:])
    return sorted((subdir, name) for subdir, name, v in packages if v in old)


def channel_packages(channel):
    """(subdir, file name, version) of every package file in the channel, from its repodata."""
    packages = []
    for subdir in sorted(set(SUBDIRS.values())):
        try:
            repodata = json.loads(fetch(f"{channel}/{subdir}/repodata.json"))
        except urllib.error.HTTPError as e:
            if e.code == 404:
                continue
            raise
        for key in ("packages", "packages.conda"):
            for name, record in (repodata.get(key) or {}).items():
                packages.append((subdir, name, record["version"]))
    return packages


def prefix_token():
    """A prefix.dev token: PREFIX_DEV_TOKEN, else one minted from this GitHub Actions job's OIDC token."""
    if os.environ.get("PREFIX_DEV_TOKEN"):
        return os.environ["PREFIX_DEV_TOKEN"]
    url, bearer = os.environ.get("ACTIONS_ID_TOKEN_REQUEST_URL"), os.environ.get("ACTIONS_ID_TOKEN_REQUEST_TOKEN")
    if not url or not bearer:
        sys.exit("no PREFIX_DEV_TOKEN and no GitHub Actions OIDC token (permissions: id-token: write)")
    oidc = json.loads(fetch(f"{url}&audience=prefix.dev", {"Authorization": f"Bearer {bearer}"}))["value"]
    return fetch(f"{PREFIX_DEV}/api/oidc/mint_token", {"Content-Type": "application/json"},
                 json.dumps({"token": oidc}).encode(), "POST").decode().strip()


def prune(channel, keep, dry_run):
    channel = channel.rstrip("/")
    name = channel.removeprefix(f"{PREFIX_DEV}/")
    doomed = to_delete(channel_packages(channel), keep)
    if not doomed:
        print(f"{channel}: {keep} versions or fewer, nothing to delete")
        return
    token = None if dry_run else prefix_token()
    for subdir, file in doomed:
        print(f"{'would delete' if dry_run else 'delete'} {subdir}/{file}")
        if not dry_run:
            fetch(f"{PREFIX_DEV}/api/v1/delete/{name}/{subdir}/{urllib.parse.quote(file)}",
                  {"Authorization": f"Bearer {token}"}, method="DELETE")


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = parser.add_subparsers(dest="command", required=True)
    p = sub.add_parser("stage")
    p.add_argument("--release", required=True)
    p.add_argument("--platform", required=True, choices=sorted(SUBDIRS))
    p.add_argument("--out", required=True, type=Path)
    p.add_argument("--repo", default=REPO)
    p = sub.add_parser("prune")
    p.add_argument("--channel", required=True)
    p.add_argument("--keep", required=True, type=int)
    p.add_argument("--dry-run", action="store_true")
    args = parser.parse_args()
    if args.command == "stage":
        stage(args.repo, args.release, args.platform, args.out.resolve())
    else:
        prune(args.channel, args.keep, args.dry_run)


if __name__ == "__main__":
    main()
