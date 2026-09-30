#!/usr/bin/env python3
"""Build, install and verify the pinned rdocx / rpptx toolchain (CLIs and Python bindings).

Standard library only, Python >= 3.9, Linux and macOS. The pin lives in ../rdocx.lock.json: an upstream git
commit and, per platform, the SHA-256 of the two CLIs, of the two wheels, and of every file the wheels install.

  rdocx_env.py status [--allow-local]   exit 0 when the installed build matches the lock, or is an
                                        unchanged local build of the pinned commit (only where the lock has
                                        no hashes for this platform, or with --allow-local), and `current`
                                        points at it
  rdocx_env.py install [--allow-local]  install from a dist folder whose files match the lock, else download
                                        them from the release named in the lock and check them the same way
  rdocx_env.py install --build          build the pinned commit from source if neither is available
  rdocx_env.py install --from DIR       install from DIR (for example files staged from another machine)
  rdocx_env.py build [--target linux-aarch64]
                                        build the pinned commit into dist/<commit>/<platform>/, print SHA-256
  rdocx_env.py lock [--write] [--platform P]
                                        compare (or record) a platform's dist hashes with the lock
  rdocx_env.py lock --write --release [--platform P]
                                        download the release's files, hash them here, record every platform
                                        (when the lock names no release yet, find the pinned commit's one
                                        among this repository's releases and record its URL first)
  rdocx_env.py bump REF                 pin an upstream tag, branch or full commit: new commit, release and
                                        hashes emptied until the release is built and recorded
  rdocx_env.py paths                    print RDOCX=..., RPPTX=..., RDOCX_PY=... for eval in a shell
  rdocx_env.py test [PYTEST ARGS]       run tests/ on the installed build, from a separate environment that
                                        holds the hash-pinned test dependencies

Trust model. Source: the commit named by its full hash is exported object by object with `git cat-file`, and
the SHA-1 of every object (commit, trees, blobs) is recomputed and compared with its id, with replace refs
disabled, so neither a replace ref, nor attributes, nor a substituted object can change what is built. It is
built with `cargo --locked` and maturin pinned by hash. Prebuilt files: each is copied once into a private
staging folder, hashed there, compared with the lock, and installed from that copy. Files built in the same
run are trusted by construction and installed as a local build, accepted only where the lock has no hashes
for the platform or with --allow-local; `lock --write` records them after review. `status` re-hashes the
CLIs, the wrapper, every file of the two packages (and refuses any other file there), the interpreter
link, pyvenv.cfg and the start-up files (.pth, sitecustomize) recorded at install.

Releases: the build workflow of this repository publishes, per upstream commit, a release `rdocx-<UTC build
date YYYYMMDD>-<first 12 characters of the commit>` whose assets are `<platform>.<file>` plus a SHA256SUMS
listing them. The tag cannot be derived from the commit alone: the lock's `release` URL is the only source of
truth, filled by `lock --write --release` from the GitHub API listing of the releases (GH_TOKEN or GITHUB_TOKEN
is sent when set). A downloaded file is trusted only through the lock: it lands in
RDOCX_HOME/dist/<commit>/<platform>/ and is installed by the same staging and hash check as any dist folder.

Environment: RDOCX_HOME (default ~/.local/share/rdocx-skills), RDOCX_SRC (a local clone of rdocx; default:
a sibling `rdocx` folder of this repository), RDOCX_DIST (a folder holding <commit>/<platform>/),
RDOCX_RELEASE_URL (a mirror of the release assets, instead of the lock's `release`), CARGO_BUILD_JOBS (fewer
parallel jobs on small machines).
"""
import argparse
import gzip
import hashlib
import json
import os
import platform
import re
import shlex
import shutil
import subprocess
import sys
import tempfile
import urllib.request
import zipfile
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
LOCK_PATH = REPO / "rdocx.lock.json"
HOME = Path(os.environ.get("RDOCX_HOME", Path.home() / ".local" / "share" / "rdocx-skills"))
GIT_ENV = dict(os.environ, GIT_NO_REPLACE_OBJECTS="1")
RELEASES_REPO = "hadim/rdocx-skills"  # the GitHub repository whose releases the lock names
TRIPLES = {"linux-x86_64": "x86_64-unknown-linux-gnu", "linux-aarch64": "aarch64-unknown-linux-gnu"}
ZIGBUILD_VERSION = "0.20.1"


def say(msg):
    print(msg, file=sys.stderr)


def die(msg, code=1):
    say(f"error: {msg}")
    sys.exit(code)


def load_lock():
    lock = json.loads(LOCK_PATH.read_text())
    if not re.fullmatch(r"[0-9a-f]{40}", lock.get("commit", "")):
        die("the lock's commit must be a full 40-character hexadecimal hash")
    return lock


def platform_key():
    system = {"darwin": "macos"}.get(platform.system().lower(), platform.system().lower())
    machine = platform.machine().lower()
    machine = {"amd64": "x86_64", "arm64": "arm64" if system == "macos" else "aarch64"}.get(machine, machine)
    return f"{system}-{machine}"


def sha256(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for block in iter(lambda: f.read(1 << 20), b""):
            h.update(block)
    return h.hexdigest()


def run(cmd, **kw):
    say("+ " + " ".join(str(c) for c in cmd))
    subprocess.run([str(c) for c in cmd], check=True, **kw)


def install_dir(lock):
    return HOME / lock["commit"][:12]


def expected(lock, plat):
    return lock.get("artifacts", {}).get(plat, {})


def expected_files(lock, plat):
    return lock.get("installed_files", {}).get(plat, {})


def artifact_names(folder):
    """Names of the artifacts in a dist folder; a gzipped CLI (`rdocx.gz`, for transfers that limit file
    sizes) counts under its plain name."""
    names = set()
    for p in Path(folder).iterdir():
        name = p.name[:-3] if p.name.endswith(".gz") else p.name
        if name in ("rdocx", "rpptx") or name.endswith(".whl"):
            names.add(name)
    return sorted(names)


def copy_artifact(folder, name, dest):
    """Copy `name` (or decompress `name.gz`) from `folder` to `dest`, which is what gets hashed and used."""
    plain, packed = Path(folder) / name, Path(folder) / f"{name}.gz"
    if plain.is_file():
        shutil.copyfile(plain, dest)
    elif packed.is_file():
        with gzip.open(packed, "rb") as src, open(dest, "wb") as out:
            shutil.copyfileobj(src, out)
    else:
        return False
    return True


def stage(folder, names):
    """Copy the artifacts into a private folder once. Returns (folder, {name: sha256}) for what was found."""
    tmp = Path(tempfile.mkdtemp(prefix="rdocx-stage-"))
    sums = {}
    for name in names:
        if copy_artifact(folder, name, tmp / name):
            sums[name] = sha256(tmp / name)
    return tmp, sums


def check(sums, want):
    report, ok = [], bool(want)
    for name, digest in sorted(want.items()):
        got = sums.get(name)
        report.append(f"{'missing ' if got is None else 'ok      ' if got == digest else 'MISMATCH'} {name} {got or ''}")
        ok &= got == digest
    return ok, report


def dist_candidates(lock, plat):
    rel = Path(lock["commit"]) / plat
    for root in (os.environ.get("RDOCX_DIST"), REPO / "dist", HOME / "dist"):
        if root and (Path(root) / rel).is_dir():
            yield Path(root) / rel


def source_candidates():
    for root in (os.environ.get("RDOCX_SRC"), REPO.parent / "rdocx"):
        if root and (Path(root) / ".git").exists():
            yield Path(root)


def write_lock(lock):
    LOCK_PATH.write_text(json.dumps(lock, indent=2) + "\n")


def record(lock, plat, folder, sums):
    """Record the artifacts of `plat` (staged in `folder`, with their `sums`), every file their wheels install,
    and the package versions read from the wheel names."""
    files = {}
    for name in sums:
        if name.endswith(".whl"):
            files.update(wheel_files(Path(folder) / name))
            package, version = name.split("-")[:2]
            lock.setdefault("versions", {})[package] = version
    lock.setdefault("artifacts", {})[plat] = sums
    lock.setdefault("installed_files", {})[plat] = files
    say(f"recorded {len(sums)} artifacts and {len(files)} installed files for {plat}")


def wheel_files(wheel):
    """{path installed by the wheel: sha256} for every file of the wheel outside its .dist-info folder."""
    with zipfile.ZipFile(wheel) as z:
        return {n: hashlib.sha256(z.read(n)).hexdigest() for n in z.namelist()
                if not n.endswith("/") and ".dist-info/" not in n}


# ---------------------------------------------------------------- releases
def github_releases(repo):
    """Every release of `repo` from the GitHub REST API, page by page (a token from GH_TOKEN or GITHUB_TOKEN,
    if set, raises the rate limit)."""
    token = os.environ.get("GH_TOKEN") or os.environ.get("GITHUB_TOKEN")
    releases, page = [], 1
    while True:
        req = urllib.request.Request(f"https://api.github.com/repos/{repo}/releases?per_page=100&page={page}",
                                     headers={"Accept": "application/vnd.github+json"})
        if token:
            req.add_header("Authorization", f"Bearer {token}")
        with urllib.request.urlopen(req, timeout=60) as res:
            batch = json.load(res)
        releases += batch
        if len(batch) < 100:
            return releases
        page += 1


def release_url(repo, releases, commit):
    """Download URL of the one release of `commit`: its tag ends with `-<first 12 characters of the commit>`."""
    tags = sorted(r["tag_name"] for r in releases if r.get("tag_name", "").endswith(f"-{commit[:12]}"))
    if not tags:
        die(f"{repo} has no release of commit {commit[:12]} yet: run the build workflow on it first")
    if len(tags) > 1:
        die(f"{repo} has several releases of commit {commit[:12]} ({', '.join(tags)}): name one in the lock")
    return f"https://github.com/{repo}/releases/download/{tags[0]}"


def release_base(lock):
    """Base URL of the pinned commit's release assets (RDOCX_RELEASE_URL overrides the lock's `release`)."""
    return (os.environ.get("RDOCX_RELEASE_URL") or lock.get("release") or "").rstrip("/")


def fetch(url, dest):
    """Download `url` to `dest` through a temporary sibling. Nothing is trusted here: callers check hashes."""
    tmp = Path(f"{dest}.part")
    with urllib.request.urlopen(url, timeout=120) as src, open(tmp, "wb") as out:
        shutil.copyfileobj(src, out)
    os.replace(tmp, dest)


def release_listing(base):
    """{platform: {file: sha256}} from the release's SHA256SUMS, whose lines name assets `<platform>.<file>`."""
    tmp = Path(tempfile.mkdtemp(prefix="rdocx-release-"))
    try:
        fetch(f"{base}/SHA256SUMS", tmp / "SHA256SUMS")
        listing = {}
        for line in (tmp / "SHA256SUMS").read_text().splitlines():
            digest, asset = line.split()
            plat, _, name = asset.partition(".")
            if not (re.fullmatch(r"[0-9a-f]{64}", digest) and re.fullmatch(r"[a-z]+-[a-z0-9_]+", plat)
                    and re.fullmatch(r"[A-Za-z0-9_+-][A-Za-z0-9._+-]*", name)):
                die(f"unexpected line in {base}/SHA256SUMS: {line!r}")
            listing.setdefault(plat, {})[name] = digest
        return listing
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def download_release(lock, plat, names):
    """Download `names` of `plat` into HOME/dist/<commit>/<plat>/, a dist folder like any other."""
    base, out = release_base(lock), HOME / "dist" / lock["commit"] / plat
    out.mkdir(parents=True, exist_ok=True)
    for name in names:
        say(f"download {base}/{plat}.{name}")
        fetch(f"{base}/{plat}.{name}", out / name)
    return out


def resolve_ref(upstream, ref):
    """The commit of an upstream tag (peeled) or branch; a full commit hash is taken as is."""
    if re.fullmatch(r"[0-9a-f]{40}", ref):
        return ref
    out = subprocess.run(["git", "ls-remote", upstream, ref, f"{ref}^{{}}"], capture_output=True, text=True,
                         env=GIT_ENV)
    if out.returncode:
        die(f"git ls-remote {upstream} failed: {out.stderr.strip()}")
    refs = {name: oid for oid, name in (line.split("\t") for line in out.stdout.splitlines())}
    for name in (f"refs/tags/{ref}^{{}}", f"refs/tags/{ref}", f"refs/heads/{ref}"):
        if name in refs:
            return refs[name]
    die(f"{ref} is neither a tag nor a branch of {upstream}; give a full 40-character commit hash")


# ---------------------------------------------------------------- verified source export
class ObjectReader:
    """Reads git objects with `cat-file --batch` and checks each one's SHA-1 against its id."""

    def __init__(self, repo):
        self.proc = subprocess.Popen(["git", "-C", str(repo), "cat-file", "--batch"], stdin=subprocess.PIPE,
                                     stdout=subprocess.PIPE, env=GIT_ENV)

    def read(self, oid):
        self.proc.stdin.write(oid.encode() + b"\n")
        self.proc.stdin.flush()
        header = self.proc.stdout.readline().split()
        if len(header) != 3:
            die(f"git object {oid} is missing")
        kind, size = header[1].decode(), int(header[2])
        data = self.proc.stdout.read(size)
        self.proc.stdout.read(1)
        if hashlib.sha1(f"{kind} {size}\0".encode() + data).hexdigest() != oid:
            die(f"git object {oid} does not match its id: refusing the source")
        return kind, data

    def close(self):
        self.proc.stdin.close()
        self.proc.wait()


def export_tree(reader, oid, dest):
    kind, data = reader.read(oid)
    if kind != "tree":
        die(f"{oid} is not a tree")
    dest.mkdir(parents=True, exist_ok=True)
    i = 0
    while i < len(data):
        sp, nul = data.index(b" ", i), data.index(b"\0", i)
        mode, name, child = data[i:sp].decode(), data[sp + 1:nul].decode(), data[nul + 1:nul + 21].hex()
        i = nul + 21
        if name in (".", "..", ".git") or "/" in name or "\\" in name:
            die(f"unsafe path {name!r} in the source tree")
        target = dest / name
        if mode == "40000":
            export_tree(reader, child, target)
        elif mode in ("100644", "100755"):
            _, blob = reader.read(child)
            target.write_bytes(blob)
            os.chmod(target, 0o755 if mode == "100755" else 0o644)
        elif mode == "120000":
            _, blob = reader.read(child)
            link = blob.decode()
            if os.path.isabs(link) or ".." in Path(link).parts:
                die(f"symlink {name} points outside the source tree")
            os.symlink(link, target)
        elif mode == "160000":
            say(f"warning: submodule {name} skipped")
        else:
            die(f"unknown mode {mode} for {name}")


def export_source(lock, dest):
    commit = lock["commit"]
    for src in source_candidates():
        if subprocess.run(["git", "-C", str(src), "cat-file", "-e", f"{commit}^{{commit}}"], capture_output=True,
                          env=GIT_ENV).returncode == 0:
            say(f"source: local clone {src}")
            break
    else:
        src = HOME / "src.git"
        if not src.exists():
            run(["git", "init", "--quiet", "--bare", src], env=GIT_ENV)
            run(["git", "-C", src, "remote", "add", "origin", lock["upstream"]], env=GIT_ENV)
        # a shallow fetch of the one commit brings its whole tree in one pack (a partial clone would fetch the
        # blobs one by one while they are read)
        run(["git", "-C", src, "fetch", "--quiet", "--depth", "1", "origin", commit], env=GIT_ENV)
        say(f"source: {lock['upstream']}, fetched into {src}")
    reader = ObjectReader(src)
    kind, data = reader.read(commit)
    if kind != "commit" or not data.startswith(b"tree "):
        die(f"{commit} is not a commit")
    export_tree(reader, data[5:45].decode(), dest)
    reader.close()
    say(f"source: commit {commit[:12]} exported, every object verified")
    return dest


# ---------------------------------------------------------------- build
def pip_env(venv, requirements):
    """A venv holding exactly the hash-pinned requirements, recreated when the requirements change."""
    stamp = venv / ".requirements.sha256"
    want = sha256(requirements)
    if not (stamp.exists() and stamp.read_text() == want):
        if venv.exists():
            shutil.rmtree(venv)
        run([sys.executable, "-m", "venv", venv])
        run([venv / "bin" / "python", "-m", "pip", "install", "--quiet", "--require-hashes", "--no-deps",
             "--only-binary", ":all:", "-r", requirements])
        stamp.write_text(want)
    return venv


def build(lock, plat, target=None):
    if not shutil.which("cargo"):
        die("cargo not found: install the Rust toolchain (rustup) first; rust-toolchain.toml in rdocx pins the version")
    cross = bool(target) and target != plat
    if cross and (target not in TRIPLES or not plat.startswith("linux")):
        die(f"cross-building is supported between Linux platforms only ({', '.join(TRIPLES)})")
    if cross:
        out = subprocess.run(["cargo", "zigbuild", "--version"], capture_output=True, text=True)
        if out.returncode or ZIGBUILD_VERSION not in out.stdout:
            die(f"cross-building needs cargo-zigbuild {ZIGBUILD_VERSION}: "
                f"cargo install --locked cargo-zigbuild --version {ZIGBUILD_VERSION}")
    plat = target or plat
    commit = lock["commit"]
    work = HOME / "build" / commit[:12]
    if work.exists():
        shutil.rmtree(work)
    src = export_source(lock, work / "src")
    env = dict(os.environ, CARGO_TARGET_DIR=str(HOME / "target" / plat))
    benv = pip_env(HOME / "buildenv", REPO / "scripts" / "build-requirements.txt")
    wheels = work / "wheels"
    if cross:
        triple = TRIPLES[plat]
        zenv = pip_env(HOME / "crossenv", REPO / "scripts" / "cross-requirements.txt")
        env["PATH"] = f"{zenv / 'bin'}:{benv / 'bin'}:{env['PATH']}"
        run(["rustup", "target", "add", triple], cwd=src, env=env)
        run(["cargo", "zigbuild", "--release", "--locked", "--target", f"{triple}.2.17", "-p", "rdocx-cli", "-p",
             "rpptx-cli"], cwd=src, env=env)
        bins = Path(env["CARGO_TARGET_DIR"]) / triple / "release"
        extra = ["--zig", "--target", triple, "--compatibility", "manylinux2014"]
    else:
        run(["cargo", "build", "--release", "--locked", "-p", "rdocx-cli", "-p", "rpptx-cli"], cwd=src, env=env)
        bins = Path(env["CARGO_TARGET_DIR"]) / "release"
        extra = []
    for crate in ("rdocx-py", "rpptx-py"):
        run([benv / "bin" / "maturin", "build", "--release", "--locked", "-m", f"crates/{crate}/Cargo.toml",
             "-o", wheels, *extra], cwd=src, env=env)
    out_root = REPO / "dist" if os.access(REPO, os.W_OK) else HOME / "dist"
    out = out_root / commit / plat
    if out.exists():
        shutil.rmtree(out)
    out.mkdir(parents=True)
    for name in ("rdocx", "rpptx"):
        shutil.copy2(bins / name, out / name)
    for whl in wheels.glob("*.whl"):
        shutil.copy2(whl, out / whl.name)
    sums = {name: sha256(out / name) for name in artifact_names(out)}
    (out / "SHA256SUMS").write_text("".join(f"{d}  {n}\n" for n, d in sums.items()))
    say(f"built into {out}")
    want = expected(lock, plat)
    if want and want != sums:
        say("note: these files differ from the ones recorded in the lock (builds are not bit-reproducible); "
            "they install as a local build")
    for n, d in sums.items():
        print(f"{d}  {n}")
    return out, sums


# ---------------------------------------------------------------- install and status
def site_packages(venv):
    """The venv's site-packages, found without running its interpreter."""
    found = sorted(Path(venv).glob("lib/python3*/site-packages"))
    if len(found) != 1:
        raise OSError(f"cannot find one site-packages folder in {venv}")
    return found[0]


def wrapper_text(dest):
    return (f"#!/bin/sh\nexport PYTHONDONTWRITEBYTECODE=1\n"
            f"exec {shlex.quote(str(dest / 'venv' / 'bin' / 'python'))} \"$@\"\n")


def environment_record(dest):
    """What `status` re-checks outside the two packages: the interpreter link, pyvenv.cfg and every start-up
    file of site-packages (.pth files, sitecustomize, usercustomize)."""
    venv = dest / "venv"
    base = site_packages(venv)
    startup = {p.name: sha256(p) for p in base.iterdir()
               if p.is_file() and (p.suffix == ".pth" or p.name in ("sitecustomize.py", "usercustomize.py"))}
    return {"interpreter": os.path.realpath(venv / "bin" / "python"), "pyvenv_cfg": sha256(venv / "pyvenv.cfg"),
            "startup_files": startup}


def point_current(dest):
    """Point HOME/current at `dest` (the skills and helpers run HOME/current/bin), and rewrite the wrapper."""
    wrapper = dest / "bin" / "python"  # a symlink would hide the venv from the interpreter
    wrapper.write_text(wrapper_text(dest))
    os.chmod(wrapper, 0o755)
    current = HOME / "current"
    if current.is_symlink() or current.exists():
        current.unlink()
    current.symlink_to(dest)


def install_from(lock, plat, folder, local_sums=None):
    """Install from `folder`. Without `local_sums`, every file must match the lock; with them (files just
    built from the pinned commit), they must match those sums and the install is marked as a local build."""
    want = local_sums if local_sums is not None else expected(lock, plat)
    staged, sums = stage(folder, sorted(want))
    try:
        ok, report = check(sums, want)
        for line in report:
            say("  " + line)
        if not ok:
            die(f"{folder}: files do not match {'the lock' if local_sums is None else 'the build'} for {plat}; not installing")
        dest = install_dir(lock)
        if dest.exists():
            shutil.rmtree(dest)
        (dest / "bin").mkdir(parents=True)
        for name in ("rdocx", "rpptx"):
            shutil.copyfile(staged / name, dest / "bin" / name)
            os.chmod(dest / "bin" / name, 0o755)
        run([sys.executable, "-m", "venv", dest / "venv"])
        wheels = [staged / n for n in want if n.endswith(".whl")]
        run([dest / "venv" / "bin" / "python", "-m", "pip", "install", "--quiet", "--no-index", "--no-deps", "--no-compile",
             *wheels])
        files = {}
        for w in wheels:
            files.update(wheel_files(w))
    finally:
        shutil.rmtree(staged, ignore_errors=True)
    source = "lock" if local_sums is None else "local-build"
    (dest / "installed.json").write_text(json.dumps({"commit": lock["commit"], "platform": plat, "source": source,
                                                     "sha256": want, "files": files,
                                                     "environment": environment_record(dest)}, indent=1))
    point_current(dest)
    say(f"installed into {dest} (also {HOME / 'current'}), {source}")


def status(lock, plat, allow_local=False, quiet=False):
    """True when the install matches the lock and `current` points at it; with allow_local, also when it is
    an unchanged local build of the pinned commit."""
    dest = install_dir(lock)
    marker = dest / "installed.json"
    if not marker.exists():
        if not quiet:
            say(f"not installed: {dest}")
        return False
    problems = []
    try:
        info = json.loads(marker.read_text())
    except ValueError:
        info, problems = {}, ["installed.json is unreadable"]
    local = info.get("source") == "local-build"
    ref_bins = info.get("sha256", {}) if local else expected(lock, plat)
    ref_files = info.get("files", {}) if local else expected_files(lock, plat)
    if info.get("commit") != lock["commit"] or info.get("platform") != plat:
        problems.append("installed for another commit or platform")
    if local and expected(lock, plat) and not allow_local:
        problems.append("a local build, while the lock has hashes for this platform (reinstall, or --allow-local)")
    current = HOME / "current"
    if not current.exists() or current.resolve() != dest.resolve():
        problems.append(f"{current} does not point at {dest} (rerun install)")
    for name in ("rdocx", "rpptx"):
        path = dest / "bin" / name
        if not path.is_file() or sha256(path) != ref_bins.get(name):
            problems.append(f"{name} does not match the {'build record' if local else 'lock'}")
    wrapper = dest / "bin" / "python"
    if not wrapper.is_file() or wrapper.read_text() != wrapper_text(dest):
        problems.append("the bin/python wrapper was changed (rerun install)")
    if not ref_files:
        problems.append("no hashes of the installed Python files to check against")
    else:
        try:
            base = site_packages(dest / "venv")
            for rel, digest in ref_files.items():
                if not (base / rel).is_file() or sha256(base / rel) != digest:
                    problems.append(f"python file {rel} does not match")
                    break
            for top in sorted({rel.split("/")[0] for rel in ref_files}):
                for p in (base / top).rglob("*"):
                    rel = p.relative_to(base).as_posix()
                    if (p.is_file() or p.is_symlink()) and rel not in ref_files:
                        problems.append(f"unexpected file {rel} in the installed packages")
                        break
            env = info.get("environment")
            if env is None or environment_record(dest) != env:
                problems.append("the Python environment (interpreter, pyvenv.cfg or start-up files) changed")
        except OSError as e:
            problems.append(f"the Python environment is broken ({e})")
    ok = not problems
    if not quiet:
        say(f"platform {plat}, commit {lock['commit'][:12]}, install {dest}")
        if problems:
            say("NOT VERIFIED: " + "; ".join(problems))
        elif local:
            say("local build of the pinned commit, unchanged since install; its hashes are not in the lock "
                "(record them with `lock --write` after review)")
        else:
            say("verified against the lock: CLIs, wrapper, every file of the rdocx and rpptx packages, start-up files")
    return ok


def test_env(lock):
    """A separate venv with the hash-pinned test dependencies, which sees the installed packages through a
    .pth file: running the suite never adds anything to the runtime environment."""
    venv = pip_env(HOME / "testenv", REPO / "tests" / "requirements.txt")
    (site_packages(venv) / "rdocx_runtime.pth").write_text(str(site_packages(install_dir(lock) / "venv")) + "\n")
    return venv / "bin" / "python"


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)
    p = sub.add_parser("status")
    p.add_argument("--allow-local", action="store_true", help="accept an unchanged local build of the pinned commit")
    p = sub.add_parser("install")
    p.add_argument("--allow-local", action="store_true", help="accept an unchanged local build of the pinned commit")
    p.add_argument("--build", action="store_true", help="build from source when no verified dist is found")
    p.add_argument("--from", dest="src", help="install from this folder (hashes checked against the lock)")
    p = sub.add_parser("build")
    p.add_argument("--target", choices=sorted(TRIPLES), help="cross-build for another Linux platform")
    p = sub.add_parser("lock")
    p.add_argument("--write", action="store_true")
    p.add_argument("--platform", help="platform of the dist folder to check or record (default: this machine; "
                                      "with --release: every platform of the release)")
    p.add_argument("--release", action="store_true", help="record the files of the release named in the lock")
    p = sub.add_parser("bump")
    p.add_argument("ref", help="upstream tag, branch or full commit hash")
    sub.add_parser("paths")
    p = sub.add_parser("test")
    p.add_argument("--allow-local", action="store_true", help="accept an unchanged local build of the pinned commit")
    args, extra = ap.parse_known_args()
    if extra and args.cmd != "test":
        ap.error(f"unrecognized arguments: {' '.join(extra)}")
    if platform.system() not in ("Linux", "Darwin"):
        die("only Linux and macOS are supported")
    lock, plat = load_lock(), platform_key()

    if args.cmd == "status":
        sys.exit(0 if status(lock, plat, allow_local=args.allow_local) else 1)
    if args.cmd == "paths":
        d = install_dir(lock) / "bin"
        print(f"RDOCX={d / 'rdocx'}\nRPPTX={d / 'rpptx'}\nRDOCX_PY={d / 'python'}")
        return
    if args.cmd == "build":
        build(lock, plat, args.target)
        return
    if args.cmd == "test":
        if not status(lock, plat, allow_local=args.allow_local or not expected(lock, plat)):
            die("the pinned build is not installed and verified: run `install` first")
        py = test_env(lock)
        env = dict(os.environ, RDOCX_BIN_DIR=str(install_dir(lock) / "bin"), PYTHONDONTWRITEBYTECODE="1")
        sys.exit(subprocess.run([str(py), "-m", "pytest", "-p", "no:cacheprovider", str(REPO / "tests"), *extra],
                                env=env).returncode)
    if args.cmd == "bump":
        commit = resolve_ref(lock["upstream"], args.ref)
        if commit == lock["commit"]:
            say(f"{args.ref} is {commit}, already pinned")
            return
        lock.update(commit=commit, ref=args.ref, release="", versions={}, artifacts={}, installed_files={})
        write_lock(lock)
        say(f"pinned {args.ref} = {commit}; no release nor hashes recorded yet. Next: `install --build` and `test` "
            f"here, then push: the build workflow publishes the release rdocx-<date>-{commit[:12]}; then "
            "`lock --write --release` finds it and records it.")
        return
    if args.cmd == "lock" and args.release:
        if not args.write:
            die("--release records hashes: use it with --write")
        if not lock.get("release"):
            lock["release"] = release_url(RELEASES_REPO, github_releases(RELEASES_REPO), lock["commit"])
            say(f"release: {lock['release']}")
        base = release_base(lock)
        if not base:
            die("the lock names no release")
        listing = release_listing(base)
        if args.platform:
            listing = {args.platform: listing.get(args.platform) or die(f"no {args.platform} files in {base}")}
        for p, want in sorted(listing.items()):
            folder = download_release(lock, p, sorted(want))
            staged, sums = stage(folder, sorted(want))
            try:
                ok, report = check(sums, want)
                if not ok:
                    die(f"{p}: the downloaded files do not match the release's SHA256SUMS\n" + "\n".join(report))
                record(lock, p, staged, sums)
            finally:
                shutil.rmtree(staged, ignore_errors=True)
        others = sorted(set(lock.get("artifacts", {})) - set(listing))
        if others:
            say(f"note: {', '.join(others)} keep hashes that do not come from this release")
        write_lock(lock)
        return
    if args.cmd == "lock":
        plat = args.platform or plat
        folder = next(dist_candidates(lock, plat), None)
        if folder is None:
            die(f"no dist folder for {plat} at commit {lock['commit'][:12]}")
        staged, sums = stage(folder, artifact_names(folder))
        try:
            if args.write:
                record(lock, plat, staged, sums)
                write_lock(lock)
            else:
                ok, report = check(sums, expected(lock, plat))
                print("\n".join(report) or "no hashes recorded for this platform")
                sys.exit(0 if ok else 1)
        finally:
            shutil.rmtree(staged, ignore_errors=True)
        return
    if args.cmd == "install":
        allow_local = args.allow_local or not expected(lock, plat)
        dest = install_dir(lock)
        if (dest / "installed.json").exists():
            point_current(dest)  # the skills run HOME/current: make it this pin, then verify everything
            if status(lock, plat, allow_local=allow_local, quiet=True):
                say("already installed and unchanged")
                status(lock, plat, allow_local=allow_local)
                return
        if args.src:
            install_from(lock, plat, args.src)
        else:
            for folder in dist_candidates(lock, plat):
                staged, sums = stage(folder, sorted(expected(lock, plat)))
                shutil.rmtree(staged, ignore_errors=True)
                if check(sums, expected(lock, plat))[0]:
                    install_from(lock, plat, folder)
                    break
            else:
                folder = None
                if expected(lock, plat) and release_base(lock):
                    try:
                        folder = download_release(lock, plat, sorted(expected(lock, plat)))
                    except OSError as e:
                        say(f"no download from {release_base(lock)}: {e}")
                if folder is not None:
                    install_from(lock, plat, folder)  # refuses any file that does not match the lock
                elif not args.build:
                    die(f"no verified build for {plat} at {lock['commit'][:12]}; rerun with --build to compile it "
                        "(10 to 30 minutes), or install --from a folder of prebuilt files", code=2)
                else:
                    out, sums = build(lock, plat)
                    install_from(lock, plat, out, local_sums=sums)
                    allow_local = True
        if not status(lock, plat, allow_local=allow_local):
            sys.exit(1)


if __name__ == "__main__":
    main()
