"""scripts/rdocx_env.py: a fresh install from a verified dist folder of this machine (the repository's dist/,
or the release files downloaded into RDOCX_HOME/dist), then each kind of change `status` must catch and
`install` must repair. Skipped when there is none."""
import datetime
import hashlib
import json
import os
import subprocess
import sys
from pathlib import Path

import pytest

from conftest import verified_dist

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "plugin" / "scripts"))
import rdocx_env  # noqa: E402

LOCK = rdocx_env.load_lock()
PLAT = rdocx_env.platform_key()
DIST = verified_dist()


def env_cmd(home, *args):
    env = dict(os.environ, RDOCX_HOME=str(home))
    env.pop("RDOCX_DIST", None)
    if DIST is not None:  # <root>/<commit>/<platform>
        env["RDOCX_DIST"] = str(DIST.parents[1])
    return subprocess.run([sys.executable, ROOT / "plugin" / "scripts" / "rdocx_env.py", *args], env=env, capture_output=True, text=True)


@pytest.fixture(scope="module")
def home(tmp_path_factory):
    if DIST is None:
        pytest.skip(f"no verified dist folder for {PLAT}")
    home = tmp_path_factory.mktemp("rdocx-home")
    res = env_cmd(home, "install")
    assert res.returncode == 0, res.stderr
    return home


def dest(home):
    return home / LOCK["commit"][:12]


def site(home):
    return rdocx_env.site_packages(dest(home) / "venv")


def extra_pth(home):
    (site(home) / "zz_extra.pth").write_text("import sys\n")


def edited_wrapper(home):
    (dest(home) / "bin" / "python").write_text("#!/bin/sh\necho changed\n")


def stray_cache(home):
    cache = site(home) / "rdocx" / "__pycache__"
    cache.mkdir(exist_ok=True)
    (cache / "x.cpython-311.pyc").write_bytes(b"x")


def forged_local_build(home):
    marker = dest(home) / "installed.json"
    info = json.loads(marker.read_text())
    cli = dest(home) / "bin" / rdocx_env.cli_file(PLAT, "rdocx")
    cli.write_bytes((dest(home) / "bin" / rdocx_env.cli_file(PLAT, "rpptx")).read_bytes())
    info["source"] = "local-build"
    info["sha256"][cli.name] = hashlib.sha256(cli.read_bytes()).hexdigest()
    marker.write_text(json.dumps(info))


def repointed_current(home):
    other = home / "other" / "bin"
    other.mkdir(parents=True, exist_ok=True)
    rdocx_env.link_folder(home / "current", other.parent)


def changed_package_file(home):
    with open(site(home) / "rdocx" / "shared.py", "a") as f:
        f.write("# changed\n")


CHANGES = [extra_pth, edited_wrapper, stray_cache, forged_local_build, repointed_current, changed_package_file]


def test_fresh_install_is_verified(home):
    res = env_cmd(home, "status")
    assert res.returncode == 0 and "verified against the lock" in res.stderr
    assert (home / "current").resolve() == dest(home).resolve()


@pytest.mark.parametrize("change", CHANGES, ids=[c.__name__ for c in CHANGES])
def test_status_catches_and_install_repairs(home, change):
    change(home)
    assert env_cmd(home, "status").returncode == 1
    assert env_cmd(home, "install").returncode == 0
    assert env_cmd(home, "status").returncode == 0


def test_install_keeps_the_pin_and_the_previous_install_only(home):
    """Each pin would otherwise leave about 100 MB behind: `install` keeps the pinned install and the most
    recent other one, with their downloads, and removes the older ones; caches and environments stay."""
    olds = ["a" * 12, "b" * 12, "c" * 12]
    for age, name in enumerate(olds):
        marker = home / name / "installed.json"
        marker.parent.mkdir()
        marker.write_text("{}")
        os.utime(marker, (1_000_000 + age, 1_000_000 + age))  # c is the most recent of the three
        (home / "dist" / (name + "0" * 28)).mkdir(parents=True)
        (home / "build" / name).mkdir(parents=True)
    for keep in ("target", "testenv"):
        (home / keep).mkdir(exist_ok=True)
    res = env_cmd(home, "install")
    assert res.returncode == 0, res.stderr
    assert dest(home).is_dir() and (home / ("c" * 12)).is_dir() and (home / "dist" / ("c" * 12 + "0" * 28)).is_dir()
    for name in ("a" * 12, "b" * 12):
        assert not (home / name).exists() and not (home / "dist" / (name + "0" * 28)).exists()
        assert not (home / "build" / name).exists()
    assert (home / "target").is_dir() and (home / "testenv").is_dir()
    assert env_cmd(home, "status").returncode == 0



def test_a_source_export_carries_the_commit_time(tmp_path, monkeypatch):
    """Cargo compiles a crate again when one of its files is newer than its last build: every export of a commit
    gives each file the commit's time, so that the restored build folder of that commit compiles nothing again."""
    when = "2026-10-17T09:49:21+01:00"
    env = dict(os.environ, GIT_AUTHOR_DATE=when, GIT_COMMITTER_DATE=when, GIT_AUTHOR_NAME="t", GIT_COMMITTER_NAME="t",
               GIT_AUTHOR_EMAIL="t@example.com", GIT_COMMITTER_EMAIL="t@example.com")
    repo = tmp_path / "rdocx"
    (repo / "crates" / "a").mkdir(parents=True)
    (repo / "crates" / "a" / "lib.rs").write_text("pub fn a() {}\n")
    (repo / "Cargo.toml").write_text("[workspace]\n")
    for args in (["init", "-q"], ["add", "-A"], ["-c", "commit.gpgsign=false", "commit", "-q", "-m", "x"]):
        subprocess.run(["git", "-C", str(repo), *args], check=True, env=env)
    commit = subprocess.run(["git", "-C", str(repo), "rev-parse", "HEAD"], check=True, capture_output=True,
                            text=True).stdout.strip()
    monkeypatch.setenv("RDOCX_SRC", str(repo))
    out = rdocx_env.export_source({"commit": commit, "upstream": "unused"}, tmp_path / "src")
    files = [p for p in out.rglob("*") if p.is_file()]
    stamp = int(datetime.datetime.fromisoformat(when).timestamp())
    assert len(files) == 2 and {int(p.stat().st_mtime) for p in files} == {stamp}
    # a build folder restored from another commit can be newer than this one: build() then gives the export's time
    out = rdocx_env.export_source({"commit": commit, "upstream": "unused"}, tmp_path / "src2", mtime=stamp + 99)
    assert {int(p.stat().st_mtime) for p in out.rglob("*") if p.is_file()} == {stamp + 99}
