"""scripts/rdocx_env.py: a fresh install from the verified dist folder, then each kind of change `status` must
catch and `install` must repair. Skipped when this machine's dist folder is not present."""
import hashlib
import json
import os
import subprocess
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "scripts"))
import rdocx_env  # noqa: E402

LOCK = rdocx_env.load_lock()
PLAT = rdocx_env.platform_key()
DIST = ROOT / "dist" / LOCK["commit"] / PLAT


def env_cmd(home, *args):
    env = dict(os.environ, RDOCX_HOME=str(home))
    env.pop("RDOCX_DIST", None)
    return subprocess.run([sys.executable, ROOT / "scripts" / "rdocx_env.py", *args], env=env, capture_output=True, text=True)


@pytest.fixture(scope="module")
def home(tmp_path_factory):
    if not DIST.is_dir() or not rdocx_env.expected(LOCK, PLAT):
        pytest.skip(f"no verified dist folder for {PLAT}")
    home = tmp_path_factory.mktemp("rdocx-home")
    res = env_cmd(home, "install")
    assert res.returncode == 0, res.stderr
    return home


def dest(home):
    return home / LOCK["commit"][:12]


def site(home):
    return next((dest(home) / "venv").glob("lib/python3*/site-packages"))


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
    cli = dest(home) / "bin" / "rdocx"
    cli.write_bytes((dest(home) / "bin" / "rpptx").read_bytes())
    info["source"] = "local-build"
    info["sha256"]["rdocx"] = hashlib.sha256(cli.read_bytes()).hexdigest()
    marker.write_text(json.dumps(info))


def repointed_current(home):
    other = home / "other" / "bin"
    other.mkdir(parents=True, exist_ok=True)
    (home / "current").unlink()
    (home / "current").symlink_to(other.parent)


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
