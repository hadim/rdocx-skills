"""scripts/rdocx_env.py and the releases it installs from: the builds of this repository (assets
`<platform>.<file>`). `install` downloads the pinned files and checks them against the lock, `lock --write
--release` records the release (and finds it by commit when the lock names none yet), `bump` moves the pin. Each release is served from a local
folder through RDOCX_RELEASE_URL (file://, one folder per tag), and the script runs from a copy of the
repository without its dist folder. The download tests are skipped when this machine has no verified build to
serve."""
import json
import os
import shutil
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
WANT = rdocx_env.expected(LOCK, PLAT)
RDOCX, RPPTX = rdocx_env.cli_file(PLAT, "rdocx"), rdocx_env.cli_file(PLAT, "rpptx")
# this repository's release tags: UTC build date and time (to the minute), and date only (the older releases)
OWN_TAG = f"rdocx-202610041405-{LOCK['commit'][:12]}"
OWN_TAG_DAY = f"rdocx-20261004-{LOCK['commit'][:12]}"


def repo_copy(root, lock=LOCK):
    (root / "scripts").mkdir(parents=True)
    shutil.copy(ROOT / "plugin" / "scripts" / "rdocx_env.py", root / "scripts")
    (root / "rdocx.lock.json").write_text(json.dumps(lock, indent=2) + "\n")
    return root


def cmd(repo, home, mirror, *args):
    env = dict(os.environ, RDOCX_HOME=str(home), RDOCX_RELEASE_URL=Path(mirror).as_uri())
    env.pop("RDOCX_DIST", None)
    env.pop("RDOCX_SRC", None)
    return subprocess.run([sys.executable, repo / "scripts" / "rdocx_env.py", *args], env=env, capture_output=True,
                          text=True)


def sums_file(folder):
    files = sorted(p for p in folder.iterdir() if p.name != "SHA256SUMS")
    (folder / "SHA256SUMS").write_text("".join(f"{rdocx_env.sha256(p)}  {p.name}\n" for p in files))


@pytest.fixture(scope="module")
def dist():
    folder = verified_dist()
    if folder is None:
        pytest.skip(f"no verified build for {PLAT} to serve as a release")
    return folder


def own_release(dist, mirror, tag):
    """(mirror, lock): this repository's release `tag` of the pinned commit, as the build workflow publishes it."""
    lock = json.loads(json.dumps(LOCK))
    lock["release"] = f"https://github.com/hadim/rdocx-skills/releases/download/{tag}"
    lock["artifacts"], lock["installed_files"] = {PLAT: WANT}, {PLAT: rdocx_env.expected_files(LOCK, PLAT)}
    folder = mirror / tag
    folder.mkdir()
    for name in WANT:
        rdocx_env.copy_artifact(dist, name, folder / f"{PLAT}.{name}")
    sums_file(folder)
    return mirror, lock


@pytest.fixture(scope="module")
def own(dist, tmp_path_factory):
    return own_release(dist, tmp_path_factory.mktemp("own"), OWN_TAG)


@pytest.fixture(scope="module")
def own_day(dist, tmp_path_factory):
    return own_release(dist, tmp_path_factory.mktemp("own_day"), OWN_TAG_DAY)


@pytest.fixture(params=["own", "own_day"])
def release(request):
    return request.param, *request.getfixturevalue(request.param)


def test_install_downloads_the_release_and_verifies_it(release, tmp_path):
    _, mirror, lock = release
    repo, home = repo_copy(tmp_path / "repo", lock), tmp_path / "home"
    res = cmd(repo, home, mirror, "install")
    assert res.returncode == 0, res.stderr
    assert f"download {mirror.as_uri()}/{lock['release'].rsplit('/', 1)[-1]}/{PLAT}.{RDOCX}" in res.stderr
    res = cmd(repo, home, mirror, "status")
    assert res.returncode == 0 and "verified against the lock" in res.stderr


def test_install_refuses_a_changed_release_file(release, tmp_path):
    _, mirror, lock = release
    changed = Path(shutil.copytree(mirror, tmp_path / "changed"))
    cli = next(p for p in sorted(changed.rglob("*")) if p.is_file() and p.name != "SHA256SUMS"
               and not p.name.endswith(".whl"))  # a CLI
    with open(cli, "ab") as f:
        f.write(b"\0")
    repo, home = repo_copy(tmp_path / "repo", lock), tmp_path / "home"
    res = cmd(repo, home, changed, "install")
    assert res.returncode == 1 and "do not match the lock" in res.stderr
    assert not (home / LOCK["commit"][:12] / "installed.json").exists()


def test_install_without_a_release_exits_2(release, tmp_path):
    _, _, lock = release
    (tmp_path / "empty").mkdir()
    res = cmd(repo_copy(tmp_path / "repo", lock), tmp_path / "home", tmp_path / "empty", "install")
    assert res.returncode == 2 and "no verified build" in res.stderr


def test_lock_write_release_records_the_same_hashes(release, tmp_path):
    _, mirror, lock = release
    partial = json.loads(json.dumps(lock))
    for key in ("artifacts", "installed_files"):
        partial.get(key, {}).pop(PLAT, None)
    partial["versions"] = {}
    repo = repo_copy(tmp_path / "repo", partial)
    res = cmd(repo, tmp_path / "home", mirror, "lock", "--write", "--release")
    assert res.returncode == 0, res.stderr
    assert json.loads((repo / "rdocx.lock.json").read_text()) == lock


def test_lock_write_release_refuses_files_that_differ_from_its_listing(release, tmp_path):
    _, mirror, lock = release
    changed = Path(shutil.copytree(mirror, tmp_path / "changed"))
    with open(next(p for p in changed.rglob("*.whl") if "rpptx" in p.name), "ab") as f:
        f.write(b"\0")
    repo = repo_copy(tmp_path / "repo", lock)
    res = cmd(repo, tmp_path / "home", changed, "lock", "--write", "--release")
    assert res.returncode == 1 and "do not match the release's SHA256SUMS" in res.stderr
    assert json.loads((repo / "rdocx.lock.json").read_text()) == lock


def test_bump_to_a_commit_empties_the_hashes(tmp_path):
    repo, other = repo_copy(tmp_path / "repo"), "0123456789abcdef0123456789abcdef01234567"
    res = cmd(repo, tmp_path / "home", tmp_path, "bump", other)
    assert res.returncode == 0, res.stderr
    lock = json.loads((repo / "rdocx.lock.json").read_text())
    assert lock["commit"] == other and lock["ref"] == other
    assert lock["release"] == "" and "rdocx-<YYYYMMDDHHMM>-0123456789ab" in res.stderr
    assert lock["artifacts"] == lock["installed_files"] == lock["versions"] == {}
    res = cmd(repo, tmp_path / "home", tmp_path, "bump", other)
    assert res.returncode == 0 and "already pinned" in res.stderr
    assert json.loads((repo / "rdocx.lock.json").read_text()) == lock


@pytest.mark.parametrize("tag", ["rdocx-202609301405-0123456789ab", "rdocx-20260930-0123456789ab"])
def test_release_url_is_the_release_whose_tag_ends_with_the_commit(tag):
    commit = "0123456789abcdef0123456789abcdef01234567"
    tags = ["rdocx-20260929-f2fa36d1d18a", "rdocx-202610011200-f2fa36d1d18b", tag]
    assert (rdocx_env.release_url("owner/repo", tags, commit)
            == f"https://github.com/owner/repo/releases/download/{tag}")


@pytest.mark.parametrize("tags, message", [
    ([], "no release of commit 0123456789ab"),
    (["rdocx-20260929-0123456789ab", "rdocx-20260930-0123456789ab"], "several releases of commit 0123456789ab"),
    (["rdocx-20260930-0123456789ab", "rdocx-202610011200-0123456789ab"], "several releases of commit 0123456789ab"),
])
def test_release_url_refuses_no_or_several_releases(tags, message, capsys):
    with pytest.raises(SystemExit):
        rdocx_env.release_url("owner/repo", tags, "0123456789ab" + "0" * 28)
    assert message in capsys.readouterr().err


def test_release_tags_come_from_git_without_a_token(monkeypatch):
    out = ("1111111111111111111111111111111111111111\trefs/tags/rdocx-20260929-f2fa36d1d18a\n"
           "2222222222222222222222222222222222222222\trefs/tags/rdocx-202610011200-0123456789ab\n"
           "3333333333333333333333333333333333333333\trefs/tags/rdocx-202610011200-0123456789ab^{}\n")
    seen = []

    def run(cmd, **kwargs):
        seen.append((cmd, kwargs["env"]))
        return subprocess.CompletedProcess(cmd, 0, out, "")

    monkeypatch.setattr(rdocx_env.subprocess, "run", run)
    monkeypatch.setenv("GITHUB_TOKEN", "secret")
    assert rdocx_env.release_tags("owner/repo") == ["rdocx-20260929-f2fa36d1d18a", "rdocx-202610011200-0123456789ab"]
    assert [cmd for cmd, _ in seen] == [["git", "ls-remote", "--tags", "https://github.com/owner/repo"]]
    assert "secret" not in " ".join(seen[0][0])