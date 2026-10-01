"""scripts/rdocx_env.py and the releases of this repository: `install` downloads the pinned files and checks them
against the lock, `lock --write --release` records a release (and finds it by commit among the repository's releases when the lock
names none yet), `bump` moves the pin. The release is served from a
local folder through RDOCX_RELEASE_URL (file://), and the script runs from a copy of the repository without its
dist folder. The download tests are skipped when this machine has no verified build to serve."""
import io
import json
import os
import shutil
import subprocess
import sys
from pathlib import Path

import pytest

from conftest import verified_dist

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "scripts"))
import rdocx_env  # noqa: E402

LOCK = rdocx_env.load_lock()
PLAT = rdocx_env.platform_key()
WANT = rdocx_env.expected(LOCK, PLAT)
RDOCX, RPPTX = rdocx_env.cli_file(PLAT, "rdocx"), rdocx_env.cli_file(PLAT, "rpptx")


def repo_copy(root, lock=LOCK):
    (root / "scripts").mkdir(parents=True)
    shutil.copy(ROOT / "scripts" / "rdocx_env.py", root / "scripts")
    (root / "rdocx.lock.json").write_text(json.dumps(lock, indent=2) + "\n")
    return root


def cmd(repo, home, release, *args):
    env = dict(os.environ, RDOCX_HOME=str(home), RDOCX_RELEASE_URL=Path(release).as_uri())
    env.pop("RDOCX_DIST", None)
    env.pop("RDOCX_SRC", None)
    return subprocess.run([sys.executable, repo / "scripts" / "rdocx_env.py", *args], env=env, capture_output=True,
                          text=True)


@pytest.fixture(scope="module")
def release(tmp_path_factory):
    """A release folder as the build workflow publishes it: `<platform>.<file>` and SHA256SUMS."""
    folder = verified_dist()
    if folder is None:
        pytest.skip(f"no verified build for {PLAT} to serve as a release")
    out = tmp_path_factory.mktemp("release")
    for name in WANT:
        rdocx_env.copy_artifact(folder, name, out / f"{PLAT}.{name}")
    (out / "SHA256SUMS").write_text("".join(f"{d}  {PLAT}.{n}\n" for n, d in sorted(WANT.items())))
    return out


def test_install_downloads_the_release_and_verifies_it(release, tmp_path):
    repo, home = repo_copy(tmp_path / "repo"), tmp_path / "home"
    res = cmd(repo, home, release, "install")
    assert res.returncode == 0, res.stderr
    assert f"download {release.as_uri()}/{PLAT}.{RDOCX}" in res.stderr
    res = cmd(repo, home, release, "status")
    assert res.returncode == 0 and "verified against the lock" in res.stderr


def test_install_refuses_a_changed_release_file(release, tmp_path):
    changed = Path(shutil.copytree(release, tmp_path / "changed"))
    with open(changed / f"{PLAT}.{RDOCX}", "ab") as f:
        f.write(b"\0")
    repo, home = repo_copy(tmp_path / "repo"), tmp_path / "home"
    res = cmd(repo, home, changed, "install")
    assert res.returncode == 1 and "do not match the lock" in res.stderr
    assert not (home / LOCK["commit"][:12] / "installed.json").exists()


def test_install_without_a_release_exits_2(tmp_path):
    if not WANT:
        pytest.skip(f"the lock has no hashes for {PLAT}")
    (tmp_path / "empty").mkdir()
    res = cmd(repo_copy(tmp_path / "repo"), tmp_path / "home", tmp_path / "empty", "install")
    assert res.returncode == 2 and "no verified build" in res.stderr


def test_lock_write_release_records_the_same_hashes(release, tmp_path):
    lock = json.loads(json.dumps(LOCK))
    lock["artifacts"].pop(PLAT)
    lock["installed_files"].pop(PLAT)
    lock["versions"] = {}
    repo = repo_copy(tmp_path / "repo", lock)
    res = cmd(repo, tmp_path / "home", release, "lock", "--write", "--release")
    assert res.returncode == 0, res.stderr
    assert json.loads((repo / "rdocx.lock.json").read_text()) == LOCK


def test_lock_write_release_refuses_files_that_differ_from_its_listing(release, tmp_path):
    changed = Path(shutil.copytree(release, tmp_path / "changed"))
    with open(changed / f"{PLAT}.{RPPTX}", "ab") as f:
        f.write(b"\0")
    repo = repo_copy(tmp_path / "repo")
    res = cmd(repo, tmp_path / "home", changed, "lock", "--write", "--release")
    assert res.returncode == 1 and "do not match the release's SHA256SUMS" in res.stderr
    assert json.loads((repo / "rdocx.lock.json").read_text()) == LOCK


def test_bump_to_a_commit_empties_the_hashes(tmp_path):
    repo, other = repo_copy(tmp_path / "repo"), "0123456789abcdef0123456789abcdef01234567"
    res = cmd(repo, tmp_path / "home", tmp_path, "bump", other)
    assert res.returncode == 0, res.stderr
    lock = json.loads((repo / "rdocx.lock.json").read_text())
    assert lock["commit"] == other and lock["ref"] == other
    assert lock["release"] == "" and "rdocx-<date>-0123456789ab" in res.stderr
    assert lock["artifacts"] == lock["installed_files"] == lock["versions"] == {}
    res = cmd(repo, tmp_path / "home", tmp_path, "bump", other)
    assert res.returncode == 0 and "already pinned" in res.stderr
    assert json.loads((repo / "rdocx.lock.json").read_text()) == lock


def test_release_url_is_the_release_whose_tag_ends_with_the_commit():
    commit = "0123456789abcdef0123456789abcdef01234567"
    releases = [{"tag_name": "rdocx-20260929-f2fa36d1d18a"}, {"tag_name": "rdocx-20260930-0123456789ab"}]
    assert (rdocx_env.release_url("owner/repo", releases, commit)
            == "https://github.com/owner/repo/releases/download/rdocx-20260930-0123456789ab")


@pytest.mark.parametrize("tags, message", [
    ([], "no release of commit 0123456789ab"),
    (["rdocx-20260929-0123456789ab", "rdocx-20260930-0123456789ab"], "several releases of commit 0123456789ab"),
])
def test_release_url_refuses_no_or_several_releases(tags, message, capsys):
    with pytest.raises(SystemExit):
        rdocx_env.release_url("owner/repo", [{"tag_name": t} for t in tags], "0123456789ab" + "0" * 28)
    assert message in capsys.readouterr().err


def test_github_releases_reads_every_page_and_sends_the_token(monkeypatch):
    pages, seen = {1: [{"tag_name": f"t{i}"} for i in range(100)], 2: [{"tag_name": "last"}]}, []

    def urlopen(req, timeout):
        seen.append((req.full_url, req.get_header("Authorization")))
        return io.BytesIO(json.dumps(pages[int(req.full_url.rsplit("=", 1)[1])]).encode())

    monkeypatch.setattr(rdocx_env.urllib.request, "urlopen", urlopen)
    monkeypatch.delenv("GH_TOKEN", raising=False)
    monkeypatch.setenv("GITHUB_TOKEN", "secret")
    assert len(rdocx_env.github_releases("owner/repo")) == 101
    assert seen == [(f"https://api.github.com/repos/owner/repo/releases?per_page=100&page={n}", "Bearer secret")
                    for n in (1, 2)]
