"""scripts/rdocx_env.py and the releases it installs from: the upstream releases of rdocx and rpptx (one per
family, CLI archives and wheels) and the builds of this repository (assets `<platform>.<file>`). `install`
downloads the pinned files and checks them against the lock, `lock --write --release` records the releases (and
finds them by commit when the lock names none yet), `bump` moves the pin. Each release is served from a local
folder through RDOCX_RELEASE_URL (file://, one folder per tag), and the script runs from a copy of the
repository without its dist folder. The download tests are skipped when this machine has no verified build to
serve."""
import io
import json
import os
import shutil
import subprocess
import sys
import tarfile
import zipfile
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
OWN_TAG = f"rdocx-20261004-{LOCK['commit'][:12]}"


def repo_copy(root, lock=LOCK):
    (root / "scripts").mkdir(parents=True)
    shutil.copy(ROOT / "scripts" / "rdocx_env.py", root / "scripts")
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


def archive(folder, family, cli):
    """A CLI archive as upstream publishes it: the CLI next to a README and a LICENSE."""
    triple = rdocx_env.UPSTREAM_PLATFORMS[PLAT][0]
    if PLAT.startswith("windows-"):
        with zipfile.ZipFile(folder / f"{family}-{triple}.zip", "w") as z:
            z.writestr("README.md", "readme\n")
            z.write(cli, cli.name)
    else:
        with tarfile.open(folder / f"{family}-{triple}.tar.gz", "w:gz") as t:
            t.add(cli, cli.name)
            info = tarfile.TarInfo("LICENSE")
            info.size = 8
            t.addfile(info, io.BytesIO(b"license\n"))


@pytest.fixture(scope="module")
def dist():
    folder = verified_dist()
    if folder is None:
        pytest.skip(f"no verified build for {PLAT} to serve as a release")
    return folder


@pytest.fixture(scope="module")
def upstream(dist, tmp_path_factory):
    """(mirror, lock): the two upstream releases of this platform's files, one folder per tag, and the lock that
    pins them (its `downloads` are the hashes of these archives)."""
    mirror, lock = tmp_path_factory.mktemp("upstream"), json.loads(json.dumps(LOCK))
    lock["release"] = {f: f"https://github.com/tensorbee/rdocx/releases/download/{t}"
                       for f, t in (("rdocx", "v0.0.1"), ("rpptx", "rpptx-v0.0.2"))}
    staged = tmp_path_factory.mktemp("staged")
    downloads = {}
    for family, url in lock["release"].items():
        folder = mirror / url.rsplit("/", 1)[-1]
        folder.mkdir()
        cli = staged / rdocx_env.cli_file(PLAT, family)
        rdocx_env.copy_artifact(dist, cli.name, cli)
        archive(folder, family, cli)
        wheel = next(n for n in WANT if n.startswith(f"{family}-") and n.endswith(".whl"))
        shutil.copyfile(dist / wheel, folder / wheel)
        sums_file(folder)
        downloads.update(rdocx_env.release_sums(folder.as_uri()))
    lock["downloads"] = {PLAT: downloads}
    return mirror, lock


@pytest.fixture(scope="module")
def own(dist, tmp_path_factory):
    """(mirror, lock): this repository's release of the pinned commit, as the build workflow publishes it."""
    mirror, lock = tmp_path_factory.mktemp("own"), json.loads(json.dumps(LOCK))
    lock["release"] = f"https://github.com/hadim/rdocx-skills/releases/download/{OWN_TAG}"
    lock.pop("downloads", None)
    lock["artifacts"], lock["installed_files"] = {PLAT: WANT}, {PLAT: rdocx_env.expected_files(LOCK, PLAT)}
    folder = mirror / OWN_TAG
    folder.mkdir()
    for name in WANT:
        rdocx_env.copy_artifact(dist, name, folder / f"{PLAT}.{name}")
    sums_file(folder)
    return mirror, lock


@pytest.fixture(params=["upstream", "own"])
def release(request):
    return request.param, *request.getfixturevalue(request.param)


def test_install_downloads_the_release_and_verifies_it(release, tmp_path):
    kind, mirror, lock = release
    repo, home = repo_copy(tmp_path / "repo", lock), tmp_path / "home"
    res = cmd(repo, home, mirror, "install")
    assert res.returncode == 0, res.stderr
    first = f"v0.0.1/rdocx-{rdocx_env.UPSTREAM_PLATFORMS[PLAT][0]}" if kind == "upstream" else f"{OWN_TAG}/{PLAT}.{RDOCX}"
    assert f"download {mirror.as_uri()}/{first}" in res.stderr
    res = cmd(repo, home, mirror, "status")
    assert res.returncode == 0 and "verified against the lock" in res.stderr


def test_install_refuses_a_changed_release_file(release, tmp_path):
    _, mirror, lock = release
    changed = Path(shutil.copytree(mirror, tmp_path / "changed"))
    cli = next(p for p in sorted(changed.rglob("*")) if p.is_file() and p.name != "SHA256SUMS"
               and not p.name.endswith(".whl"))  # a CLI, or a CLI archive
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
    for key in ("artifacts", "installed_files", "downloads"):
        partial.get(key, {}).pop(PLAT, None)
    partial["versions"] = {}
    repo = repo_copy(tmp_path / "repo", partial)
    res = cmd(repo, tmp_path / "home", mirror, "lock", "--write", "--release", "--skip-attestation")
    assert res.returncode == 0, res.stderr
    assert json.loads((repo / "rdocx.lock.json").read_text()) == lock


def test_lock_write_release_refuses_files_that_differ_from_its_listing(release, tmp_path):
    _, mirror, lock = release
    changed = Path(shutil.copytree(mirror, tmp_path / "changed"))
    with open(next(p for p in changed.rglob("*.whl") if "rpptx" in p.name), "ab") as f:
        f.write(b"\0")
    repo = repo_copy(tmp_path / "repo", lock)
    res = cmd(repo, tmp_path / "home", changed, "lock", "--write", "--release", "--skip-attestation")
    assert res.returncode == 1 and "do not match the release's SHA256SUMS" in res.stderr
    assert json.loads((repo / "rdocx.lock.json").read_text()) == lock


def test_lock_write_release_needs_gh_to_check_provenance(upstream, tmp_path, monkeypatch):
    mirror, lock = upstream
    repo = repo_copy(tmp_path / "repo", lock)
    monkeypatch.setenv("PATH", str(tmp_path))  # no gh there
    res = cmd(repo, tmp_path / "home", mirror, "lock", "--write", "--release")
    assert res.returncode == 1 and "needs the GitHub CLI (gh)" in res.stderr
    assert json.loads((repo / "rdocx.lock.json").read_text()) == lock


def test_bump_to_a_commit_empties_the_hashes(tmp_path):
    repo, other = repo_copy(tmp_path / "repo"), "0123456789abcdef0123456789abcdef01234567"
    res = cmd(repo, tmp_path / "home", tmp_path, "bump", other)
    assert res.returncode == 0, res.stderr
    lock = json.loads((repo / "rdocx.lock.json").read_text())
    assert lock["commit"] == other and lock["ref"] == other
    assert lock["release"] == "" and "rdocx-<date>-0123456789ab" in res.stderr
    assert lock["artifacts"] == lock["installed_files"] == lock["versions"] == {} and "downloads" not in lock
    res = cmd(repo, tmp_path / "home", tmp_path, "bump", other)
    assert res.returncode == 0 and "already pinned" in res.stderr
    assert json.loads((repo / "rdocx.lock.json").read_text()) == lock


def test_upstream_tags_are_the_family_tags_at_the_commit():
    commit, other = "a" * 40, "b" * 40
    out = "".join(f"{oid}\trefs/tags/{name}\n" for oid, name in [
        ("c" * 40, "v0.15.0"), (commit, "v0.15.0^{}"),  # annotated
        (commit, "rpptx-v0.13.1"),  # lightweight
        (commit, "py-rdocx-v0.15.0"), (commit, "s88"),  # other tags at the commit
        (other, "v0.14.0"), (other, "rpptx-v0.12.1"),
    ])
    assert rdocx_env.upstream_tags(out, commit) == {"rdocx": "v0.15.0", "rpptx": "rpptx-v0.13.1"}
    assert rdocx_env.upstream_tags(out, other) == {"rdocx": "v0.14.0", "rpptx": "rpptx-v0.12.1"}
    assert rdocx_env.upstream_tags(out, "d" * 40) == {}


def test_upstream_assets_pick_the_archive_and_the_wheel_of_a_platform(capsys):
    sums = {n: "0" * 64 for n in ["rdocx-0.15.0-cp39-abi3-manylinux_2_28_x86_64.whl",
                                   "rdocx-0.15.0-cp39-abi3-musllinux_1_2_x86_64.whl",
                                   "rdocx-0.15.0-cp39-abi3-win_amd64.whl", "rdocx-0.15.0.tar.gz",
                                   "rdocx-x86_64-unknown-linux-gnu.tar.gz", "rdocx-x86_64-unknown-linux-musl.tar.gz",
                                   "rdocx-x86_64-pc-windows-msvc.zip"]}
    assert rdocx_env.upstream_assets("rdocx", sums, "linux-x86_64") == [
        "rdocx-x86_64-unknown-linux-gnu.tar.gz", "rdocx-0.15.0-cp39-abi3-manylinux_2_28_x86_64.whl"]
    assert rdocx_env.upstream_assets("rdocx", sums, "windows-x86_64") == [
        "rdocx-x86_64-pc-windows-msvc.zip", "rdocx-0.15.0-cp39-abi3-win_amd64.whl"]
    assert rdocx_env.upstream_assets("rdocx", sums, "macos-arm64") == []
    del sums["rdocx-0.15.0-cp39-abi3-win_amd64.whl"]
    with pytest.raises(SystemExit):
        rdocx_env.upstream_assets("rdocx", sums, "windows-x86_64")
    assert "expected one of each" in capsys.readouterr().err


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
