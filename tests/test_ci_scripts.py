""".github/scripts: suite-summary.sh groups a suite log for the job summaries and the release notes,
plugin_release.py names the plugin's releases, writes their version and the head of their notes, tested-tree.sh tells
whether ci.yml passed on a tree. Synthetic logs, locks and API answers only; no network, no release touched."""
import datetime
import json
import os
import shutil
import subprocess
import sys
from pathlib import Path

import pytest

from conftest import BASH

ROOT = Path(__file__).resolve().parent.parent
SCRIPTS = ROOT / ".github" / "scripts"
sys.path.insert(0, str(SCRIPTS))
import plugin_release  # noqa: E402

REPO = "hadim/rdocx-skills"
COMMIT = "ffb8d3d67f420d3554a74fd49d206b99e5a993cd"
RUN = "https://github.com/hadim/rdocx-skills/actions/runs/1"


def test_the_summary_groups_docs_failures_apart(tmp_path):
    log = tmp_path / "pytest.log"
    log.write_text(
        "FAILED tests/test_pptx.py::test_closed - [XPASS(strict)] gap render-embedded-font\n"
        "FAILED tests/test_skill_docs.py::test_every_python_api_is_documented_or_excluded[rdocx] - AssertionError\n"
        "ERROR tests/test_skill_docs.py::test_setup - RuntimeError\n"
        "FAILED tests/test_docx_edit.py::test_replace - AssertionError\n"
        "==== 3 failed, 1 error, 9 passed in 1.00s ====\n", encoding="utf-8")
    out = subprocess.run([BASH, SCRIPTS / "suite-summary.sh", log, "linux-x86_64"], capture_output=True, text=True,
                         check=True).stdout
    assert out.startswith("### linux-x86_64\n\n`3 failed, 1 error, 9 passed in 1.00s`\n")
    closed, rest = out.split("Gaps closed")[1].split("Docs not yet updated for the new build")
    docs, other = rest.split("Other failures (regressions or changed APIs)")
    assert "test_closed" in closed and "test_skill_docs" not in closed
    assert docs.count("tests/test_skill_docs.py::") == 2 and "test_replace" not in docs
    assert "test_replace" in other and "test_skill_docs" not in other and "XPASS" not in other


def test_the_summary_without_docs_failures_has_no_docs_group(tmp_path):
    log = tmp_path / "pytest.log"
    log.write_text("FAILED tests/test_docx_edit.py::test_replace - AssertionError\n==== 1 failed ====\n",
                   encoding="utf-8")
    out = subprocess.run([BASH, SCRIPTS / "suite-summary.sh", log, "macos-arm64"], capture_output=True, text=True,
                         check=True).stdout
    assert "Docs not yet updated" not in out and "Gaps closed" not in out
    assert "Other failures" in out and "test_replace" in out


@pytest.mark.parametrize("ref, repo, label", [
    ("v0.16.1", "tensorbee/rdocx", "rdocx 0.16.1 \u00b7 rpptx 0.14.1"),  # an upstream release: the versions say it all
    ("rpptx-v0.14.1", "tensorbee/rdocx", "rdocx 0.16.1 \u00b7 rpptx 0.14.1"),
    ("v0.16.1", "hadim/rdocx", "rdocx 0.16.1 \u00b7 rpptx 0.14.1"),  # the same tag, mirrored on a fork
    ("main", "tensorbee/rdocx", "rdocx 0.16.1 \u00b7 rpptx 0.14.1 + main"),
    ("integration/open-prs-2026-10-14", "hadim/rdocx",
     "rdocx 0.16.1 \u00b7 rpptx 0.14.1 + integration/open-prs-2026-10-14 from hadim/rdocx"),
    (COMMIT, "hadim/rdocx", f"rdocx 0.16.1 \u00b7 rpptx 0.14.1 + {COMMIT[:12]} from hadim/rdocx"),
    ("", "tensorbee/rdocx", "rdocx 0.16.1 \u00b7 rpptx 0.14.1"),
])
def test_a_build_is_named_after_its_versions_and_what_it_was_built_from(ref, repo, label):
    assert plugin_release.label({"rdocx": "0.16.1", "rpptx": "0.14.1"}, ref, repo) == label


def test_a_label_needs_both_versions():
    with pytest.raises(SystemExit):
        plugin_release.label({"rdocx": "0.16.1"}, "v0.16.1")


def test_the_version_is_the_day_then_a_counter_for_the_same_day():
    day = datetime.date(2026, 10, 5)
    assert plugin_release.next_version(day, []) == "2026.10.5"
    assert plugin_release.next_version(day, ["v2026.10.4", "rdocx-202610050000-0123456789ab"]) == "2026.10.5"
    assert plugin_release.next_version(day, ["v2026.10.5"]) == "2026.10.5.2"
    assert plugin_release.next_version(day, ["v2026.10.5", "v2026.10.5.2"]) == "2026.10.5.3"


def test_the_version_goes_into_plugin_json_after_the_name(tmp_path):
    manifest = tmp_path / "plugin.json"
    manifest.write_text('{\n  "name": "rdocx",\n  "description": "d",\n  "author": {"name": "A"}\n}\n')
    for version in ("2026.10.5", "2026.10.5.2"):  # a second release replaces the first version
        res = subprocess.run([sys.executable, SCRIPTS / "plugin_release.py", "set-version", manifest, version],
                             capture_output=True, text=True)
        assert res.returncode == 0, res.stderr
        assert manifest.read_text() == (f'{{\n  "name": "rdocx",\n  "version": "{version}",\n  "description": "d",\n'
                                        '  "author": {"name": "A"}\n}\n')
    assert json.loads(manifest.read_text())["version"] == "2026.10.5.2"


def test_the_release_notes_describe_the_release_s_own_files_and_how_to_install(tmp_path):
    lock = {"upstream": "https://github.com/tensorbee/rdocx", "commit": COMMIT, "ref": "main",
            "release": f"https://github.com/{REPO}/releases/download/v2026.10.5",
            "versions": {"rdocx": "0.16.1", "rpptx": "0.14.1"},
            "artifacts": {p: {} for p in ("windows-x86_64", "linux-x86_64", "macos-arm64", "linux-aarch64")}}
    (tmp_path / "lock.json").write_text(json.dumps(lock))
    res = subprocess.run([sys.executable, SCRIPTS / "plugin_release.py", "notes", "--repo", REPO, "--version",
                          "2026.10.5", "--run-url", RUN, tmp_path / "lock.json"],
                         capture_output=True, text=True, encoding="utf-8")
    assert res.returncode == 0, res.stderr
    assert res.stdout == plugin_release.notes(lock, REPO, "2026.10.5", RUN)
    assert "rdocx 0.16.1 and rpptx 0.14.1" in res.stdout and "(`main`)" in res.stdout
    assert "linux-aarch64, linux-x86_64, macos-arm64 and windows-x86_64" in res.stdout
    assert f"[This run]({RUN})" in res.stdout and f"`gh attestation verify FILE -R {REPO}`" in res.stdout
    assert f"`/plugin marketplace add {REPO}`, then" in res.stdout


def test_release_notes_need_recorded_files():
    lock = {"upstream": "https://github.com/tensorbee/rdocx", "commit": COMMIT,
            "versions": {"rdocx": "0.16.1", "rpptx": "0.14.1"}}
    with pytest.raises(SystemExit):
        plugin_release.notes(lock, REPO, "2026.10.5", RUN)


RUN_OK = {"path": ".github/workflows/ci.yml", "event": "pull_request", "conclusion": "success"}
MARK = {"expired": False, "workflow_run": {"id": 7, "repository_id": 1, "head_repository_id": 1}}


@pytest.mark.parametrize("artifacts,run,passed", [
    ([], RUN_OK, "false"),
    ([MARK], RUN_OK, "true"),
    ([MARK], {**RUN_OK, "event": "push"}, "true"),
    ([{**MARK, "expired": True}], RUN_OK, "false"),
    ([{**MARK, "workflow_run": {"id": 7, "repository_id": 1, "head_repository_id": 2}}], RUN_OK, "false"),
    ([MARK], {**RUN_OK, "path": ".github/workflows/other.yml"}, "false"),
    ([MARK], {**RUN_OK, "event": "workflow_dispatch"}, "false"),
    ([MARK], {**RUN_OK, "conclusion": "failure"}, "false"),
])
def test_a_tree_passed_only_on_a_live_mark_from_this_repository(tmp_path, artifacts, run, passed):
    """tested-tree.sh: the mark of a green ci.yml run (push or pull request) on HEAD's tree, from this repository,
    not expired."""
    repo = tmp_path / "repo"
    repo.mkdir()
    git = ["git", "-C", str(repo), "-c", "user.name=t", "-c", "user.email=t@t"]
    subprocess.run(["git", "init", "-q", str(repo)], check=True)
    (repo / "f").write_text("x\n")
    subprocess.run(git + ["add", "f"], check=True)
    subprocess.run(git + ["commit", "-qm", "c"], check=True)
    tree = subprocess.run(git + ["rev-parse", "HEAD^{tree}"], capture_output=True, text=True, check=True).stdout.strip()
    bin_dir = tmp_path / "bin"
    bin_dir.mkdir()
    (tmp_path / "artifacts.json").write_text(json.dumps({"artifacts": artifacts}))
    (tmp_path / "run.json").write_text(json.dumps(run))
    # a stand-in for `gh api URL --jq FILTER`: answers the artifact query for HEAD's tree and run 7, through jq
    gh = bin_dir / "gh"
    gh.write_text(f'#!/bin/sh\ncase "$2" in\n'
                  f'  *"/actions/artifacts?name=tested-tree-{tree}&"*) f="{(tmp_path / "artifacts.json").as_posix()}" ;;\n'
                  f'  */actions/runs/7) f="{(tmp_path / "run.json").as_posix()}" ;;\n'
                  f'  *) exit 3 ;;\nesac\nexec jq -r "$4" "$f"\n')
    gh.chmod(0o755)
    env = {**os.environ, "GITHUB_REPOSITORY": REPO, "PATH": f"{bin_dir}{os.pathsep}{os.environ['PATH']}"}
    res = subprocess.run([BASH, SCRIPTS / "tested-tree.sh"], cwd=repo, env=env, capture_output=True, text=True)
    assert res.returncode == 0, res.stderr
    assert res.stdout.splitlines() == [f"tree={tree}", f"passed={passed}"]
