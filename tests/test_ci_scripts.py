""".github/scripts: suite-summary.sh groups a suite log for the release notes, release_notes.py writes the final
suite of main into the notes of the release the lock pins (only one of this repository's own builds) and keeps the
build-time results below it, plugin_release.py names the builds and the plugin's releases and writes their version.
Synthetic notes, logs and locks only; no network, no release touched."""
import datetime
import json
import subprocess
import sys
from pathlib import Path

import pytest

from conftest import BASH

ROOT = Path(__file__).resolve().parent.parent
SCRIPTS = ROOT / ".github" / "scripts"
sys.path.insert(0, str(SCRIPTS))
import plugin_release  # noqa: E402
import release_notes  # noqa: E402

REPO = "hadim/rdocx-skills"
COMMIT = "ffb8d3d67f420d3554a74fd49d206b99e5a993cd"
BUILD_SHA = "cbd27d520ab378b61655302b28267aa50d7829fc"
MAIN_SHA = "0123456789abcdef0123456789abcdef01234567"
RUN = "https://github.com/hadim/rdocx-skills/actions/runs/1"
PREAMBLE = (f"hadim/rdocx@{COMMIT} ({COMMIT}), built for linux-x86_64 by https://github.com/{REPO}/actions/runs/9 "
            f"from {REPO}@{BUILD_SHA}.\n\nThese files are trusted only through rdocx.lock.json.\n")
LEGACY = PREAMBLE + "\n## Acceptance suite\n### linux-x86_64\n\n`1 failed, 9 passed`\n\nOther failures:\n\n```\nFAILED x\n```\n"
ADDED = "\n## Added later: macos-x86_64\n\nbuilt later\n\n### macos-x86_64\n\n`10 passed`\n"
SUMMARY_A = "### linux-x86_64\n\n`10 passed in 1s`\n"
SUMMARY_B = "### linux-x86_64\n\n`9 passed, 1 failed in 1s`\n"


@pytest.mark.parametrize("stamp", ["202610121405", "20261012"])  # build date and time; date only (older)
def test_only_this_repositorys_own_builds_are_named(stamp):
    own = f"https://github.com/{REPO}/releases/download/rdocx-{stamp}-{COMMIT[:12]}"
    assert release_notes.own_release_tag(own, REPO, COMMIT) == f"rdocx-{stamp}-{COMMIT[:12]}"
    assert release_notes.own_release_tag(own + "/", REPO, COMMIT) == f"rdocx-{stamp}-{COMMIT[:12]}"
    upstream = {"rdocx": "https://github.com/tensorbee/rdocx/releases/tag/v0.16.0"}
    assert release_notes.own_release_tag(upstream, REPO, COMMIT) is None
    assert release_notes.own_release_tag(own.replace(REPO, "tensorbee/rdocx"), REPO, COMMIT) is None
    assert release_notes.own_release_tag(own, REPO, "f" * 40) is None  # a release of another commit
    assert release_notes.own_release_tag(f"https://github.com/{REPO}/releases/download/v0.16.0", REPO, COMMIT) is None
    assert release_notes.own_release_tag("", REPO, COMMIT) is None


def test_a_release_built_without_a_suite_gets_the_suite_after_its_preamble():
    """build.yml no longer runs the suite: the notes of a new release are a preamble only."""
    body = PREAMBLE + "\nThe acceptance suite runs in ci.yml on the pull request that pins this release.\n"
    out = release_notes.rewrite(body, [SUMMARY_A], MAIN_SHA, RUN)
    assert out.startswith(PREAMBLE.rstrip("\n"))
    assert out.count("## Acceptance suite (skills main") == 1 and "## Build-time suite" not in out
    assert out.index("pins this release.") < out.index("## Acceptance suite (skills main") < out.index("### linux-x86_64")
    again = release_notes.rewrite(out, [SUMMARY_B], MAIN_SHA, RUN)
    assert again.count("## Acceptance suite") == 1 and "9 passed, 1 failed" in again


@pytest.mark.parametrize("stamp", ["2026101214", "2026101214051", "2026-10-12", "2026101"])
def test_a_tag_stamped_otherwise_is_not_named(stamp):
    url = f"https://github.com/{REPO}/releases/download/rdocx-{stamp}-{COMMIT[:12]}"
    assert release_notes.own_release_tag(url, REPO, COMMIT) is None


def test_the_final_suite_heads_the_notes_and_the_build_time_suite_is_kept():
    out = release_notes.rewrite(LEGACY + ADDED, [SUMMARY_A], MAIN_SHA, RUN)
    final = f"## Acceptance suite (skills main {MAIN_SHA[:12]}, {RUN})"
    build = f"## Build-time suite (preliminary, skills {BUILD_SHA[:12]})"
    assert out.startswith(PREAMBLE.rstrip("\n") + "\n\n" + final + "\n\n" + SUMMARY_A)
    assert "\n## Acceptance suite\n" not in out
    assert out.index(final) < out.index(build) < out.index("## Added later: macos-x86_64")
    assert "FAILED x" in out.split(build)[1]  # the build-time results, unchanged
    assert "`10 passed`" in out.split("## Added later")[1]


def test_rewriting_again_replaces_the_final_suite():
    once = release_notes.rewrite(LEGACY + ADDED, [SUMMARY_A], MAIN_SHA, RUN)
    again = release_notes.rewrite(once, [SUMMARY_B], "f" * 40, RUN + "0")
    assert again == release_notes.rewrite(LEGACY + ADDED, [SUMMARY_B], "f" * 40, RUN + "0")
    assert again.count("## Acceptance suite") == 1 and again.count("## Build-time suite") == 1
    assert SUMMARY_A not in again
    assert release_notes.rewrite(once + "\n", [SUMMARY_A], MAIN_SHA, RUN) == once  # gh's trailing newline


def test_notes_of_the_new_build_format_keep_their_heading():
    body = PREAMBLE + f"\n## Build-time suite (preliminary, skills {BUILD_SHA[:12]})\n\nRun early.\n\n" + SUMMARY_B
    out = release_notes.rewrite(body, [SUMMARY_A], MAIN_SHA, RUN)
    assert out.count("## Build-time suite") == 1
    assert out.index("## Acceptance suite (skills main") < out.index("## Build-time suite")


def test_the_command_writes_the_notes_to_stdout(tmp_path):
    notes, summary = tmp_path / "notes.md", tmp_path / "suite-linux-x86_64.md"
    notes.write_text(LEGACY, encoding="utf-8")
    summary.write_text(SUMMARY_A, encoding="utf-8")
    res = subprocess.run([sys.executable, SCRIPTS / "release_notes.py", "rewrite", "--notes", notes, "--sha", MAIN_SHA,
                          "--run-url", RUN, summary], capture_output=True, text=True, encoding="utf-8")
    assert res.returncode == 0, res.stderr
    assert res.stdout == release_notes.rewrite(LEGACY, [SUMMARY_A], MAIN_SHA, RUN)
    assert notes.read_text(encoding="utf-8") == LEGACY  # the input is left as it was


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


def test_the_command_takes_the_summaries_after_the_options_as_ci_yml_passes_them(tmp_path):
    notes = tmp_path / "notes.md"
    notes.write_text(LEGACY, encoding="utf-8")
    summaries = []
    for name, text in (("suite-linux-x86_64.md", SUMMARY_A), ("suite-macos-arm64.md", SUMMARY_B)):
        path = tmp_path / name
        path.write_text(text, encoding="utf-8")
        summaries.append(path)
    res = subprocess.run([sys.executable, SCRIPTS / "release_notes.py", "rewrite", "--notes", notes, "--sha", MAIN_SHA,
                          "--run-url", RUN, *summaries], capture_output=True, text=True, encoding="utf-8")
    assert res.returncode == 0, res.stderr
    assert res.stdout == release_notes.rewrite(LEGACY, [SUMMARY_A, SUMMARY_B], MAIN_SHA, RUN)


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


def test_the_release_notes_name_the_pinned_build_and_how_to_install(tmp_path):
    lock = {"upstream": "https://github.com/hadim/rdocx", "commit": COMMIT, "ref": "integration/x",
            "release": f"https://github.com/{REPO}/releases/download/rdocx-202610141943-{COMMIT[:12]}",
            "versions": {"rdocx": "0.16.0", "rpptx": "0.14.0"}}
    (tmp_path / "lock.json").write_text(json.dumps(lock))
    res = subprocess.run([sys.executable, SCRIPTS / "plugin_release.py", "notes", "--repo", REPO, "--version",
                          "2026.10.5", tmp_path / "lock.json"], capture_output=True, text=True, encoding="utf-8")
    assert res.returncode == 0, res.stderr
    assert res.stdout == plugin_release.notes(lock, REPO, "2026.10.5")
    assert "rdocx 0.16.0 and rpptx 0.14.0" in res.stdout and "(`integration/x`)" in res.stdout
    assert f"/releases/tag/rdocx-202610141943-{COMMIT[:12]})" in res.stdout
    assert f"`/plugin marketplace add {REPO}@release`" in res.stdout
