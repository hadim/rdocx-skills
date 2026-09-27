"""Shared fixtures of the acceptance suite. Run it with `python3 scripts/rdocx_env.py test`, which uses the
pinned, verified build; or directly with a Python that has rdocx and rpptx, with RDOCX_BIN_DIR pointing at
the folder that holds the two CLIs."""
import os
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).parent))
from gaps import GAPS  # noqa: E402

TESTS = Path(__file__).parent
BIN = Path(os.environ.get("RDOCX_BIN_DIR", Path(os.environ.get("RDOCX_HOME", Path.home() / ".local/share/rdocx-skills")) / "current" / "bin"))
STAMP = "2026-09-27T12:00:00Z"


def pytest_configure(config):
    config.addinivalue_line("markers", "gap(key): known gap of the pinned build, see tests/gaps.py (strict xfail)")


def pytest_collection_modifyitems(config, items):
    for item in items:
        for mark in item.iter_markers("gap"):
            key = mark.args[0]
            if key not in GAPS:
                raise pytest.UsageError(f"{item.nodeid}: unknown gap key {key!r}")
            summary, ticket = GAPS[key]
            item.add_marker(pytest.mark.xfail(strict=True, reason=f"gap {key} ({ticket}): {summary}"))


def tool(name):
    """The pinned CLI only: a test never runs another build found on PATH."""
    path = BIN / name
    if not path.exists():
        pytest.fail(f"{path} not found: install the pinned build (rdocx_env.py install) or set RDOCX_BIN_DIR")
    return str(path)


@pytest.fixture(scope="session")
def rdocx_cli():
    return tool("rdocx")


@pytest.fixture(scope="session")
def rpptx_cli():
    return tool("rpptx")


def run(cmd, check=False, **kw):
    return subprocess.run([str(c) for c in cmd], capture_output=True, text=True, check=check, **kw)


@pytest.fixture(scope="session")
def fixture_dir(tmp_path_factory):
    return tmp_path_factory.mktemp("fixtures")


@pytest.fixture(scope="session")
def report_docx(fixture_dir):
    """The synthetic report: Google Docs export then Word edit traits, 13 pages in LibreOffice."""
    out = fixture_dir / "fixture-report.docx"
    run([sys.executable, TESTS / "fixtures" / "make_fixture_docx.py", out], check=True)
    return out


@pytest.fixture(scope="session")
def deck_pptx(fixture_dir):
    """The synthetic deck: 7 slides, 4:3, groups, connectors, pictures, table, notes, one overflowing box."""
    out = fixture_dir / "fixture-deck.pptx"
    run([sys.executable, TESTS / "fixtures" / "make_fixture_pptx.py", out], check=True)
    return out


@pytest.fixture
def copy_of(tmp_path):
    def _copy(src, name=None):
        dst = tmp_path / (name or Path(src).name)
        shutil.copy2(src, dst)
        return dst
    return _copy


def part(path, name):
    with zipfile.ZipFile(path) as z:
        return z.read(name)


def parts(path):
    with zipfile.ZipFile(path) as z:
        return {n: z.read(n) for n in z.namelist()}
