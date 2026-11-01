"""Shared paths, tool locations and timing helpers for the comparison benchmark.

Every section script imports this module, so that the three stacks run with the same
tools, the same fonts and the same inputs:

- ``rdocx``: the rdocx-skills plugin, i.e. the pinned rdocx and rpptx build
  (CLIs and Python bindings) plus ``docx_ops.py`` and ``pptx_ops.py``.
- ``builtin``: Claude's built-in ``docx`` and ``pptx`` skills as they document themselves:
  docx-js and pptxgenjs to create, unzip / edit XML / zip to edit, their helper scripts
  (``merge_runs.py``, ``comment.py``, ``accept_changes.py``, ``add_slide.py``, ``clean.py``,
  ``slide_text.py``, ``thumbnail.py``, ``office/validate.py``), pandoc to read, LibreOffice
  and pdftoppm to render. The skill's scripts are proprietary: they are used from the
  installed skill folder (``BUILTIN_SKILLS_DIR``) and never copied into this repository.
- ``pylib``: python-docx and python-pptx alone (plus lxml when the library has no API), what
  an agent writes without a skill, and the path the built-in pptx skill mentions.
"""

from __future__ import annotations

import glob
import json
import os
import platform
import shutil
import statistics
import subprocess
import sys
import time
from pathlib import Path

BENCH = Path(__file__).resolve().parent
REPO = BENCH.parent
CACHE = BENCH / ".cache"
RESULTS = BENCH / "results"
WORK = CACHE / "work"

# --- rdocx-skills stack -------------------------------------------------------------------
RDOCX_HOME = Path(os.environ.get("RDOCX_HOME", Path.home() / ".local/share/rdocx-skills"))
R = RDOCX_HOME / "current" / "bin"
RDOCX = R / "rdocx"
RPPTX = R / "rpptx"
R_PY = R / "python"  # the Python that has rdocx and rpptx
DOCX_OPS_DIR = REPO / "plugin" / "skills" / "docx" / "scripts"
PPTX_OPS_DIR = REPO / "plugin" / "skills" / "pptx" / "scripts"

# --- built-in skills stack ----------------------------------------------------------------
def _find_builtin_skills() -> Path | None:
    env = os.environ.get("BUILTIN_SKILLS_DIR")
    if env:
        return Path(env)
    # Claude Code syncs the account's skills here; take the most recent copy that has both.
    candidates = sorted(
        (Path(p).parent.parent for p in glob.glob(str(Path.home() / ".claude/skills/synced/*/docx/SKILL.md"))),
        key=lambda p: (p / "docx" / "SKILL.md").stat().st_mtime,
        reverse=True,
    )
    for c in candidates:
        if (c / "pptx" / "SKILL.md").exists():
            return c
    return None


BUILTIN_SKILLS_DIR = _find_builtin_skills()
BUILTIN_DOCX = BUILTIN_SKILLS_DIR / "docx" if BUILTIN_SKILLS_DIR else None
BUILTIN_PPTX = BUILTIN_SKILLS_DIR / "pptx" if BUILTIN_SKILLS_DIR else None

PY = CACHE / "venv" / "bin" / "python"  # python-docx 1.2.0, python-pptx 1.0.2, lxml, defusedxml, Pillow
NODE = shutil.which("node") or "node"
NODE_MODULES = CACHE / "node" / "node_modules"  # docx (docx-js), pptxgenjs
PANDOC = CACHE / "bin" / "pandoc"
PDFTOPPM = shutil.which("pdftoppm") or "pdftoppm"
SOFFICE = shutil.which("soffice") or "soffice"

# --- fonts --------------------------------------------------------------------------------
# Carlito and Caladea are metric-compatible with Calibri and Cambria. Both renderers get them:
# LibreOffice through FONTCONFIG_FILE, rdocx and rpptx through --font-dir / font_dir=.
FONTS_DIR = CACHE / "fonts"
FONTCONFIG_FILE = CACHE / "fonts.conf"


def env_builtin() -> dict:
    """Environment for the built-in stack: its Python, node modules, fonts."""
    env = os.environ.copy()
    env["FONTCONFIG_FILE"] = str(FONTCONFIG_FILE)
    env["NODE_PATH"] = str(NODE_MODULES)
    env["PATH"] = f"{CACHE / 'bin'}:{PY.parent}:{env['PATH']}"
    env["PYTHONDONTWRITEBYTECODE"] = "1"
    return env


def env_rdocx() -> dict:
    env = os.environ.copy()
    env["FONTCONFIG_FILE"] = str(FONTCONFIG_FILE)
    env["PYTHONPATH"] = f"{DOCX_OPS_DIR}:{PPTX_OPS_DIR}"
    env["PYTHONDONTWRITEBYTECODE"] = "1"
    env["PATH"] = f"{R}:{env['PATH']}"
    return env


def env_pylib() -> dict:
    env = os.environ.copy()
    env["PYTHONDONTWRITEBYTECODE"] = "1"
    env["PATH"] = f"{PY.parent}:{env['PATH']}"
    return env


# --- running and timing -------------------------------------------------------------------
def run(cmd, *, env=None, cwd=None, timeout=900, check=False, input=None):
    """Run a command; return a dict with rc, stdout, stderr and wall seconds."""
    t0 = time.perf_counter()
    try:
        p = subprocess.run(
            [str(c) for c in cmd], env=env, cwd=cwd, capture_output=True, text=True,
            timeout=timeout, input=input,
        )
        rc, out, err = p.returncode, p.stdout, p.stderr
    except subprocess.TimeoutExpired as e:
        rc, out, err = -9, e.stdout or "", f"timeout after {timeout}s"
        if isinstance(out, bytes):
            out = out.decode(errors="replace")
    dt = time.perf_counter() - t0
    if check and rc != 0:
        raise RuntimeError(f"{cmd} failed ({rc}):\n{out[-2000:]}\n{err[-2000:]}")
    return {"rc": rc, "stdout": out, "stderr": err, "seconds": dt}


def timeit(make_cmd, *, repeat=5, warmup=1, env=None, cwd=None, setup=None, timeout=900):
    """Time a command as an agent runs it (process start included).

    ``make_cmd(i)`` returns the argv of run i (so that each run writes a fresh output);
    ``setup(i)`` runs before each run, untimed. Returns median, min, max and every sample.
    """
    samples, last = [], None
    for i in range(warmup + repeat):
        if setup:
            setup(i)
        last = run(make_cmd(i), env=env, cwd=cwd, timeout=timeout)
        if last["rc"] != 0:
            return {"ok": False, "rc": last["rc"], "error": (last["stderr"] or last["stdout"])[-1500:]}
        if i >= warmup:
            samples.append(last["seconds"])
    return {
        "ok": True,
        "median": statistics.median(samples),
        "min": min(samples),
        "max": max(samples),
        "samples": samples,
    }


def fresh_dir(*parts) -> Path:
    d = WORK.joinpath(*parts)
    if d.exists():
        shutil.rmtree(d)
    d.mkdir(parents=True)
    return d


# --- versions and results -----------------------------------------------------------------
def versions() -> dict:
    def out(cmd, env=None):
        r = run(cmd, env=env, timeout=120)
        return (r["stdout"] or r["stderr"]).strip().splitlines()[0] if r["rc"] == 0 else f"error {r['rc']}"

    lock = json.loads((REPO / "plugin" / "rdocx.lock.json").read_text())
    return {
        "date": time.strftime("%Y-%m-%d"),
        "machine": f"{platform.system()} {platform.machine()}, {os.cpu_count()} CPUs",
        "rdocx_pin": lock.get("commit", "")[:12],
        "rdocx": out([RDOCX, "--version"]),
        "rpptx": out([RPPTX, "--version"]),
        "python-docx": out([PY, "-c", "import docx; print(docx.__version__)"]),
        "python-pptx": out([PY, "-c", "import pptx; print(pptx.__version__)"]),
        "docx-js": json.loads((NODE_MODULES / "docx" / "package.json").read_text())["version"],
        "pptxgenjs": json.loads((NODE_MODULES / "pptxgenjs" / "package.json").read_text())["version"],
        "node": out([NODE, "--version"]),
        "pandoc": out([PANDOC, "--version"]),
        "libreoffice": out([SOFFICE, "--version"]),
        "pdftoppm": out([PDFTOPPM, "-v"]),
        "builtin_skills_sha256": builtin_skills_digest(),
    }


def builtin_skills_digest() -> str:
    """SHA-256 over the built-in docx and pptx skills' files: they carry no version number."""
    import hashlib

    if not BUILTIN_SKILLS_DIR:
        return "not found"
    h = hashlib.sha256()
    for skill in ("docx", "pptx"):
        for f in sorted((BUILTIN_SKILLS_DIR / skill).rglob("*")):
            if f.is_file() and "__pycache__" not in f.parts:
                h.update(f.relative_to(BUILTIN_SKILLS_DIR).as_posix().encode() + b"\0" + f.read_bytes())
    return h.hexdigest()[:16]


def scrub(text: str) -> str:
    """Replace machine-specific paths in results with placeholders (the repository is public)."""
    pairs = [
        (str(BUILTIN_SKILLS_DIR) if BUILTIN_SKILLS_DIR else None, "$BUILTIN_SKILLS_DIR"),
        (str(BENCH), "$BENCH"),
        (str(RDOCX_HOME), "$RDOCX_HOME"),
        (str(Path.home()), "~"),
    ]
    for old, new in pairs:
        if old:
            text = text.replace(old, new)
    return text


def save_results(name: str, data) -> Path:
    RESULTS.mkdir(parents=True, exist_ok=True)
    p = RESULTS / f"{name}.json"
    p.write_text(scrub(json.dumps(data, indent=2, ensure_ascii=False)) + "\n")
    return p


if __name__ == "__main__":
    json.dump(versions(), sys.stdout, indent=2)
    print()
