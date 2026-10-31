"""The checks an agent can run on a .docx / .pptx before handing it over, each as that stack runs it.

Every checker takes a file path and returns ``{"flagged": bool, "message": str, "seconds": float}``: flagged when
the check reports a problem or fails to read the file (an exception, a non-zero exit, a timeout, no output).
``seconds`` is the wall-clock time of the command, process start included.

- ``rdocx_validate``: ``rdocx validate F`` / ``rpptx validate F``, flagged on a non-zero exit (warnings exit 0).
- ``rdocx_open``: the rdocx / rpptx Python binding loads the file and reads its text (``probe_rdocx.py``).
- ``builtin_validate``: the built-in skill's ``scripts/office/validate.py F``, as the skill runs it on a file built
  from scratch (no ``--original``), flagged on a non-zero exit.
- ``builtin_validate_original``: the same with ``--original BASE``, as the skill's edit flow runs it on an edited
  file (the base the mutation started from; on a base itself, the base). Only where an original exists.
- ``pylib_open``: python-docx / python-pptx load the file and iterate it (``probe_pylib.py``).
- ``libreoffice``: PDF conversion through the skill's ``scripts/office/soffice.py``, flagged when no PDF comes out
  (LibreOffice repairs what it can silently, so a PDF says only that it could read something).
- ``pandoc``: ``pandoc -t markdown F`` (docx only), flagged on a non-zero exit.
"""

from __future__ import annotations

import sys
import tempfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parent))
import common  # noqa: E402

CHECKERS = {
    "docx": ["rdocx_validate", "rdocx_open", "builtin_validate", "builtin_validate_original", "pylib_open",
             "libreoffice", "pandoc"],
    "pptx": ["rdocx_validate", "rdocx_open", "builtin_validate", "builtin_validate_original", "pylib_open",
             "libreoffice"],
}
NEEDS_ORIGINAL = {"builtin_validate_original"}

LABELS = {
    "rdocx_validate": "rdocx/rpptx validate (CLI)",
    "rdocx_open": "rdocx/rpptx Python open + text",
    "builtin_validate": "built-in scripts/office/validate.py",
    "builtin_validate_original": "built-in validate.py --original (the skill's edit flow)",
    "pylib_open": "python-docx/python-pptx open + iterate",
    "libreoffice": "LibreOffice PDF conversion (soffice.py)",
    "pandoc": "pandoc -t markdown",
}

TIMEOUT = 180


def _short(text: str, limit=400) -> str:
    lines = [ln.strip() for ln in text.strip().splitlines() if ln.strip()]
    return " | ".join(lines)[:limit]


def _result(r, flagged, message):
    return {"flagged": bool(flagged), "message": _short(message), "seconds": round(r["seconds"], 3)}


def fmt_of(path: Path) -> str:
    return "docx" if path.suffix.lower() in (".docx", ".dotx") else "pptx"


def rdocx_validate(path: Path):
    cli = common.RDOCX if fmt_of(path) == "docx" else common.RPPTX
    r = common.run([cli, "validate", path], env=common.env_rdocx(), timeout=TIMEOUT)
    return _result(r, r["rc"] != 0, f"exit {r['rc']}: {r['stdout']} {r['stderr']}")


def rdocx_open(path: Path):
    r = common.run([common.R_PY, HERE / "probe_rdocx.py", path], env=common.env_rdocx(), timeout=TIMEOUT)
    return _result(r, r["rc"] != 0, (r["stdout"] + " " + r["stderr"]).strip() or f"exit {r['rc']}")


def builtin_validate(path: Path):
    skill = common.BUILTIN_DOCX if fmt_of(path) == "docx" else common.BUILTIN_PPTX
    r = common.run([common.PY, skill / "scripts" / "office" / "validate.py", path], env=common.env_builtin(),
                   timeout=TIMEOUT)
    return _result(r, r["rc"] != 0, f"exit {r['rc']}: {r['stdout']} {r['stderr']}")


def builtin_validate_original(path: Path, original: Path):
    skill = common.BUILTIN_DOCX if fmt_of(path) == "docx" else common.BUILTIN_PPTX
    r = common.run([common.PY, skill / "scripts" / "office" / "validate.py", path, "--original", original],
                   env=common.env_builtin(), timeout=TIMEOUT)
    return _result(r, r["rc"] != 0, f"exit {r['rc']}: {r['stdout']} {r['stderr']}")


def pylib_open(path: Path):
    r = common.run([common.PY, HERE / "probe_pylib.py", path], env=common.env_pylib(), timeout=TIMEOUT)
    return _result(r, r["rc"] != 0, (r["stdout"] + " " + r["stderr"]).strip() or f"exit {r['rc']}")


def libreoffice(path: Path):
    skill = common.BUILTIN_DOCX if fmt_of(path) == "docx" else common.BUILTIN_PPTX
    with tempfile.TemporaryDirectory(prefix="lo-", dir=common.WORK) as out:
        r = common.run([common.PY, skill / "scripts" / "office" / "soffice.py", "--headless", "--convert-to", "pdf",
                        "--outdir", out, path], env=common.env_builtin(), timeout=TIMEOUT)
        pdf = Path(out) / (path.stem + ".pdf")
        ok = pdf.exists() and pdf.stat().st_size > 0
        msg = "pdf written" if ok else f"no pdf (exit {r['rc']}): {r['stdout']} {r['stderr']}"
        if ok and "Error" in (r["stdout"] + r["stderr"]):
            msg += f" ({_short(r['stdout'] + r['stderr'], 200)})"
        return _result(r, not ok, msg)


def pandoc(path: Path):
    r = common.run([common.PANDOC, "-f", "docx", "-t", "markdown", path], env=common.env_builtin(), timeout=TIMEOUT)
    msg = f"exit {r['rc']}: {r['stderr']}" if r["rc"] else f"exit 0, {len(r['stdout'])} chars of markdown"
    return _result(r, r["rc"] != 0, msg)


FUNCS = {f.__name__: f for f in (rdocx_validate, rdocx_open, builtin_validate, builtin_validate_original,
                                  pylib_open, libreoffice, pandoc)}


def check(name: str, path: Path, original: Path | None = None):
    """Run one checker; ``None`` when it needs an original file and none is given."""
    if name in NEEDS_ORIGINAL:
        return None if original is None else FUNCS[name](Path(path), Path(original))
    return FUNCS[name](Path(path))
