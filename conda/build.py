"""Repackage a release's files (conda/recipe.yaml): the CLIs into the environment's bin folder, the wheels into
its site-packages, as they are."""
import os
import shutil
import subprocess
import sys
from pathlib import Path

windows = sys.platform == "win32"
bin_dir = Path(os.environ["LIBRARY_BIN"] if windows else Path(os.environ["PREFIX"]) / "bin")
bin_dir.mkdir(parents=True, exist_ok=True)
for cli in ("rdocx", "rpptx"):
    name = cli + (".exe" if windows else "")
    shutil.copy(name, bin_dir / name)
    os.chmod(bin_dir / name, 0o755)
wheels = sorted(str(p) for p in Path.cwd().glob("*.whl"))
if len(wheels) != 2:
    sys.exit(f"expected the rdocx and rpptx wheels, found {wheels}")
subprocess.run([sys.executable, "-m", "pip", "install", "--no-deps", "--no-index", "--no-build-isolation",
                "--no-compile", "--disable-pip-version-check", *wheels], check=True)
