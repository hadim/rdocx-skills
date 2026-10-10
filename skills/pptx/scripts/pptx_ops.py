"""Helpers for editing .pptx decks safely with rpptx. Import from a script run with the pinned Python
(`~/.local/share/rdocx-skills/current/bin/python`, with this folder on PYTHONPATH), or run as a command:

  python pptx_ops.py replace IN.pptx OUT.pptx --edit OLD NEW COUNT [--edit ...]   counted, all-or-nothing
  python pptx_ops.py overflow IN.pptx                                            text frames that overflow
                                                                   (`rpptx fit IN.pptx` from the CLI: exit 1 on overflow)
  python pptx_ops.py shapes IN.pptx [--slide N]                                  shape tree with geometry

`replace` chains `rpptx replace --expect` through temporary files and publishes the result only when every
count matched (in Python, `prs.try_replace_text(old, new, expect=n)` does one replacement in memory). Counts include speaker notes.
Outputs are written to a temporary file next to the target, flushed to disk and renamed, with the input's
file mode. Exit codes: 0 done, 1 refused (nothing written), 2 usage.
"""
import argparse
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

import rpptx

BIN = Path(os.environ.get("RDOCX_BIN_DIR", Path(os.environ.get("RDOCX_HOME", Path.home() / ".local/share/rdocx-skills")) / "current" / "bin"))


class EditError(RuntimeError):
    """A counted edit did not match: nothing was written."""


def cli(*args, check=True):
    """Run the pinned `rpptx` CLI (never another one found on PATH)."""
    exe = BIN / ("rpptx.exe" if os.name == "nt" else "rpptx")
    if not exe.exists():
        raise FileNotFoundError(f"{exe} not found: install the pinned build (rdocx_env.py install) or set RDOCX_BIN_DIR")
    res = subprocess.run([str(exe), *map(str, args)], capture_output=True, encoding="utf-8")  # the CLI writes UTF-8
    if check and res.returncode:
        raise RuntimeError((res.stdout + res.stderr).strip())
    return res


def _publish(write, out, src=None):
    """write(tmp) into a temporary file next to `out`, give it the mode of `src` (else the default mode under
    the umask), flush it to disk, rename it to `out`. Refuses `out == src`."""
    out = Path(out)
    if src is not None and Path(src).resolve() == out.resolve():
        raise EditError("refusing to write over the input; write to a new path, check it, then replace")
    fd, tmp = tempfile.mkstemp(prefix=f".{out.stem}.", suffix=out.suffix, dir=out.parent)
    os.close(fd)
    try:
        write(tmp)
        # flushed through a handle open for writing (Windows refuses to flush one open for reading, EBADF), before
        # setting the mode, which may make the file read-only
        with open(tmp, "rb+") as f:
            os.fsync(f.fileno())
        if src is not None and Path(src).exists():
            os.chmod(tmp, Path(src).stat().st_mode & 0o7777)
        else:
            umask = os.umask(0)
            os.umask(umask)
            os.chmod(tmp, 0o666 & ~umask)
        os.replace(tmp, out)
        try:
            fd = os.open(out.parent, os.O_RDONLY)
            try:
                os.fsync(fd)
            finally:
                os.close(fd)
        except OSError:
            pass
    finally:
        if os.path.exists(tmp):
            os.remove(tmp)
    return out


def save_atomic(prs, out, src=None):
    """Save `prs` to `out` atomically (rpptx's own save() writes in place). Refuses `out == src`."""
    return _publish(prs.save, out, src)


def replace_batch(src, out, edits):
    """Apply [(old, new, expected_count), ...] in order; every count must match or nothing is written."""
    if Path(src).resolve() == Path(out).resolve():
        raise EditError("refusing to write over the input")
    with tempfile.TemporaryDirectory() as tmp:
        cur = Path(src)
        wrong = []
        for k, (old, new, expected) in enumerate(edits):
            nxt = Path(tmp) / f"step{k}.pptx"
            res = cli("replace", cur, "-p", old, "-v", new, "--expect", expected, "-o", nxt, check=False)
            if res.returncode:
                wrong.append(f"{old!r}: {(res.stdout + res.stderr).strip()}")
                continue
            cur = nxt
        if wrong:
            raise EditError("; ".join(wrong))
        return _publish(lambda t: shutil.copyfile(cur, t), out, src)


def fix_template_content_type(path):
    """Rewrite a template main part's content type to the presentation one, in place (atomically, every other
    entry unchanged). Returns True if the file was changed. Kept for the scripts that call it: rpptx's save
    now writes the content type the path extension names."""
    import zipfile
    path = Path(path)
    with zipfile.ZipFile(path) as z:
        items = [(info, z.read(info.filename)) for info in z.infolist()]
    old, new = b"presentationml.template.main+xml", b"presentationml.presentation.main+xml"
    if not any(info.filename == "[Content_Types].xml" and old in data for info, data in items):
        return False

    def write(tmp):
        with zipfile.ZipFile(tmp, "w") as out:
            for info, data in items:
                out.writestr(info, data.replace(old, new) if info.filename == "[Content_Types].xml" else data)
    mode = path.stat().st_mode & 0o7777
    _publish(write, path)
    os.chmod(path, mode)
    return True


def solid_png(width, height, rgb=(128, 128, 128)):
    """A PNG of one colour, made with the standard library (the pinned Python has no Pillow)."""
    import struct
    import zlib
    row = b"\x00" + bytes(rgb) * width

    def chunk(kind, data):
        return struct.pack(">I", len(data)) + kind + data + struct.pack(">I", zlib.crc32(kind + data) & 0xFFFFFFFF)

    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0))
            + chunk(b"IDAT", zlib.compress(row * height)) + chunk(b"IEND", b""))


def image_size(data):
    """(width, height) of PNG, JPEG or GIF bytes, with the standard library; None for other formats."""
    import struct
    if data[:8] == b"\x89PNG\r\n\x1a\n":
        return struct.unpack(">II", data[16:24])
    if data[:6] in (b"GIF87a", b"GIF89a"):
        return struct.unpack("<HH", data[6:10])
    if data[:2] == b"\xff\xd8":
        i = 2
        while i + 9 < len(data):
            if data[i] != 0xFF:
                i += 1
                continue
            marker, length = data[i + 1], struct.unpack(">H", data[i + 2:i + 4])[0]
            if 0xC0 <= marker <= 0xCF and marker not in (0xC4, 0xC8, 0xCC):
                h, w = struct.unpack(">HH", data[i + 5:i + 9])
                return w, h
            i += 2 + length
    return None


def overflowing(path, width_factor=1.0):
    """[(slide_index, shape_id, name)] of text frames whose text does not fit, from rpptx's own line breaker (the
    frames `rpptx fit` reports, which also gives the font scale each one needs)."""
    return [(f.slide_index, f.shape_id, f.name) for f in rpptx.Presentation(path).text_layout(width_factor=width_factor)
            if f.overflow]


def walk(shapes, depth=0):
    for sh in shapes:
        yield depth, sh
        if int(sh.shape_type or 0) == 6:
            yield from walk(sh.shapes, depth + 1)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)
    p = sub.add_parser("replace")
    p.add_argument("src")
    p.add_argument("out")
    p.add_argument("--edit", nargs=3, action="append", metavar=("OLD", "NEW", "COUNT"), required=True)
    p = sub.add_parser("overflow")
    p.add_argument("src")
    p = sub.add_parser("shapes")
    p.add_argument("src")
    p.add_argument("--slide", type=int, help="one-based slide number")
    a = ap.parse_args()
    try:
        if a.cmd == "replace":
            replace_batch(a.src, a.out, [(o, n, int(c)) for o, n, c in a.edit])
            print(f"wrote {a.out}")
        elif a.cmd == "overflow":
            for s, sid, name in overflowing(a.src):
                print(f"slide {s + 1}  shape {sid}  {name}")
        elif a.cmd == "shapes":
            prs = rpptx.Presentation(a.src)
            for k, slide in enumerate(prs.slides):
                if a.slide and k + 1 != a.slide:
                    continue
                print(f"slide {k + 1}")
                for depth, sh in walk(slide.shapes):
                    geo = (sh.left, sh.top, sh.width, sh.height)
                    text = (sh.text[:50] + "...") if sh.has_text_frame and len(sh.text) > 50 else (sh.text if sh.has_text_frame else "")
                    kind = getattr(sh.shape_type, "name", sh.shape_type)
                    print(f"{'  ' * (depth + 1)}{sh.shape_id:>4} {kind!s:<12} {sh.name:<24} {geo} {text!r}")
    except EditError as e:
        print(f"refused: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
