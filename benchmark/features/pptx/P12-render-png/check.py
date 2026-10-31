"""P12 oracle (standard library only for the images): one PNG per slide (8), each 1333x750 (+-2 px) as 13.333 x 7.5
inches at 100 dpi, none blank, not all identical; and the two non-text elements drawn: the chart frame of slide 6
holds bars in the series colour (1E2761), and slide 7 shows the photo (over 50,000 saturated pixels, the
master's small logo being the only other colour)."""
import hashlib
import struct
import sys
import zlib
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "lib"))
import oracle  # noqa: E402


def png_decode(data):
    """(width, height, compressed size, digest, rows of (r, g, b)) of an 8-bit RGB or RGBA PNG."""
    if data[:8] != b"\x89PNG\r\n\x1a\n":
        return None
    w, h, depth, ctype = struct.unpack(">IIBB", data[16:26])
    i, idat = 8, b""
    while i < len(data):
        n, kind = struct.unpack(">I4s", data[i:i + 8])
        if kind == b"IDAT":
            idat += data[i + 8:i + 8 + n]
        i += 12 + n
    raw = zlib.decompress(idat)
    bpp = {2: 3, 6: 4, 0: 1, 4: 2}.get(ctype)
    if depth != 8 or bpp is None:
        return w, h, len(idat), hashlib.sha256(raw).hexdigest(), None
    stride, rows, prev, pos = w * bpp, [], bytearray(w * bpp), 0
    for _ in range(h):
        f, line = raw[pos], bytearray(raw[pos + 1:pos + 1 + stride])
        pos += 1 + stride
        for x in range(stride):
            a = line[x - bpp] if x >= bpp else 0
            b, c = prev[x], (prev[x - bpp] if x >= bpp else 0)
            if f == 1:
                line[x] = (line[x] + a) & 255
            elif f == 2:
                line[x] = (line[x] + b) & 255
            elif f == 3:
                line[x] = (line[x] + (a + b) // 2) & 255
            elif f == 4:
                p = a + b - c
                pa, pb, pc = abs(p - a), abs(p - b), abs(p - c)
                line[x] = (line[x] + (a if pa <= pb and pa <= pc else b if pb <= pc else c)) & 255
        rows.append(bytes(line) if bpp >= 3 else bytes(v for v in line for _ in range(3)))
        prev = line
    pix = [[tuple(r[x * bpp:x * bpp + 3]) if bpp >= 3 else (r[x],) * 3 for x in range(w)] for r in rows]
    return w, h, len(idat), hashlib.sha256(raw).hexdigest(), pix


def near(pix, box, rgb, tol=25):
    x0, y0, x1, y1 = box
    return sum(1 for row in pix[y0:y1] for p in row[x0:x1] if all(abs(p[k] - rgb[k]) <= tol for k in range(3)))


def check(out, inputs):
    pngs = sorted(out.glob("*.png"), key=lambda p: int("".join(c for c in p.stem.rsplit("slide", 1)[-1] if c.isdigit())
                                                         or 0))
    errors = []
    if len(pngs) != 8:
        return [f"{len(pngs)} PNG files, expected 8"]
    infos = []
    for k, p in enumerate(pngs, 1):
        data = p.read_bytes()
        if k in (6, 7):
            info = png_decode(data)
        else:
            w, h = struct.unpack(">II", data[16:24]) if data[:8] == b"\x89PNG\r\n\x1a\n" else (0, 0)
            info = (w, h, len(data), hashlib.sha256(data).hexdigest(), None)
        infos.append(info)
        if info is None or info[0] == 0:
            errors.append(f"{p.name} is not a PNG")
            continue
        w, h, size, _, _ = info
        if abs(w - 1333) > 2 or abs(h - 750) > 2:
            errors.append(f"{p.name} is {w}x{h}, expected 1333x750")
        if size < 3000:
            errors.append(f"{p.name} looks blank ({size} bytes of image data)")
    if len({i[3] for i in infos if i}) < len(infos):
        errors.append("some slides render identically")
    bars = near(infos[5][4], (80, 160, 930, 640), (0x1E, 0x27, 0x61)) if infos[5] and infos[5][4] else 0
    photo = sum(1 for row in infos[6][4] for p in row if max(p) - min(p) > 80) if infos[6] and infos[6][4] else 0
    if bars < 5000:
        errors.append(f"slide 6: the chart's bars are not drawn ({bars} pixels in the series colour)")
    if photo < 50000:
        errors.append(f"slide 7: the picture is not drawn ({photo} saturated pixels)")
    return errors, {"files": [p.name for p in pngs], "bar_pixels": bars, "photo_pixels": photo}


oracle.main(check, pptx_output=False)
