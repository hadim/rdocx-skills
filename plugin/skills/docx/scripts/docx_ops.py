"""Helpers for editing .docx files safely with rdocx. Import from a script run with the pinned Python
(`~/.local/share/rdocx-skills/current/bin/python`, with this folder on PYTHONPATH), or run as a command:

  python docx_ops.py text IN.docx                                                every paragraph of every story, by part
  python docx_ops.py count IN.docx TEXT                                          occurrences of TEXT by story kind
  python docx_ops.py replace IN.docx OUT.docx --edit OLD NEW COUNT [--edit ...] [--allow-unreached]
  python docx_ops.py comment IN.docx OUT.docx --anchor TEXT --text COMMENT --author NAME [--occurrence N] [--in-tables] [--date ISO]
  python docx_ops.py toc IN.docx OUT.docx                                        rebuild the table of contents
  python docx_ops.py pages IN.docx                                               page count from rdocx's layout

Every function that writes takes an output path, writes a temporary file next to it, flushes it to disk and
renames it into place, keeps the input's file mode, and refuses to write over its input. Exit codes: 0 done, 1 refused (nothing written), 2 usage.
"""
import argparse
import collections
import io
import json
import os
import re
import subprocess
import sys
import tempfile
import xml.etree.ElementTree as ET
import zipfile
from datetime import datetime, timezone
from pathlib import Path

import rdocx

BIN = Path(os.environ.get("RDOCX_BIN_DIR", Path(os.environ.get("RDOCX_HOME", Path.home() / ".local/share/rdocx-skills")) / "current" / "bin"))


class EditError(RuntimeError):
    """An edit could not be done exactly as asked: nothing was written, and `doc` was not changed."""


def cli(*args, check=True):
    """Run the pinned `rdocx` CLI (never another one found on PATH)."""
    exe = BIN / ("rdocx.exe" if os.name == "nt" else "rdocx")
    if not exe.exists():
        raise FileNotFoundError(f"{exe} not found: install the pinned build (rdocx_env.py install) or set RDOCX_BIN_DIR")
    res = subprocess.run([str(exe), *map(str, args)], capture_output=True, encoding="utf-8")  # the CLI writes UTF-8
    if check and res.returncode:
        raise RuntimeError((res.stdout + res.stderr).strip())
    return res


W_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
MC_FALLBACK = "{http://schemas.openxmlformats.org/markup-compatibility/2006}Fallback"
STORY_PART = re.compile(r"word/(document|header|footer|footnotes|endnotes)(\d*)\.xml$")
STORY_KIND = {"document": "body", "header": "header", "footer": "footer", "footnotes": "footnote", "endnotes": "endnote"}


def now():
    """The current UTC time as an RFC 3339 string, the form rdocx takes for comment and revision dates."""
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def _fsync(path):
    """Flush a file to disk, through a handle open for writing: Windows refuses to flush one open for reading
    (EBADF). Call it before setting the file's mode, which may make it read-only."""
    with open(path, "rb+") as f:
        os.fsync(f.fileno())


def _fsync_dir(folder):
    """Flush a rename to disk (best effort: some file systems refuse to open a folder)."""
    try:
        fd = os.open(folder, os.O_RDONLY)
        try:
            os.fsync(fd)
        finally:
            os.close(fd)
    except OSError:
        pass


def save_atomic(doc, out, src=None):
    """Save `doc` to `out` through a temporary file in the same folder, flushed to disk, then renamed. The
    file gets the mode of `src` when given (else the default mode under the umask). Refuses `out == src`."""
    out = Path(out)
    if src is not None and Path(src).resolve() == out.resolve():
        raise EditError("refusing to write over the input; write to a new path, check it, then replace")
    fd, tmp = tempfile.mkstemp(prefix=f".{out.stem}.", suffix=out.suffix, dir=out.parent)
    os.close(fd)
    try:
        doc.save(tmp)
        _fsync(tmp)
        if src is not None and Path(src).exists():
            os.chmod(tmp, Path(src).stat().st_mode & 0o7777)
        else:
            umask = os.umask(0)
            os.umask(umask)
            os.chmod(tmp, 0o666 & ~umask)
        os.replace(tmp, out)
        _fsync_dir(out.parent)
    finally:
        if os.path.exists(tmp):
            os.remove(tmp)
    return out


def textmap(path):
    """`rdocx text --json`: body paragraphs and table cells with body_index, path, style, runs (accepted view:
    tracked insertions in, deletions out). Headers, footers, notes and text boxes are in the record's
    `stories`, not here: see all_text."""
    return json.loads(cli("text", "--json", path).stdout)["paragraphs"]


def _package(source):
    """The .docx bytes of a Document, bytes, or a path."""
    if hasattr(source, "to_bytes"):
        return source.to_bytes()
    if isinstance(source, (bytes, bytearray)):
        return bytes(source)
    return Path(source).read_bytes()


def _part_paragraphs(xml):
    """[(style id, text)] for every w:p of one XML part, in document order. A paragraph's text is its own
    w:t and w:tab ("\\t") (a text box's paragraphs are listed on their own, not inside the paragraph that anchors
    them), content controls, tracked insertions, hyperlinks and fields included, deletions (w:delText) out, and
    mc:Fallback copies skipped, so a Word text box counts once."""
    out = []

    def walk(el, acc):
        if el.tag == MC_FALLBACK:
            return
        if el.tag == W_NS + "p":
            slot = len(out)
            out.append(None)
            mine = []
            style = el.find(f"{W_NS}pPr/{W_NS}pStyle")
            for child in el:
                walk(child, mine)
            out[slot] = (style.get(W_NS + "val") if style is not None else None, "".join(mine))
            return
        if el.tag == W_NS + "t":
            acc.append(el.text or "")
        for child in el:
            if child.tag == W_NS + "tab" and el.tag == W_NS + "r":  # a tab character, not a tab stop of w:pPr
                acc.append("\t")
                continue
            walk(child, acc)

    walk(ET.fromstring(xml), [])
    return out


def story_paragraphs(source):
    """[(part name, story kind, style id, text)] for every paragraph of the body (table cells, text boxes,
    content controls included), the headers and footers (each variant part), the footnotes and endnotes, read
    from the package XML: what a reader of the document sees, in the accepted view of tracked changes (a
    paragraph that a tracked deletion removes is listed with empty text)."""
    out = []
    with zipfile.ZipFile(io.BytesIO(_package(source))) as z:
        names = sorted((n for n in z.namelist() if STORY_PART.match(n)),
                       key=lambda n: (n != "word/document.xml", n))
        for name in names:
            kind = STORY_KIND[STORY_PART.match(name).group(1)]
            out += [(name, kind, style, text) for style, text in _part_paragraphs(z.read(name))]
    return out


def all_text(source):
    """[(story kind, paragraph text)] for every paragraph of every story (see story_paragraphs)."""
    return [(kind, text) for _, kind, _, text in story_paragraphs(source)]


def count(source, text):
    """Occurrences of `text` by story kind, as a reader sees them: {"body": n, "footer": n, ...}. A header or
    footer text counts once per variant part (default, first page, even pages), like replacement does."""
    got = collections.Counter()
    for kind, t in all_text(source):
        if text in t:
            got[kind] += t.count(text)
    return dict(got)


def replace_batch(src, out, edits, allow_unreached=False):
    """Apply [(old, new, expected_count), ...] in order, all or nothing. A dry run on an in-memory copy comes
    first; nothing is written, and EditError names every problem, when a replacement count differs from its
    expected count, or when `old` is still in the document after its replacement (a match across the edge of a
    content control, a tracked insertion or a simple field), unless allow_unreached=True. Counts are rdocx's:
    headers and footers once per variant part, a Word text box once."""
    doc = rdocx.Document.open(src)
    probe = rdocx.Document.from_bytes(doc.to_bytes())
    wrong = []
    for old, new, expected in edits:
        seen = sum(count(probe, old).values())
        got = probe.try_replace_text(old, new)
        if got != expected:
            wrong.append(f"{old!r}: expected {expected}, found {got}")
        if allow_unreached:
            continue
        if old not in new:
            left = [(name, style, t) for name, _, style, t in story_paragraphs(probe) if old in t]
            if left:
                where = "; ".join(f"{name}{f' [{style}]' if style else ''}: {t[:60]!r}" for name, style, t in left[:5])
                toc = any(style and re.match(r"(?i)(toc|tm|inhalt|verzeichnis)\s*\d", style) for _, style, _ in left)
                wrong.append(f"{old!r}: {sum(t.count(old) for _, _, t in left)} occurrence(s) left after the replacement, "
                             f"out of its reach ({where}{'; ...' if len(left) > 5 else ''})"
                             + ("; table of contents entries come back with rebuild_toc" if toc else "")
                             + "; pass allow_unreached=True once handled")
        elif seen > got:
            wrong.append(f"{old!r}: {seen} occurrence(s) seen, {got} replaced; pass allow_unreached=True once handled")
    if wrong:
        raise EditError("; ".join(wrong))
    for old, new, _ in edits:
        doc.try_replace_text(old, new)
    return save_atomic(doc, out, src)


def _starts(text, anchor):
    """Every start offset of `anchor` in `text`, overlapping ones included."""
    out, k = [], text.find(anchor)
    while k != -1:
        out.append(k)
        k = text.find(anchor, k + 1)
    return out


def locate(doc, anchor, occurrence=1):
    """(flow index, run offset, body index) of the n-th occurrence of `anchor` within the text of one paragraph
    that is a direct child of the body. Paragraphs in table cells and in block content controls (a table of
    contents, Google Docs blocks) are skipped: comments cannot be anchored there by body index. The offset
    counts the characters of the paragraph's runs (`"".join(r.text for r in p.runs)`), which leave out the
    text inside a simple field, a smart tag or a custom XML element that `Paragraph.text` shows: an anchor
    that sits in such an element, or crosses its edge, is refused."""
    seen = 0
    for i, p in enumerate(doc.paragraphs):
        shown = _starts(p.text, anchor)
        if not shown:
            continue
        try:
            bi = doc.find_content_index(doc.paragraphs[i])
        except ValueError:  # not a direct child of the body
            continue
        if seen + len(shown) < occurrence:
            seen += len(shown)
            continue
        offsets = _starts("".join(r.text for r in doc.paragraphs[i].runs), anchor)
        if len(offsets) != len(shown):
            raise EditError(f"anchor {anchor!r} occurrence {occurrence}: in its paragraph the text sits inside a simple "
                            "field, a smart tag or a custom XML element (or crosses its edge), whose text is not in "
                            "Paragraph.runs: anchor on text outside that element")
        return i, offsets[occurrence - seen - 1], bi
    raise EditError(f"anchor {anchor!r} occurrence {occurrence} not found in the body's own paragraphs "
                    "(table cells and content-control blocks are not searched)")


def isolate(doc, flow_index, start, end):
    """Split runs so that characters [start, end) of paragraph `flow_index` form whole runs. Offsets count the
    characters of the paragraph's runs joined, as `locate` returns them (not `Paragraph.text`, which also holds
    the text of simple fields, smart tags and custom XML). Returns the (first, last_exclusive) run indices
    covering them. The paragraph must be a direct child of the body: split_run takes its body index, as
    RunPosition and find_content_index do."""
    before = doc.paragraphs[flow_index].text
    body_index = doc.find_content_index(doc.paragraphs[flow_index])

    def split_at(offset):
        pos = 0
        for k, r in enumerate(doc.paragraphs[flow_index].runs):
            n = len(r.text)
            if offset == pos:
                return k
            if pos < offset < pos + n:
                doc.split_run(body_index, k, offset - pos)
                return k + 1
            pos += n
        return len(doc.paragraphs[flow_index].runs)

    split_at(end)
    first = split_at(start)
    last = split_at(end)
    if doc.paragraphs[flow_index].text != before:
        raise EditError("splitting the runs changed the paragraph text")
    return first, last


def anchored_text(data, comment_id):
    """The text a comment is anchored on in a saved document (bytes or path), or None if it has no range.

    Read through rdocx (`Comment.anchor_text`): the accepted view of tracked changes, through content controls
    (Google Docs' goog_rdk wrappers included) and in table cells, the paragraphs of a range over several joined
    with "\n". rdocx gives "" for a comment that has a reference mark but no range: that maps to None, like a
    reply, an unknown id or a comment with no marks at all."""
    doc = rdocx.Document.from_bytes(data) if isinstance(data, bytes) else rdocx.Document.open(data)
    comment = next((c for c in doc.comments if c.id == comment_id), None)
    return (comment.anchor_text or None) if comment is not None else None


def _comment(doc, anchor, text, author, initials, occurrence, date):
    flow, start, bi = locate(doc, anchor, occurrence)
    first, last = isolate(doc, flow, start, start + len(anchor))
    rng = rdocx.RunRange(start=rdocx.RunPosition(body_index=bi, run_index=first),
                         end=rdocx.RunPosition(body_index=bi, run_index=last))
    return doc.add_comment(rng, author=author, text=text, initials=initials, date=date)


def comment_landings(data):
    """{comment id: True if its range starts inside a table cell} for every comment range of document.xml."""
    out = {}

    def walk(el, in_cell):
        if el.tag == W_NS + "commentRangeStart":
            out[int(el.get(W_NS + "id"))] = in_cell
        for child in el:
            walk(child, in_cell or el.tag == W_NS + "tc")

    with zipfile.ZipFile(io.BytesIO(data) if isinstance(data, bytes) else data) as z:
        walk(ET.fromstring(z.read("word/document.xml")), False)
    return out


def _cell_occurrence(doc, anchor, occurrence):
    """The zero-based occurrence number that `add_comment_on_text` gives the n-th occurrence of `anchor`
    inside a table cell. rdocx numbers the anchor's occurrences over the whole main story in document order;
    anchoring every one of them on a copy (the text does not change) shows which ones sit in a cell."""
    probe = rdocx.Document.from_bytes(doc.to_bytes())
    ids = []
    while True:
        try:
            ids.append(probe.add_comment_on_text(anchor, author="probe", text="probe", occurrence=len(ids)))
        except rdocx.RdocxError:
            break
    landing = comment_landings(probe.to_bytes())
    in_cells = [k for k, cid in enumerate(ids) if landing.get(cid)]
    if len(in_cells) < occurrence:
        raise EditError(f"anchor {anchor!r} occurrence {occurrence} not found in table cells "
                        f"({len(in_cells)} there, {len(ids)} in the main story)")
    return in_cells[occurrence - 1]


def comment_on_text(doc, anchor, text, author, initials=None, occurrence=1, date=None, in_tables=False):
    """Anchor a comment on exactly `anchor`, dated `date` (RFC 3339; default: now, UTC). `occurrence` counts
    from 1 in the body's own paragraphs, or, with `in_tables=True`, in the paragraphs of table cells (nested
    tables included), in document order. The whole operation is first run on a copy and the anchored text
    read back: if rdocx refuses the range or anchors it elsewhere, EditError is raised and `doc` is left
    unchanged. Returns the comment id."""
    date = date or now()
    if in_tables:
        k = _cell_occurrence(doc, anchor, occurrence)

        def act(d):
            return d.add_comment_on_text(anchor, author=author, text=text, occurrence=k, initials=initials, date=date)
    else:
        def act(d):
            return _comment(d, anchor, text, author, initials, occurrence, date)
    probe = rdocx.Document.from_bytes(doc.to_bytes())
    why = "nothing was written; report it as a new gap"
    try:
        cid = act(probe)
    except rdocx.RdocxError as e:
        raise EditError(f"rdocx refused the range ({e}): {why}") from e
    data = probe.to_bytes()
    got = anchored_text(data, cid)
    if got != anchor:
        raise EditError(f"the comment would be anchored on {got!r} instead of {anchor!r}: {why}")
    if in_tables and not comment_landings(data).get(cid):
        raise EditError(f"the comment would be anchored outside the table cells: {why}")
    return act(doc)


def rebuild_toc(doc):
    """doc.rebuild_toc(). Kept for the scripts that call it: a fresh open of a Word file now rebuilds
    directly."""
    return doc.rebuild_toc()


def fix_template_content_type(path):
    """Rewrite a template main part's content type to the document one, in place (atomically, every other
    entry unchanged). Returns True if the file was changed. Kept for the scripts that call it: rdocx's save
    now writes the content type the path extension names."""
    path = Path(path)
    with zipfile.ZipFile(path) as z:
        items = [(info, z.read(info.filename)) for info in z.infolist()]
    old, new = b"wordprocessingml.template.main+xml", b"wordprocessingml.document.main+xml"
    if not any(info.filename == "[Content_Types].xml" and old in data for info, data in items):
        return False
    fd, tmp = tempfile.mkstemp(prefix=f".{path.stem}.", suffix=path.suffix, dir=path.parent)
    os.close(fd)
    try:
        with zipfile.ZipFile(tmp, "w") as out:
            for info, data in items:
                out.writestr(info, data.replace(old, new) if info.filename == "[Content_Types].xml" else data)
        _fsync(tmp)
        os.chmod(tmp, path.stat().st_mode & 0o7777)
        os.replace(tmp, path)
        _fsync_dir(path.parent)
    finally:
        if os.path.exists(tmp):
            os.remove(tmp)
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


def pages(path):
    """Page count from rdocx's own layout (on Word's line heights; confirm in Word when it matters)."""
    return max(f.physical_page for f in rdocx.Document.open(path).layout())


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)
    p = sub.add_parser("text")
    p.add_argument("src")
    p = sub.add_parser("count")
    p.add_argument("src")
    p.add_argument("text")
    p = sub.add_parser("replace")
    p.add_argument("src")
    p.add_argument("out")
    p.add_argument("--edit", nargs=3, action="append", metavar=("OLD", "NEW", "COUNT"), required=True)
    p.add_argument("--allow-unreached", action="store_true")
    p = sub.add_parser("comment")
    p.add_argument("src")
    p.add_argument("out")
    p.add_argument("--anchor", required=True)
    p.add_argument("--text", required=True)
    p.add_argument("--author", required=True)
    p.add_argument("--initials")
    p.add_argument("--occurrence", type=int, default=1)
    p.add_argument("--in-tables", action="store_true", help="count and anchor in table cells")
    p.add_argument("--date", help="RFC 3339, default now (UTC)")
    p = sub.add_parser("toc")
    p.add_argument("src")
    p.add_argument("out")
    p = sub.add_parser("pages")
    p.add_argument("src")
    a = ap.parse_args()
    try:
        if a.cmd == "text":
            for name, kind, style, t in story_paragraphs(a.src):
                print(f"{kind}\t{name}\t{style or ''}\t{t}")
        elif a.cmd == "count":
            print(json.dumps(count(rdocx.Document.open(a.src), a.text)))
        elif a.cmd == "replace":
            replace_batch(a.src, a.out, [(o, n, int(c)) for o, n, c in a.edit], a.allow_unreached)
            print(f"wrote {a.out}")
        elif a.cmd == "comment":
            doc = rdocx.Document.open(a.src)
            cid = comment_on_text(doc, a.anchor, a.text, a.author, a.initials, a.occurrence, a.date, a.in_tables)
            save_atomic(doc, a.out, a.src)
            print(f"comment {cid} written to {a.out}")
        elif a.cmd == "toc":
            doc = rdocx.Document.open(a.src)
            rep = rebuild_toc(doc)
            save_atomic(doc, a.out, a.src)
            print(f"{rep.entry_count} entries, {len(rep.diagnostics)} diagnostics, written to {a.out}")
        elif a.cmd == "pages":
            print(pages(a.src))
    except EditError as e:
        print(f"refused: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
