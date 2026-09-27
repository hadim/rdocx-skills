"""The two acceptance matrices, as tests: every operation on a document carrying one identity attribute
(Word and Google Docs write them on every save) or one producer serialisation trait must behave exactly as
on the control. A new trait met in a real file becomes a row here."""
import os
import re
import zipfile

import pytest
import rdocx

from conftest import STAMP, run
from matrix_fixtures import TRAITS, build, rewrite

W14 = "{http://schemas.microsoft.com/office/word/2010/wordml}"

IDENTITY_ROWS = [("paragraphs", a) for a in ("w:rsidR", "w:rsidRDefault", "w:rsidP", "w14:paraId", "w14:textId")] + \
                [("runs", a) for a in ("w:rsidR", "w:rsidRPr", "w:rsidDel")] + \
                [("field runs", a) for a in ("w:rsidR", "w:rsidRPr")] + \
                [("footer field runs", a) for a in ("w:rsidR", "w:rsidRPr")] + \
                [("content control", a) for a in ("w:id", "w:tag")] + \
                [("table rows", a) for a in ("w:rsidR", "w:rsidTr", "w14:paraId")]

# Known failing cells: (row, column) -> gap key
IDENTITY_GAPS = {**{(("runs", a), "toc"): "toc-rsid-field-runs" for a in ("w:rsidR", "w:rsidRPr", "w:rsidDel")},
                 **{(("field runs", a), "toc"): "toc-rsid-field-runs" for a in ("w:rsidR", "w:rsidRPr")},
                 **{(("content control", a), "cmp0"): "compare-sdt-id" for a in ("w:id", "w:tag")},
                 **{(("table rows", a), "save"): "tr-identity-lost" for a in ("w:rsidR", "w:rsidTr", "w14:paraId")}}
TRAIT_GAPS = {("inline content control on first runs", "replace"): "sdt-replace",
              ("inline content control on first runs", "self"): "sdt-replace",
              ("w:orient=\"portrait\" on pgSz", "self"): "compare-own-save-noise",
              ("packed footer fields, no cached result", "cmpfld"): "compare-packed-fields",
              ("empty comments part", "cmpfld"): "empty-comments-reserialised",
              ("empty comments part", "self"): "empty-comments-reserialised"}


def count_attr(path, attr):
    with zipfile.ZipFile(path) as z:
        x = z.read("word/document.xml") + z.read("word/footer1.xml")
    if attr == "w:tag":
        return x.count(b"<w:tag ")
    return x.count(b"<w:id ") if attr == "w:id" else x.count(b" " + attr.encode() + b"=")


def revisions(cli, a, b, out):
    if os.path.exists(out):
        os.remove(out)
    res = run([cli, "compare", a, b, "--author", "R", "--timestamp", STAMP, "-o", out])
    assert res.returncode == 0, (res.stdout + res.stderr).strip()
    return len(rdocx.Document.open(out).revisions)


def params(rows, columns, gaps):
    out = []
    for r in rows:
        for c in columns:
            key = gaps.get((r, c))
            marks = [pytest.mark.gap(key)] if key else []
            out.append(pytest.param(r, c, marks=marks, id=f"{r if isinstance(r, str) else ' '.join(r)} | {c}"))
    return out


@pytest.mark.parametrize("row,column", params(IDENTITY_ROWS, ("save", "replace", "toc", "fields", "render", "cmp0", "cmp1"),
                                              IDENTITY_GAPS))
def test_identity_matrix(row, column, rdocx_cli, tmp_path):
    where, attr = row
    f = build(tmp_path / "case.docx", where, attr)
    if column == "save":
        d = rdocx.Document.open(f)
        assert d.try_replace_text("Body text of section 3.1.", "Body text of section three.") == 1
        d.save(tmp_path / "saved.docx")
        assert count_attr(tmp_path / "saved.docx", attr) == count_attr(f, attr)
    elif column == "replace":
        d = rdocx.Document.open(f)
        assert d.try_replace_text("lorem", "LOREM") == 3
        d.save(tmp_path / "rep.docx")
    elif column == "toc":
        res = run([rdocx_cli, "toc", "rebuild", f, "-o", tmp_path / "toc.docx"])
        assert res.returncode == 0 and "Entries: 6" in res.stdout, res.stdout + res.stderr
    elif column == "fields":
        rdocx.Document.open(f).update_page_fields()
    elif column == "render":
        assert rdocx.Document.open(f).to_pdf()[:5] == b"%PDF-"
    elif column == "cmp0":
        assert revisions(rdocx_cli, build(tmp_path / "plain.docx"), f, tmp_path / "out.docx") == 0
    elif column == "cmp1":
        g = build(tmp_path / "edit.docx", where, attr, word="ALPHA")
        assert revisions(rdocx_cli, f, g, tmp_path / "out.docx") == 2


@pytest.mark.parametrize("row,column", params(list(TRAITS), ("noop", "replace", "toc", "fields", "render", "cmpfld", "cmp1", "self"),
                                              TRAIT_GAPS))
def test_producer_traits_matrix(row, column, rdocx_cli, tmp_path):
    trait = TRAITS[row]
    f = rewrite(build(tmp_path / "case.docx"), *trait)
    if column == "noop":
        rdocx.Document.open(f).save(tmp_path / "noop.docx")
        with zipfile.ZipFile(f) as a, zipfile.ZipFile(tmp_path / "noop.docx") as b:
            assert a.read("word/document.xml") == b.read("word/document.xml")
    elif column in ("replace", "self"):
        d = rdocx.Document.open(f)
        assert d.try_replace_text("alpha", "ALPHA") == 1
        d.save(tmp_path / "rep.docx")
        if column == "self":
            assert revisions(rdocx_cli, f, tmp_path / "rep.docx", tmp_path / "out.docx") == 2
    elif column == "toc":
        res = run([rdocx_cli, "toc", "rebuild", f, "-o", tmp_path / "toc.docx"])
        assert res.returncode == 0 and "Entries: 6" in res.stdout, res.stdout + res.stderr
    elif column in ("fields", "cmpfld"):
        d = rdocx.Document.open(f)
        d.update_page_fields()
        d.save(tmp_path / "fld.docx")
        if column == "fields":
            footer = zipfile.ZipFile(tmp_path / "fld.docx").read("word/footer1.xml").decode()
            assert len(re.findall(r'fldCharType="separate"/>(?:</w:r><w:r>)?<w:t>([^<]*)</w:t>', footer)) == 2
        else:
            revisions(rdocx_cli, f, tmp_path / "fld.docx", tmp_path / "out.docx")
    elif column == "render":
        assert rdocx.Document.open(f).to_pdf()[:5] == b"%PDF-"
    elif column == "cmp1":
        g = rewrite(build(tmp_path / "edit.docx", word="ALPHA"), *trait)
        assert revisions(rdocx_cli, f, g, tmp_path / "out.docx") == 2
