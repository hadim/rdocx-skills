"""R09 input: a Word table of contents (TOC field in a docPartObj content control, hyperlinked entries with
PAGEREF fields, _Toc bookmarks on the headings) gone stale: two headings renamed, four added since. Pages are
set by page breaks before each Heading 1, and one Heading 2 falls mid-page after a page and a half of text, so
every renderer puts each heading on the same page."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import review_build as b  # noqa: E402

HEADINGS = [  # (level, text, page, filler paragraphs after it, page break before)
    (1, "Introduction", 2, 2, True),
    (1, "Scope of work", 3, 1, True),
    (2, "Deliverables", 3, 2, False),
    (1, "Payment terms", 4, 14, True),
    (2, "Late payment", 5, 2, False),
    (3, "Interest rate", 5, 1, False),
    (1, "Termination", 6, 2, True),
]
STALE = [(1, "Introduction", "1", "_Toc1001"), (1, "Scope", "2", "_Toc1002"), (1, "Payment", "3", "_Toc1004")]
EXPECTED = [(lvl, text, page) for lvl, text, page, _, _ in HEADINGS]
TAB = '<w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9350"/></w:tabs>'


def toc_entry(level, text, page, bm, first=False, last=False):
    r = b.IDS.rsid
    head = ""
    if first:
        head = (f'<w:r w:rsidR="{r()}"><w:fldChar w:fldCharType="begin"/></w:r><w:r w:rsidR="{r()}"><w:instrText '
                'xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u </w:instrText></w:r><w:r w:rsidR="{0}"><w:fldChar '
                'w:fldCharType="separate"/></w:r>'.format(r()))
    link = (f'<w:hyperlink w:anchor="{bm}" w:history="1"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/><w:noProof/>'
            f'</w:rPr><w:t>{text}</w:t></w:r><w:r><w:rPr><w:noProof/><w:webHidden/></w:rPr><w:tab/></w:r>'
            + b.field(f"PAGEREF {bm} \\h", page, "<w:rPr><w:noProof/><w:webHidden/></w:rPr>") + "</w:hyperlink>")
    return b.para(head + link, style=f"TOC{level}", ppr=TAB)


def build(out_dir):
    b.reset_ids()
    pkg = b.Package()
    entries = [toc_entry(lvl, t, pg, bm, first=(k == 0)) for k, (lvl, t, pg, bm) in enumerate(STALE)]
    end = b.para(f'<w:r w:rsidR="{b.IDS.rsid()}"><w:rPr><w:b/><w:bCs/><w:noProof/></w:rPr>'
                 '<w:fldChar w:fldCharType="end"/></w:r>')
    toc = ('<w:sdt><w:sdtPr><w:id w:val="-1427102539"/><w:docPartObj><w:docPartGallery w:val="Table of Contents"/>'
           '<w:docPartUnique/></w:docPartObj></w:sdtPr><w:sdtEndPr><w:rPr><w:b/><w:bCs/><w:noProof/></w:rPr>'
           '</w:sdtEndPr><w:sdtContent>' + b.para(b.run("Contents"), style="TOCHeading")
           + "".join(entries) + end + "</w:sdtContent></w:sdt>")
    body = [b.para(b.frag("Service agreement"), style="Title"), toc]
    marks = {"Introduction": ("1001", "_Toc1001"), "Scope of work": ("1002", "_Toc1002"),
             "Payment terms": ("1004", "_Toc1004")}
    seed = 0
    for lvl, text, _page, n, brk in HEADINGS:
        content = b.frag(text)
        if text in marks:
            content = b.bookmark(*marks[text], content)
        body.append(b.para(content, style=f"Heading{lvl}", page_break_before=brk))
        for _ in range(n):
            seed += 1
            body.append(b.para(b.frag(b.lorem(seed), cuts=(40, 160))))
    return pkg.save(Path(out_dir) / "input.docx", "".join(body) + b.sect_pr())


if __name__ == "__main__":
    print(build(sys.argv[1]))
