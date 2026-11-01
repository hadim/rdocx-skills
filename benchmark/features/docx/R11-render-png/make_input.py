"""R11 input: a four-page report (page breaks) with a header and a footer page number, tracked insertions and
deletions on pages 1 and 2, a picture, a table and a footnote."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import review_build as b  # noqa: E402

PAGES = 4
A, D = "Alex Author", "Dana Reviewer"


def build(out_dir, name="input.docx"):
    b.reset_ids()
    pkg = b.Package(title="Inspection report")
    hdr, ftr = pkg.header_footer(
        header_content=b.para(b.run("Inspection report - draft for review"), style="Header"),
        footer_content=b.para(b.run("Page ") + b.field("PAGE", "1"), style="Footer"))
    rid = pkg.image("image1.png", b.png(240, 120, (70, 130, 180)))
    body = [
        b.para(b.frag("Inspection report"), style="Title"),
        b.para(b.frag("Summary"), style="Heading1"),
        b.para(b.frag("The inspection took place on 12 October 2026 and covered ", cuts=(20,))
               + b.dele("three", A, "2026-10-14T09:00:00Z") + b.ins(b.run("four"), D, "2026-10-15T10:00:00Z")
               + b.frag(" sites. No defect needs urgent action.") + b.footnote_ref(1)),
        b.para(b.frag(b.lorem(11), cuts=(25,))),
        b.para(b.ins(b.run("A follow-up visit is planned for the spring."), D, "2026-10-15T10:00:00Z")
               + b.frag(" " + b.lorem(12, 40))),
        b.para(b.inline_picture(rid, 3048000, 1524000, 1, "Picture 1")),
        b.para(b.frag("Figure 1. Site plan.") + b.page_break()),
        b.para(b.frag("Findings"), style="Heading1"),
        b.table([["Site", "Finding", "Action"], ["North", "Worn handrail", "Repair"],
                 ["South", "Loose cover", "Monitor"]], widths=[2000, 4360, 3000]),
        b.para(b.frag("The south site ") + b.dele("needs a new survey and ", A, "2026-10-14T09:00:00Z")
               + b.frag("is otherwise in good condition. " + b.lorem(13, 50)) + b.page_break()),
        b.para(b.frag("Method"), style="Heading1"),
        b.para(b.frag(b.lorem(14, 120), cuts=(60,))),
        b.para(b.frag(b.lorem(15, 90)) + b.page_break()),
        b.para(b.frag("Appendix"), style="Heading1"),
        b.para(b.frag(b.lorem(16, 80))),
    ]
    pkg.notes({1: b.frag("Weather was dry throughout the visit.")})
    return pkg.save(Path(out_dir) / name, "".join(body) + b.sect_pr(hdr, ftr))


if __name__ == "__main__":
    print(build(sys.argv[1]))
