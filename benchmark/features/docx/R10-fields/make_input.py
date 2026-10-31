"""R10 input: three pages (page breaks) whose body holds a DATE field, a NUMPAGES field on page 1 and a PAGE field
on page 3, all with stale cached results as Word leaves them; a footer with PAGE of NUMPAGES."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import review_build as b  # noqa: E402

FIELDS = [  # (instruction, stale result, expected result: None for today's date)
    ('DATE \\@ "yyyy-MM-dd"', "2020-01-01", None),
    ("NUMPAGES", "1", "3"),
    ("PAGE", "1", "3"),
]


def build(out_dir):
    b.reset_ids()
    pkg = b.Package()
    hdr, ftr = pkg.header_footer(
        footer_content=b.para(b.run("Page ") + b.field("PAGE", "1") + b.run(" of ") + b.field("NUMPAGES", "1"),
                              style="Footer"))
    body = [
        b.para(b.frag("Inspection report"), style="Title"),
        b.para(b.frag("Report date: ") + b.field(FIELDS[0][0], FIELDS[0][1])),
        b.para(b.frag("This report has ") + b.field(FIELDS[1][0], FIELDS[1][1]) + b.frag(" pages.")),
        b.para(b.frag(b.lorem(1), cuts=(30,))),
        b.para(b.frag(b.lorem(2)) + b.page_break()),
        b.para(b.frag("Findings"), style="Heading1"),
        b.para(b.frag(b.lorem(3), cuts=(50,)) + b.page_break()),
        b.para(b.frag("Sign-off"), style="Heading1"),
        b.para(b.frag("This paragraph is on page ") + b.field(FIELDS[2][0], FIELDS[2][1]) + b.frag(".")),
    ]
    return pkg.save(Path(out_dir) / "input.docx", "".join(body) + b.sect_pr(hdr, ftr))


if __name__ == "__main__":
    print(build(sys.argv[1]))
