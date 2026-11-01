"""R15 inputs: A and B from two different templates. Same style ids, different definitions: A's body font Calibri
and blue Calibri Light headings, B's body font Georgia (in its docDefaults) and red Georgia headings; B also has a
style of its own (BNote). Both number a list with numId 1 (A decimal "1.", B lower roman "i)") and both hold a
picture named media/image1.png with different pixels."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import review_build as b  # noqa: E402

A_BODY = [
    ("Title", "Annual report"),
    ("Heading1", "Overview"),
    (None, "The year closed with four new service contracts."),
    ("ListParagraph", "Revenue grew in every region."),
    ("ListParagraph", "Two sites were refurbished."),
    (None, "PICTURE"),
    ("Heading1", "Results"),
    (None, "Results are detailed in the appendix."),
]
B_BODY = [
    ("Heading1", "Appendix: site visits"),
    (None, "Each site was visited twice during the year."),
    ("BNote", "Visits in December were postponed because of snow."),
    ("ListParagraph", "North site: handrails replaced."),
    ("ListParagraph", "South site: drainage cleared."),
    (None, "PICTURE"),
    (None, "Figure B1. South site after works."),
]
A_COLOR, B_COLOR, NOTE_COLOR = "2F5496", "C00000", "00B050"
A_PNG = b.png(160, 80, (70, 130, 180))
B_PNG = b.png(160, 80, (200, 60, 40))
NUMBERING = {"A": ("decimal", "%1."), "B": ("lowerRoman", "%1)")}


def write(path, body_spec, png, fmt, styles, seed):
    b.reset_ids()
    b.IDS.n += seed
    pkg = b.Package(title=path.stem)
    pkg.styles = styles
    pkg.numbering([(0, 1, fmt[0], fmt[1], 720, 360)])
    rid = pkg.image("image1.png", png)
    body = []
    for style, text in body_spec:
        if text == "PICTURE":
            body.append(b.para(b.inline_picture(rid, 1828800, 914400, 1, "Picture 1")))
        elif style == "ListParagraph":
            body.append(b.numbered(b.frag(text, cuts=(6,)), 1))
        else:
            body.append(b.para(b.frag(text, cuts=(len(text) // 2,) if len(text) > 20 else ()), style=style))
    return pkg.save(path, "".join(body) + b.sect_pr())


def build(out_dir):
    out_dir = Path(out_dir)
    write(out_dir / "a.docx", A_BODY, A_PNG, NUMBERING["A"], b.styles_xml(), 0)
    note = ('<w:style w:type="paragraph" w:customStyle="1" w:styleId="BNote"><w:name w:val="B Note"/>'
            '<w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:ind w:left="360"/></w:pPr><w:rPr><w:i/><w:iCs/>'
            f'<w:color w:val="{NOTE_COLOR}"/></w:rPr></w:style>')
    styles_b = b.styles_xml(heading_font="Georgia", heading_color=B_COLOR, body_font="Georgia", extra_styles=note)
    return write(out_dir / "b.docx", B_BODY, B_PNG, NUMBERING["B"], styles_b, 0x10000)


if __name__ == "__main__":
    print(build(sys.argv[1]))
