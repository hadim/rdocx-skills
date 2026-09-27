`rdocx convert --to pdf|md|html`, `rpptx convert --to pdf` and `rpptx thumbnail` overwrite any existing output, the input included

The editing commands (`replace`, `comment`, `revision`, `compare`, `toc rebuild`) refuse an existing output
through `ensure_output_paths_available` (`crates/oxml-cli-support/src/lib.rs:111`): "output already exists". So do
`rdocx render`, `rpptx render` and the image outputs of `convert` (`rdocx --to png|jpeg|tiff`, `rpptx --to png`). The other outputs do
not: `rdocx convert --to pdf|md|html` (`crates/rdocx-cli/src/commands.rs:371`), `rpptx convert --to pdf`
(`crates/rpptx-cli/src/commands.rs:311`) and `rpptx thumbnail` write over whatever is at `-o`, the input file
included, and exit 0. One mistyped extension (`-o report.docx` instead of `-o report.pdf`) replaces the
document with its PDF, with no way back. Without `-o`, the default output (the input name with the new
extension) is overwritten the same way.

```bash
python3 - <<'PY'
from docx import Document
from pptx import Presentation
d = Document(); d.add_paragraph("Lorem ipsum."); d.save("in.docx")
p = Presentation(); p.slides.add_slide(p.slide_layouts[6]); p.save("in.pptx")
PY
for t in pdf md html png; do cp in.docx a.docx; rdocx convert a.docx --to $t -o a.docx > /dev/null 2> err; echo "rdocx convert --to $t -o <input>: exit $?, starts with $(head -c 5 a.docx | tr -dc '[:print:]') $(cat err)"; done
for t in pdf png; do cp in.pptx a.pptx; rpptx convert a.pptx --to $t -o a.pptx > /dev/null 2> err; echo "rpptx convert --to $t -o <input>: exit $?, starts with $(head -c 5 a.pptx | tr -dc '[:print:]') $(cat err)"; done
cp in.pptx a.pptx; rpptx thumbnail a.pptx -o a.pptx > /dev/null 2> err; echo "rpptx thumbnail -o <input>: exit $?, starts with $(head -c 5 a.pptx | tr -dc '[:print:]') $(cat err)"
echo keep > in.pdf; rdocx convert in.docx --to pdf > /dev/null; echo "rdocx convert --to pdf, no -o, in.pdf existed: exit $?, in.pdf starts with $(head -c 5 in.pdf)"
```

```
rdocx convert --to pdf -o <input>: exit 0, starts with %PDF- 
rdocx convert --to md -o <input>: exit 0, starts with Lorem 
rdocx convert --to html -o <input>: exit 0, starts with <!DOC 
rdocx convert --to png -o <input>: exit 1, starts with PK Error: output already exists: a.docx
rpptx convert --to pdf -o <input>: exit 0, starts with %PDF- 
rpptx convert --to png -o <input>: exit 1, starts with PK Error: output already exists: a.pptx
rpptx thumbnail -o <input>: exit 0, starts with PNG 
rdocx convert --to pdf, no -o, in.pdf existed: exit 0, in.pdf starts with %PDF-
```

*Acceptance:* every command that writes a file follows the policy the editing commands and `render` already
apply: an output equal to an input is always refused, and an existing output is refused unless an explicit
`--force` is given, with or without `-o`. A test per command asserts that the input, and a pre-existing
output, are byte-identical after a refused run.

Environment: `main` at `9a7ed714` (S75), release build, linux x86_64, python-docx 1.2.0 and python-pptx
1.0.2 for the inputs.
