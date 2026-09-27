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
