"""D11 input: the sentence to annotate sits in the middle of a paragraph, its end inside a run; no footnotes yet."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from authoring_inputs import declare_rsids, new_document, para  # noqa: E402

doc = new_document("Survey summary")
doc.add_heading("Survey summary", 1)
para(doc, [("We visited each ", ""), ("region", "i"), (". The survey covered ", ""), ("120", "b"),
           (" sites. Results follow below.", "")])
para(doc, [("The results are ", ""), ("preliminary", "i"), (".", "")])
declare_rsids(doc)
doc.save(Path(sys.argv[1]) / "input.docx")
