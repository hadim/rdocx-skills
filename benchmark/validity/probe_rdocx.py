"""Open a file with the rdocx / rpptx Python binding and read its text (the "rdocx_open" checker).

Usage: R_PY probe_rdocx.py FILE   (.docx/.dotx with rdocx, .pptx/.potx with rpptx)
Exit 0 and print a short summary when the file loads and its text reads; exit 1 with the error otherwise.
"""

import sys

path = sys.argv[1]
try:
    if path.lower().endswith((".docx", ".dotx")):
        import rdocx

        doc = rdocx.Document(path)
        text = doc.text()
        tables = [[c.text for row in t.rows for c in row.cells] for t in doc.tables]
        comments = [cm.text for cm in doc.comments]
        print(f"ok: {len(doc.paragraphs)} paragraphs, {len(tables)} tables, {len(comments)} comments, {len(text)} chars")
    else:
        import rpptx

        prs = rpptx.Presentation(path)
        n = 0
        for slide in prs.slides:
            for shape in slide.shapes:
                if shape.has_text_frame:
                    n += len(shape.text_frame.text)
                if shape.has_table:
                    n += sum(len(c.text) for row in shape.table.rows for c in row.cells)
                if shape.has_chart:
                    n += len(shape.chart.categories)
            n += len(slide.notes_text or "")
        print(f"ok: {len(prs.slides)} slides, {n} chars")
except Exception as e:  # any failure to load or read counts as flagged
    print(f"{type(e).__name__}: {e}"[:600])
    sys.exit(1)
