// Base .docx made with docx-js (the builtin stack), following the built-in docx SKILL.md gotchas:
// US Letter in DXA, dual table widths in DXA, ShadingType.CLEAR, ImageRun with `type`, no "\n".
// Usage: node docx_builtin.js OUT.docx PICTURE.png   (NODE_PATH = common.NODE_MODULES)
const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell, WidthType,
  ShadingType, ImageRun, CommentRangeStart, CommentRangeEnd, CommentReference,
  InsertedTextRun, DeletedTextRun,
} = require("docx");

const [out, png] = process.argv.slice(2);
const when = new Date("2026-10-01T09:00:00Z");

function cell(text, header) {
  return new TableCell({
    width: { size: 4680, type: WidthType.DXA },
    shading: header ? { type: ShadingType.CLEAR, color: "auto", fill: "D9E2F3" } : undefined,
    children: [new Paragraph({ children: [new TextRun({ text, bold: !!header })] })],
  });
}

const doc = new Document({
  creator: "bench",
  title: "Validity base",
  styles: {
    paragraphStyles: [{
      id: "Notebox", name: "Note box", basedOn: "Normal", quickFormat: true,
      run: { italics: true }, paragraph: { indent: { left: 720 } },
    }],
  },
  comments: {
    children: [{
      id: 0, author: "Reviewer", initials: "RV", date: when,
      children: [new Paragraph({ children: [new TextRun("Check this figure.")] })],
    }],
  },
  sections: [{
    properties: { page: { size: { width: 12240, height: 15840 } } },
    children: [
      new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("Quarterly report")] }),
      new Paragraph({ style: "Notebox", children: [new TextRun("A styled note paragraph.")] }),
      new Paragraph({
        children: [
          new CommentRangeStart(0),
          new TextRun("Revenue grew by 12 percent over the quarter."),
          new CommentRangeEnd(0),
          new TextRun({ children: [new CommentReference(0)] }),
        ],
      }),
      new Paragraph({
        children: [
          new TextRun("The budget is "),
          new DeletedTextRun({ text: "ten", id: 101, author: "Reviewer", date: "2026-10-01T09:00:00Z" }),
          new InsertedTextRun({ text: "twelve", id: 102, author: "Reviewer", date: "2026-10-01T09:00:00Z" }),
          new TextRun(" thousand euros."),
        ],
      }),
      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [4680, 4680],
        rows: [["Item", "Value"], ["North", "41"], ["South", "37"]].map(
          (row, i) => new TableRow({ children: row.map((t) => cell(t, i === 0)) }),
        ),
      }),
      new Paragraph({ children: [new TextRun("Figure 1: a picture.")] }),
      new Paragraph({
        children: [new ImageRun({
          type: "png", data: fs.readFileSync(png),
          transformation: { width: 120, height: 90 },
          altText: { title: "Figure", description: "A grey picture", name: "figure" },
        })],
      }),
    ],
  }],
});

Packer.toBuffer(doc).then((buf) => fs.writeFileSync(out, buf));
