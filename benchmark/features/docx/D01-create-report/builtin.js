const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, TableOfContents, Table, TableRow, TableCell, WidthType,
  ShadingType, ImageRun, Footer, PageNumber, AlignmentType, LevelFormat,
} = require("docx");

const [png, out] = process.argv.slice(2);
const list = (reference, format, text) => ({
  reference,
  levels: [{ level: 0, format, text, alignment: AlignmentType.LEFT,
    style: { paragraph: { indent: { left: 720, hanging: 360 } } } }],
});
const item = (text, reference) => new Paragraph({ numbering: { reference, level: 0 }, children: [new TextRun(text)] });
const heading = (text, heading) => new Paragraph({ heading, children: [new TextRun(text)] });
const rows = [["Region", "Q1", "Q2", "Q3"], ["North", "10", "12", "14"], ["South", "8", "9", "11"],
  ["East", "7", "7", "9"], ["West", "12", "13", "15"]];
const cell = (text, header) => new TableCell({
  width: { size: 2256, type: WidthType.DXA },
  shading: header ? { fill: "D9E2F3", type: ShadingType.CLEAR, color: "auto" } : undefined,
  children: [new Paragraph({ children: [new TextRun({ text, bold: header })] })],
});

const doc = new Document({
  features: { updateFields: true },
  numbering: { config: [list("bullets", LevelFormat.BULLET, "\u2022"), list("numbers", LevelFormat.DECIMAL, "%1.")] },
  sections: [{
    footers: {
      default: new Footer({ children: [new Paragraph({
        children: [new TextRun({ children: ["Page ", PageNumber.CURRENT, " of ", PageNumber.TOTAL_PAGES] })],
      })] }),
    },
    children: [
      heading("Quarterly Report", HeadingLevel.TITLE),
      new TableOfContents("Table of Contents", { hyperlink: true, headingStyleRange: "1-3" }),
      heading("Introduction", HeadingLevel.HEADING_1),
      new Paragraph({ children: [new TextRun("This report has "), new TextRun({ text: "bold findings", bold: true }),
        new TextRun(" and "), new TextRun({ text: "italic remarks", italics: true }), new TextRun(".")] }),
      heading("Scope", HeadingLevel.HEADING_2),
      ...["Alpha", "Beta", "Gamma"].map((t) => item(t, "bullets")),
      heading("Details", HeadingLevel.HEADING_3),
      ...["First", "Second", "Third"].map((t) => item(t, "numbers")),
      heading("Results", HeadingLevel.HEADING_1),
      new Table({
        width: { size: 9024, type: WidthType.DXA },
        columnWidths: [2256, 2256, 2256, 2256],
        rows: rows.map((r, i) => new TableRow({ children: r.map((t) => cell(t, i === 0)) })),
      }),
      new Paragraph({ children: [new ImageRun({ type: "png", data: fs.readFileSync(png),
        transformation: { width: 302, height: 227 } })] }),
    ],
  }],
});
Packer.toBuffer(doc).then((buf) => fs.writeFileSync(out, buf));
