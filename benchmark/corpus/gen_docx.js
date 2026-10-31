// Build a .docx from a plan (content.py docx_plan) with docx-js, the way the built-in docx skill
// teaches: A4, built-in heading levels, numbering configs for lists, tables with columnWidths and
// a DXA width on every cell, ShadingType.CLEAR, ImageRun with its type, PageBreak inside a
// Paragraph, a TOC field over the headings, PAGE / NUMPAGES in the footer.
//
//   TZ=UTC NODE_PATH=<node_modules> node gen_docx.js plan.json out.docx
"use strict";
require("./deterministic");
const fs = require("fs");
const {
	AlignmentType, CommentRangeEnd, CommentRangeStart, CommentReference, DeletedTextRun, Document, Footer,
	FootnoteReferenceRun, Header, HeadingLevel, ImageRun, InsertedTextRun, LevelFormat, Packer, PageBreak,
	PageNumber, PageOrientation, Paragraph, ShadingType, Table, TableCell, TableOfContents, TableRow, TextRun,
	VerticalAlign, WidthType, BorderStyle,
} = require("docx");

const [planPath, outPath] = process.argv.slice(2);
const plan = JSON.parse(fs.readFileSync(planPath, "utf8"));

const HEAD_FILL = "1F4E79";
const ZEBRA = "EAF1F8";
const TOTAL_FILL = "D9E2F3";

function inlines(list) {
	const out = [];
	for (const x of list) {
		const fmt = { bold: !!x.b, italics: !!x.i };
		switch (x.k) {
			case "r":
				out.push(new TextRun({ text: x.text, ...fmt }));
				break;
			case "ins":
				out.push(new InsertedTextRun({ text: x.text, id: x.id, author: x.author, date: x.date, ...fmt }));
				break;
			case "del":
				out.push(new DeletedTextRun({ text: x.text, id: x.id, author: x.author, date: x.date, ...fmt }));
				break;
			case "fn":
				out.push(new FootnoteReferenceRun(x.id));
				break;
			case "cs":
				out.push(new CommentRangeStart(x.id));
				break;
			case "ce":
				out.push(new CommentRangeEnd(x.id));
				out.push(new TextRun({ children: [new CommentReference(x.id)] }));
				break;
			default:
				throw new Error("unknown inline " + x.k);
		}
	}
	return out;
}

function table(spec) {
	const total = spec.widths.reduce((a, b) => a + b, 0);
	const ncols = spec.widths.length;
	const carry = new Array(ncols).fill(0); // rows still covered by a vertical merge from above
	const rows = spec.rows.map((r) => {
		let col = 0;
		const cells = [];
		for (const c of r.cells) {
			while (col < ncols && carry[col] > 0) {
				carry[col] -= 1;
				col += 1;
			}
			const span = c.colSpan || 1;
			const w = spec.widths.slice(col, col + span).reduce((a, b) => a + b, 0);
			if (c.rowSpan) for (let k = col; k < col + span; k++) carry[k] = c.rowSpan - 1;
			const fill = r.header ? HEAD_FILL : r.total ? TOTAL_FILL : r.shade ? ZEBRA : null;
			cells.push(
				new TableCell({
					width: { size: w, type: WidthType.DXA },
					columnSpan: span > 1 ? span : undefined,
					rowSpan: c.rowSpan,
					verticalAlign: VerticalAlign.CENTER,
					shading: fill ? { fill, type: ShadingType.CLEAR, color: "auto" } : undefined,
					margins: { top: 40, bottom: 40, left: 80, right: 80 },
					children: [
						new Paragraph({
							alignment: r.header ? AlignmentType.CENTER : c.align === "right" ? AlignmentType.RIGHT : AlignmentType.LEFT,
							spacing: { after: 0 },
							children: [new TextRun({ text: c.text, bold: !!(r.header || r.total), color: r.header ? "FFFFFF" : undefined, size: 18 })],
						}),
					],
				}),
			);
			col += span;
		}
		while (col < ncols && carry[col] > 0) {
			carry[col] -= 1;
			col += 1;
		}
		return new TableRow({ tableHeader: !!r.header, children: cells });
	});
	const border = { style: BorderStyle.SINGLE, size: 4, color: "8EAADB" };
	return new Table({
		width: { size: total, type: WidthType.DXA },
		columnWidths: spec.widths,
		borders: { top: border, bottom: border, left: border, right: border, insideHorizontal: border, insideVertical: border },
		rows,
	});
}

function blocks(list) {
	const out = [];
	for (const b of list) {
		switch (b.t) {
			case "title":
				out.push(new Paragraph({ heading: HeadingLevel.TITLE, spacing: { before: 2400 }, children: [new TextRun(b.text)] }));
				break;
			case "subtitle":
				out.push(new Paragraph({ style: "Subtitle", children: [new TextRun(b.text)] }));
				break;
			case "pagebreak":
				out.push(new Paragraph({ children: [new PageBreak()] }));
				break;
			case "toc":
				out.push(new Paragraph({ style: "TOCHeading", children: [new TextRun("Contents")] }));
				out.push(new TableOfContents("Contents", { hyperlink: true, headingStyleRange: "1-3" }));
				break;
			case "h":
				out.push(new Paragraph({ heading: [HeadingLevel.HEADING_1, HeadingLevel.HEADING_2, HeadingLevel.HEADING_3][b.level - 1], children: [new TextRun(b.text)] }));
				break;
			case "p":
				out.push(new Paragraph({ children: inlines(b.inl) }));
				break;
			case "li":
				out.push(new Paragraph({ numbering: { reference: b.list === "num" ? "numbers" : "bullets", level: b.level, instance: b.instance }, children: inlines(b.inl) }));
				break;
			case "table":
				out.push(table(b));
				break;
			case "caption":
				out.push(new Paragraph({ style: "Caption", keepNext: true, children: [new TextRun(b.text)] }));
				break;
			case "img":
				out.push(
					new Paragraph({
						alignment: AlignmentType.CENTER,
						keepNext: true,
						children: [
							new ImageRun({
								type: "png",
								data: fs.readFileSync(b.path),
								transformation: { width: b.w, height: b.h },
								altText: { title: b.alt, description: b.alt, name: b.alt },
							}),
						],
					}),
				);
				break;
			default:
				throw new Error("unknown block " + b.t);
		}
	}
	return out;
}

const header = () =>
	new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: plan.meta.header, size: 16, color: "595959" })] })] });
const footer = () =>
	new Footer({
		children: [
			new Paragraph({
				alignment: AlignmentType.CENTER,
				children: [new TextRun({ size: 18, children: ["Page ", PageNumber.CURRENT, " of ", PageNumber.TOTAL_PAGES] })],
			}),
		],
	});

const A4 = { width: 11906, height: 16838 };
const sections = plan.sections.map((s, k) => ({
	properties: {
		page: {
			size: s.orientation === "landscape" ? { ...A4, orientation: PageOrientation.LANDSCAPE } : { ...A4 },
			margin: { top: 1440, bottom: 1440, left: 1300, right: 1300 },
		},
	},
	headers: k === 0 ? { default: header() } : undefined,
	footers: k === 0 ? { default: footer() } : undefined,
	children: blocks(s.blocks),
}));

const footnotes = {};
for (const [id, text] of Object.entries(plan.footnotes)) footnotes[id] = { children: [new Paragraph({ children: [new TextRun(text)] })] };

const bullet = (level, text) => ({
	level,
	format: LevelFormat.BULLET,
	text,
	alignment: AlignmentType.LEFT,
	style: { paragraph: { indent: { left: 720 * (level + 1), hanging: 360 } } },
});

const doc = new Document({
	creator: plan.meta.creator,
	title: plan.meta.title,
	description: "Synthetic benchmark document (fictional organisation).",
	lastModifiedBy: plan.meta.creator,
	styles: {
		default: {
			document: { run: { font: "Calibri", size: 22 }, paragraph: { spacing: { after: 120, line: 264 } } },
			title: { run: { font: "Cambria", size: 56, color: HEAD_FILL } },
			heading1: { run: { font: "Cambria", size: 32, bold: true, color: HEAD_FILL }, paragraph: { spacing: { before: 360, after: 120 }, keepNext: true } },
			heading2: { run: { font: "Cambria", size: 26, bold: true, color: "2E74B5" }, paragraph: { spacing: { before: 240, after: 80 }, keepNext: true } },
			heading3: { run: { font: "Calibri", size: 22, bold: true, italics: true, color: "404040" }, paragraph: { spacing: { before: 200, after: 60 }, keepNext: true } },
		},
		paragraphStyles: [
			{ id: "Subtitle", name: "Subtitle", basedOn: "Normal", next: "Normal", run: { font: "Cambria", size: 32, color: "595959" }, paragraph: { spacing: { after: 480 } } },
			{ id: "Caption", name: "Caption", basedOn: "Normal", next: "Normal", run: { size: 18, italics: true, color: "404040" }, paragraph: { spacing: { before: 60, after: 160 } } },
			{ id: "TOCHeading", name: "TOC Heading", basedOn: "Normal", next: "Normal", run: { font: "Cambria", size: 32, bold: true, color: HEAD_FILL }, paragraph: { spacing: { after: 240 } } },
		],
	},
	numbering: {
		config: [
			{ reference: "bullets", levels: [bullet(0, "\u2022"), bullet(1, "\u2013")] },
			{
				reference: "numbers",
				levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }],
			},
		],
	},
	comments: { children: plan.comments.map((c) => ({ id: c.id, author: c.author, initials: c.initials, date: new Date(c.date), children: [new Paragraph({ children: [new TextRun(c.text)] })] })) },
	footnotes,
	sections,
});

Packer.toBuffer(doc).then((buf) => fs.writeFileSync(outPath, buf));
