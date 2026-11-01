// Build a .pptx from a plan (content.py deck_plan) with pptxgenjs as a structured deck, the way
// the built-in pptx skill teaches: a THEME object, pres.theme fonts, scheme colors, one layout per
// slide frame (defineSlideMaster) with named placeholders, footer and slide number on the layouts,
// sections, speaker notes on every slide, native charts with their labels set, then
// writeFile() and the skill's applyTheme().
//
//   TZ=UTC NODE_PATH=<node_modules> node gen_pptx.js plan.json out.pptx <builtin pptx skill>/scripts/apply_theme.js
"use strict";
require("./deterministic");
const fs = require("fs");
const pptxgen = require("pptxgenjs");

const [planPath, outPath, applyThemePath] = process.argv.slice(2);
const plan = JSON.parse(fs.readFileSync(planPath, "utf8"));
const { applyTheme } = require(applyThemePath);

const THEME = {
	name: "Harbourside",
	headFontFace: "Cambria",
	bodyFontFace: "Calibri",
	colors: {
		dk1: "1B2A3A", lt1: "FFFFFF", dk2: "0B4F6C", lt2: "EEF3F6",
		accent1: "0B7A75", accent2: "E07A1F", accent3: "5B8E7D", accent4: "8C5383", accent5: "C9A227", accent6: "3D6FB6",
		hlink: "0B4F6C", folHlink: "8C5383",
	},
};
const HEX = THEME.colors;

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.33 x 7.5 in
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.title = plan.meta.title;
pres.company = plan.meta.company.replace(/&/g, "&amp;");
pres.author = "Corporate Affairs";
const C = pres.SchemeColor;

const footer = () => ({ text: { text: plan.meta.footer, options: { x: 0.6, y: 6.95, w: 9.0, h: 0.35, fontSize: 10, color: C.text2 } } });
const slideNumber = () => ({ x: 12.0, y: 6.95, w: 0.8, h: 0.35, fontSize: 10, color: C.text2, align: "right" });
const titlePh = () => ({ placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.35, w: 12.1, h: 0.95, fontSize: 36, bold: true, color: C.text2, valign: "middle" }, text: "" } });

pres.defineSlideMaster({
	title: "Title (dark)",
	background: { color: C.text2 },
	objects: [
		{ placeholder: { options: { name: "title", type: "title", x: 0.8, y: 2.3, w: 11.7, h: 1.5, fontSize: 44, bold: true, color: C.background1 }, text: "" } },
		{ placeholder: { options: { name: "subtitle", type: "body", x: 0.8, y: 3.9, w: 11.7, h: 0.9, fontSize: 22, color: C.background2 }, text: "" } },
	],
});
pres.defineSlideMaster({
	title: "Section divider",
	background: { color: C.accent1 },
	objects: [
		{ placeholder: { options: { name: "kicker", type: "body", x: 0.8, y: 2.1, w: 11.7, h: 0.6, fontSize: 20, color: C.background2 }, text: "" } },
		{ placeholder: { options: { name: "title", type: "title", x: 0.8, y: 2.7, w: 11.7, h: 1.6, fontSize: 40, bold: true, color: C.background1 }, text: "" } },
	],
	slideNumber: { x: 12.0, y: 6.95, w: 0.8, h: 0.35, fontSize: 10, color: C.background1, align: "right" },
});
pres.defineSlideMaster({
	title: "Title only",
	background: { color: C.background1 },
	margin: [0.5, 0.5, 0.7, 0.5],
	objects: [titlePh(), footer()],
	slideNumber: slideNumber(),
});
pres.defineSlideMaster({
	title: "Title and body",
	background: { color: C.background1 },
	objects: [
		titlePh(),
		{ placeholder: { options: { name: "body", type: "body", x: 0.6, y: 1.5, w: 12.1, h: 5.3, fontSize: 18, color: C.text1, valign: "top" }, text: "" } },
		footer(),
	],
	slideNumber: slideNumber(),
});
pres.defineSlideMaster({
	title: "Body and picture",
	background: { color: C.background1 },
	objects: [
		titlePh(),
		{ placeholder: { options: { name: "body", type: "body", x: 0.6, y: 1.5, w: 6.3, h: 5.3, fontSize: 18, color: C.text1, valign: "top" }, text: "" } },
		footer(),
	],
	slideNumber: slideNumber(),
});
pres.defineSlideMaster({
	title: "Two columns",
	background: { color: C.background1 },
	objects: [
		titlePh(),
		{ placeholder: { options: { name: "left", type: "body", x: 0.6, y: 1.5, w: 5.9, h: 5.3, fontSize: 16, color: C.text1, valign: "top" }, text: "" } },
		{ placeholder: { options: { name: "right", type: "body", x: 6.8, y: 1.5, w: 5.9, h: 5.3, fontSize: 16, color: C.text1, valign: "top" }, text: "" } },
		footer(),
	],
	slideNumber: slideNumber(),
});
pres.defineSlideMaster({
	title: "Chart and commentary",
	background: { color: C.background1 },
	objects: [
		titlePh(),
		{ placeholder: { options: { name: "chart", type: "chart", x: 0.6, y: 1.45, w: 8.0, h: 5.3, color: C.text1 }, text: "" } },
		{ placeholder: { options: { name: "commentary", type: "body", x: 8.9, y: 1.6, w: 3.8, h: 5.1, fontSize: 16, color: C.text1, valign: "top" }, text: "" } },
		footer(),
	],
	slideNumber: slideNumber(),
});

const bullets = (items, extra = {}) =>
	items.map((t, k) => ({ text: t, options: { bullet: true, paraSpaceAfter: 8, ...extra, breakLine: k < items.length - 1 } }));

const CHART_COLORS = [HEX.accent1, HEX.accent2, HEX.accent3, HEX.accent4, HEX.accent5, HEX.accent6];

function chartOptions(spec) {
	const o = {
		placeholder: "chart",
		showTitle: true,
		title: spec.title,
		titleFontFace: "+mn-lt",
		titleFontSize: 16,
		titleColor: HEX.dk2,
		catAxisLabelFontFace: "+mn-lt",
		valAxisLabelFontFace: "+mn-lt",
		dataLabelFontFace: "+mn-lt",
		legendFontFace: "+mn-lt",
		catAxisLabelColor: HEX.dk1,
		valAxisLabelColor: HEX.dk1,
		valGridLine: { color: "D9D9D9", size: 0.5 },
		catGridLine: { style: "none" },
		chartColors: CHART_COLORS.slice(0, Math.max(spec.series.length, spec.type === "pie" ? spec.series[0].labels.length : 1)),
		showLegend: spec.series.length > 1 || spec.type === "pie",
		legendPos: "b",
	};
	if (spec.type === "bar") Object.assign(o, { barDir: "col", barGrouping: "clustered", showValue: true, dataLabelPosition: "outEnd", dataLabelFontSize: 10 });
	if (spec.type === "line") Object.assign(o, { lineSize: 2, lineDataSymbol: "circle", lineDataSymbolSize: 6 });
	if (spec.type === "pie") Object.assign(o, { showPercent: true, showValue: false, dataLabelPosition: "bestFit", dataLabelColor: HEX.lt1 });
	return o;
}

let section = null;
for (const s of plan.slides) {
	if (s.section !== section) {
		pres.addSection({ title: s.section });
		section = s.section;
	}
	const add = (masterName) => pres.addSlide({ masterName, sectionTitle: s.section });
	let slide;
	switch (s.type) {
		case "title":
			slide = add("Title (dark)");
			slide.addText(s.title, { placeholder: "title" });
			slide.addText(s.subtitle, { placeholder: "subtitle" });
			break;
		case "section":
			slide = add("Section divider");
			slide.addText(s.subtitle, { placeholder: "kicker" });
			slide.addText(s.title, { placeholder: "title" });
			break;
		case "bullets":
		case "closing":
			slide = add("Title and body");
			slide.addText(s.title, { placeholder: "title" });
			slide.addText(bullets(s.bullets), { placeholder: "body" });
			break;
		case "bullets_pic": {
			slide = add("Body and picture");
			slide.addText(s.title, { placeholder: "title" });
			slide.addText(bullets(s.bullets, { fontSize: 16 }), { placeholder: "body" });
			const w = 5.6;
			const h = Math.min(5.0, (w * s.image.h) / s.image.w);
			slide.addImage({ path: s.image.path, x: 7.2, y: 1.6, w, h, objectName: "Figure", altText: "Chart from the report" });
			break;
		}
		case "two_col":
			slide = add("Two columns");
			slide.addText(s.title, { placeholder: "title" });
			slide.addText([{ text: s.left_head, options: { bold: true, breakLine: true, color: C.accent1 } }, ...bullets(s.bullets, { fontSize: 15 })], { placeholder: "left" });
			slide.addText([{ text: s.right_head, options: { bold: true, breakLine: true, color: C.accent2 } }, ...bullets(s.right, { fontSize: 15 })], { placeholder: "right" });
			break;
		case "chart": {
			slide = add("Chart and commentary");
			slide.addText(s.title, { placeholder: "title" });
			const type = { bar: pres.charts.BAR, line: pres.charts.LINE, pie: pres.charts.PIE }[s.chart.type];
			slide.addChart(type, s.chart.series, chartOptions(s.chart));
			slide.addText(bullets(s.commentary, { fontSize: 15 }), { placeholder: "commentary" });
			break;
		}
		case "table": {
			slide = add("Title only");
			slide.addText(s.title, { placeholder: "title" });
			const rows = s.table.map((r, i) =>
				r.map((t, j) => ({
					text: t,
					options: i === 0
						? { bold: true, color: C.background1, fill: { color: C.accent1 }, align: j === 0 ? "left" : "center" }
						: { color: C.text1, fill: { color: i % 2 === 0 ? C.background2 : C.background1 }, align: j === 0 ? "left" : "right" },
				})),
			);
			slide.addTable(rows, { x: 0.6, y: 1.5, w: 12.1, fontSize: 14, border: { type: "solid", pt: 0.5, color: "BFBFBF" }, rowH: 0.45, valign: "middle", objectName: "Data table" });
			break;
		}
		case "stats": {
			slide = add("Title only");
			slide.addText(s.title, { placeholder: "title" });
			s.stats.forEach((st, k) => {
				const x = 0.6 + k * 4.1;
				slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 1.7, w: 3.8, h: 3.2, fill: { color: C.background2 }, line: { color: C.background2 }, rectRadius: 0.12, objectName: `Stat card ${k + 1}` });
				slide.addText(st.value, { x, y: 2.0, w: 3.8, h: 1.4, fontSize: 60, bold: true, color: C.accent1, align: "center", isTextBox: true, objectName: `Stat value ${k + 1}` });
				slide.addText(st.label, { x: x + 0.2, y: 3.5, w: 3.4, h: 1.2, fontSize: 16, color: C.text1, align: "center", valign: "top", isTextBox: true, objectName: `Stat label ${k + 1}` });
			});
			slide.addText(s.caption, { x: 0.6, y: 5.3, w: 12.1, h: 1.2, fontSize: 16, italic: true, color: C.text2, isTextBox: true, objectName: "Caption" });
			break;
		}
		case "long_text":
			slide = add("Title and body");
			slide.addText(s.title, { placeholder: "title" });
			slide.addText(s.text.map((t, k) => ({ text: t, options: { fontSize: 14, paraSpaceAfter: 10, breakLine: k < s.text.length - 1 } })), { placeholder: "body" });
			break;
		default:
			throw new Error("unknown slide type " + s.type);
	}
	slide.addNotes(s.notes);
}

(async () => {
	await pres.writeFile({ fileName: outPath });
	await applyTheme(outPath, THEME);
})();
