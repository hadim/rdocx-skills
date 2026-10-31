// Base .pptx made with pptxgenjs (the builtin stack), following the built-in pptx SKILL.md:
// layout set first, two layouts from defineSlideMaster with placeholders, hex colours without '#',
// a table, a native chart, a picture, speaker notes, isTextBox on free text, sections.
// Usage: node pptx_builtin.js OUT.pptx PICTURE.png   (NODE_PATH = common.NODE_MODULES)
const pptxgen = require("pptxgenjs");

const [out, png] = process.argv.slice(2);
const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.title = "Validity base";
pres.author = "bench";
pres.company = "Bench Co";

pres.defineSlideMaster({
  title: "TITLE",
  background: { color: "FFFFFF" },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.5, y: 1.8, w: 9, h: 1.2, fontSize: 40, bold: true, color: "1E2761" }, text: "" } },
    { placeholder: { options: { name: "body", type: "body", x: 0.5, y: 3.1, w: 9, h: 0.8, fontSize: 18, color: "36454F" }, text: "" } },
  ],
});
pres.defineSlideMaster({
  title: "TITLE_ONLY",
  background: { color: "FFFFFF" },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.5, y: 0.3, w: 9, h: 0.8, fontSize: 36, bold: true, color: "1E2761" }, text: "" } },
  ],
  slideNumber: { x: 9.3, y: 5.2, color: "36454F", fontSize: 10 },
});

pres.addSection({ title: "Opening" });
const s1 = pres.addSlide({ masterName: "TITLE", sectionTitle: "Opening" });
s1.addText("Quarterly review", { placeholder: "title" });
s1.addText("Validity benchmark base deck", { placeholder: "body" });
s1.addNotes("Open with the headline number.");

pres.addSection({ title: "Figures" });
const s2 = pres.addSlide({ masterName: "TITLE_ONLY", sectionTitle: "Figures" });
s2.addText("Regional figures", { placeholder: "title" });
s2.addTable(
  [["Region", "Value"], ["North", "41"], ["South", "37"]].map((row, i) =>
    row.map((t) => ({ text: t, options: { bold: i === 0, fill: { color: i === 0 ? "CADCFC" : "FFFFFF" } } }))),
  { x: 0.5, y: 1.3, w: 6, colW: [3, 3], fontSize: 14, color: "36454F", border: { type: "solid", pt: 1, color: "36454F" } },
);
s2.addImage({ path: png, x: 7, y: 1.3, w: 2, h: 1.5, altText: "A grey picture" });
s2.addText("Source: synthetic", { x: 0.5, y: 4.8, w: 5, h: 0.4, fontSize: 10, color: "36454F", isTextBox: true, margin: 0 });
s2.addNotes("Mention the photo record.");

const s3 = pres.addSlide({ masterName: "TITLE_ONLY", sectionTitle: "Figures" });
s3.addText("Costs", { placeholder: "title" });
s3.addChart(pres.charts.BAR, [{ name: "Cost (EUR k)", labels: ["Option A", "Option B", "Option C"], values: [236, 230, 251] }], {
  x: 0.5, y: 1.2, w: 9, h: 4, barDir: "col", showTitle: true, title: "Cost by option",
  showValue: true, dataLabelPosition: "outEnd", chartColors: ["1E2761"], showLegend: false,
  valGridLine: { color: "E0E0E0", size: 1 }, catGridLine: { style: "none" },
});

pres.writeFile({ fileName: out });
