// bench: tools=pptxgenjs, apply_theme.js; processes=1; raw_xml=no
const fs = require("fs");
const path = require("path");
const pptxgen = require("pptxgenjs");
const { applyTheme } = require(path.join(process.env.PPTX_SKILL, "scripts/apply_theme.js"));

const [src, out] = process.argv.slice(2);
const content = JSON.parse(fs.readFileSync(src, "utf8"));
const T = content.theme;
const THEME = { name: T.name, headFontFace: T.headFont, bodyFontFace: T.bodyFont, colors: T.colors };

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.33 x 7.5 in
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
const C = pres.SchemeColor;
const num = () => ({ x: 12.2, y: 6.95, w: 0.8, h: 0.35, fontSize: 11, color: C.text2, align: "right" });
const title = (y = 0.4, h = 1.0) => ({ placeholder: { options: { name: "title", type: "title", x: 0.6, y, w: 12.1, h,
  fontSize: 36, bold: true, color: C.text2 }, text: "" } });
const body = (name, x, w, y = 1.7, h = 4.9) => ({ placeholder: { options: { name, type: "body", x, y, w, h, fontSize: 18,
  color: C.text1, valign: "top" }, text: "" } });

pres.defineSlideMaster({ title: "Title Slide", background: { color: C.accent1 }, objects: [
  { placeholder: { options: { name: "title", type: "title", x: 0.8, y: 2.4, w: 11.7, h: 1.4, fontSize: 44, bold: true,
    color: C.background1, align: "center" }, text: "" } },
  { placeholder: { options: { name: "subtitle", type: "body", x: 0.8, y: 3.9, w: 11.7, h: 0.8, fontSize: 20,
    color: C.background2, align: "center" }, text: "" } }] });
pres.defineSlideMaster({ title: "Section Header", background: { color: C.background2 }, slideNumber: num(), objects: [
  title(2.6, 1.2), body("subtitle", 0.6, 12.1, 3.9, 0.8)] });
pres.defineSlideMaster({ title: "Title and Content", slideNumber: num(), objects: [title(), body("body", 0.6, 12.1)] });
pres.defineSlideMaster({ title: "Title Only", slideNumber: num(), objects: [title()] });
pres.defineSlideMaster({ title: "Two Content", slideNumber: num(), objects: [title(), body("left", 0.6, 5.9),
  body("right", 6.8, 5.9)] });
pres.defineSlideMaster({ title: "Picture with Caption", slideNumber: num(), objects: [title(),
  { placeholder: { options: { name: "picture", type: "pic", x: 0.6, y: 1.6, w: 7.6, h: 5.0, color: C.text1 }, text: "" } },
  body("caption", 8.6, 4.1, 1.6, 5.0)] });

const bullets = (items) => items.map((t, i) => ({ text: t, options: { bullet: true, breakLine: i < items.length - 1,
  paraSpaceAfter: 8 } }));
const LAYOUT = { title: "Title Slide", section: "Section Header", closing: "Section Header", bullets: "Title and Content",
  table: "Title Only", chart: "Title Only", two_columns: "Two Content", picture: "Picture with Caption" };
let section = "Introduction";
for (const s of content.slides) {
  if (s.kind === "section") section = s.title;
  if (s.kind === "closing") section = "Close";
  if (s.kind === "title" || s.kind === "section" || s.kind === "closing") pres.addSection({ title: section });
  const slide = pres.addSlide({ masterName: LAYOUT[s.kind], sectionTitle: section });
  slide.addText(s.title, { placeholder: "title" });
  if (s.subtitle) slide.addText(s.subtitle, { placeholder: "subtitle" });
  if (s.kind === "bullets") slide.addText(bullets(s.bullets), { placeholder: "body" });
  if (s.kind === "two_columns") {
    slide.addText(bullets(s.left), { placeholder: "left" });
    slide.addText(bullets(s.right), { placeholder: "right" });
  }
  if (s.kind === "table") {
    const rows = s.table.map((row, r) => row.map((text) => ({ text, options: r === 0
      ? { bold: true, color: C.background1, fill: { color: C.accent1 } } : { color: C.text1 } })));
    slide.addTable(rows, { x: 0.6, y: 1.7, w: 12.1, fontSize: 16, border: { type: "solid", pt: 1, color: "D9D9D9" },
      objectName: "Results table" });
  }
  if (s.kind === "chart") {
    const data = s.chart.series.map((ser) => ({ name: ser.name, labels: s.chart.categories, values: ser.values }));
    slide.addChart(pres.charts.BAR, data, { x: 0.6, y: 1.6, w: 12.1, h: 5.2, barDir: "col", barGrouping: "clustered",
      showTitle: true, title: s.title, titleFontFace: "+mj-lt", showValue: true, dataLabelPosition: "outEnd",
      chartColors: [T.colors.accent1, T.colors.accent2], showLegend: true, legendPos: "b", legendFontFace: "+mn-lt",
      catAxisLabelFontFace: "+mn-lt", valAxisLabelFontFace: "+mn-lt", dataLabelFontFace: "+mn-lt",
      valGridLine: { color: "D9D9D9", size: 0.5 }, catGridLine: { style: "none" }, objectName: "Revenue chart" });
  }
  if (s.kind === "picture") {
    slide.addImage({ path: path.join(path.dirname(src), s.image), placeholder: "picture", x: 0.6, y: 1.6, w: 7.6, h: 5.0,
      sizing: { type: "cover", w: 7.6, h: 5.0 }, altText: s.caption });
    slide.addText(s.caption, { placeholder: "caption" });
  }
  slide.addNotes(s.notes);
}

(async () => {
  await pres.writeFile({ fileName: out });
  await applyTheme(out, THEME);
})();
