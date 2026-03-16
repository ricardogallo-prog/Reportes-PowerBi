const fs = require("fs");
const pptxgen = require("pptxgenjs");

const data = JSON.parse(fs.readFileSync("report-data.json", "utf8"));

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.title = data.meta.title;

const C = {
  bg: "1A1A1A",
  orange: "D4621A",
  white: "FFFFFF",
  midGray: "CCCCCC",
  cardBg: "252525"
};

{
  const slide = pres.addSlide();
  slide.background = { color: C.bg };

  slide.addText("SEDE CONVICCIÓN", {
    x: 0.5, y: 1.0, w: 9, h: 0.9,
    fontSize: 30, bold: true, color: C.orange, margin: 0
  });

  slide.addText("QUARTERLY REPORT", {
    x: 0.5, y: 1.7, w: 9, h: 0.5,
    fontSize: 18, bold: true, color: C.white, margin: 0
  });

  slide.addText(data.meta.periodLabel, {
    x: 0.5, y: 2.2, w: 9, h: 0.3,
    fontSize: 11, color: C.midGray, margin: 0
  });

  const stats = [
    { val: String(data.cover.totalSubmissions), label: "Total Submissions" },
    { val: String(data.cover.totalQuotes), label: "Total Quotes" },
    { val: String(data.cover.totalSales), label: "Total Sales" },
    { val: String(data.cover.totalDeclined), label: "Total Declined" }
  ];

  stats.forEach((s, i) => {
    const x = 0.6 + i * 2.2;
    slide.addShape(pres.shapes.RECTANGLE, {
      x, y: 3.0, w: 1.8, h: 1.0,
      fill: { color: C.cardBg },
      line: { color: C.orange, width: 1 }
    });

    slide.addText(s.val, {
      x, y: 3.15, w: 1.8, h: 0.35,
      fontSize: 20, bold: true, color: C.orange, align: "center", margin: 0
    });

    slide.addText(s.label, {
      x, y: 3.55, w: 1.8, h: 0.2,
      fontSize: 8, color: C.midGray, align: "center", margin: 0
    });
  });

  if (fs.existsSync(data.assets.dashboardScreenshot)) {
    slide.addImage({
      path: data.assets.dashboardScreenshot,
      x: 0.5, y: 4.3, w: 4.8, h: 2.5
    });
  }
}

pres.writeFile({ fileName: "SedeConviccion_Q4_2025.pptx" })
  .then(() => console.log("PPTX generado"))
  .catch((err) => {
    console.error(err);
    process.exit(1);
  });
