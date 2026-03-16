const fs = require("fs");
const PptxGenJS = require("pptxgenjs");

// ============================================================
// LOAD DATA
// ============================================================

let data = {
  meta: {
    title: "Sede Convicción Q4 2025",
    periodLabel: "Q4 2025 · October — November — December · FULL DETAIL"
  },
  assets: {
    dashboardScreenshot: "dashboard.png"
  },
  cover: {
    totalSubmissions: 301,
    totalQuotes: 349,
    totalSales: 140,
    totalDeclined: 49
  }
};

if (fs.existsSync("report-data.json")) {
  data = JSON.parse(fs.readFileSync("report-data.json", "utf8"));
}

// ============================================================
// PRESENTATION SETUP
// ============================================================

const pres = new PptxGenJS();
pres.layout = "LAYOUT_16x9";
pres.title = data.meta.title;

// COLORS
const C = {
  bg: "1A1A1A",
  orange: "D4621A",
  white: "FFFFFF",
  midGray: "CCCCCC",
  cardBg: "252525"
};

// ============================================================
// SLIDE 1 — COVER
// ============================================================

{
  const slide = pres.addSlide();
  slide.background = { color: C.bg };

  slide.addText("SEDE CONVICCIÓN", {
    x: 0.5,
    y: 1.0,
    w: 9,
    h: 0.9,
    fontSize: 30,
    bold: true,
    color: C.orange
  });

  slide.addText("QUARTERLY REPORT", {
    x: 0.5,
    y: 1.7,
    w: 9,
    h: 0.5,
    fontSize: 18,
    bold: true,
    color: C.white
  });

  slide.addText(data.meta.periodLabel, {
    x: 0.5,
    y: 2.2,
    w: 9,
    h: 0.3,
    fontSize: 11,
    color: C.midGray
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
      x,
      y: 3.0,
      w: 1.8,
      h: 1.0,
      fill: { color: C.cardBg },
      line: { color: C.orange, width: 1 }
    });

    slide.addText(s.val, {
      x,
      y: 3.15,
      w: 1.8,
      h: 0.35,
      fontSize: 20,
      bold: true,
      color: C.orange,
      align: "center"
    });

    slide.addText(s.label, {
      x,
      y: 3.55,
      w: 1.8,
      h: 0.2,
      fontSize: 8,
      color: C.midGray,
      align: "center"
    });
  });

  if (fs.existsSync(data.assets.dashboardScreenshot)) {
    slide.addImage({
      path: data.assets.dashboardScreenshot,
      x: 0.5,
      y: 4.3,
      w: 4.8,
      h: 2.5
    });
  }
}

// ============================================================
// SLIDES 2-20
// ============================================================

// Aquí pegas EXACTAMENTE tus slides actuales.
// NO pongas ningún pres.writeFile aquí.

// ejemplo de slide placeholder

for (let i = 2; i <= 20; i++) {
  const slide = pres.addSlide();
  slide.background = { color: C.bg };

  slide.addText(`Slide ${i}`, {
    x: 1,
    y: 1,
    fontSize: 30,
    color: C.white,
    bold: true
  });
}

// ============================================================
// SAVE FILE
// ============================================================

pres.writeFile({
  fileName: "SedeConviccion_Q4_2025_COMPLETE_FINAL.pptx"
})
.then(() => {
  console.log("\n==============================================");
  console.log("✓ PRESENTATION GENERATED SUCCESSFULLY");
  console.log("File: SedeConviccion_Q4_2025_COMPLETE_FINAL.pptx");
  console.log("==============================================\n");
})
.catch(err => {
  console.error("Error generating presentation:", err);
  process.exit(1);
});
