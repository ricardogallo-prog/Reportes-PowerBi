const pptxgen = require("pptxgenjs");

const fs = require("fs");

const branch = process.env.BRANCH_OFFICE.replace(/\s/g,"_");
const data = JSON.parse(
  fs.readFileSync(`report-clean-${branch}.json`)
);

// Color palette
const C = {
  bg: "1A1A1A",
  orange: "D4621A",
  orangeLight: "E87B35",
  orangeDark: "B04F10",
  white: "FFFFFF",
  lightGray: "F5F5F5",
  midGray: "CCCCCC",
  darkGray: "444444",
  textDark: "1A1A1A",
  green: "27AE60",
  red: "E74C3C",
  amber: "F39C12",
  tableHeader: "D4621A",
  tableRowAlt: "2A2A2A",
  tableRowBase: "222222",
  cardBg: "252525",
};

const makeShadow = () => ({ type: "outer", blur: 8, offset: 2, angle: 135, color: "000000", opacity: 0.25 });

// ============================================================
// SLIDE 1 — Cover
// ============================================================
{
  const slide = pres.addSlide();
  slide.background = { color: C.bg };

  // Left orange accent bar
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.25, h: 5.625, fill: { color: C.orange }, line: { color: C.orange } });

  // Top right accent
  slide.addShape(pres.shapes.RECTANGLE, { x: 7.5, y: 0, w: 2.5, h: 0.12, fill: { color: C.orange }, line: { color: C.orange } });

  // Main title
  slide.addText("SEDE CONVICCIÓN", {
    x: 0.5, y: 1.0, w: 9, h: 0.9,
    fontSize: 44, bold: true, color: C.orange, fontFace: "Arial Black",
    charSpacing: 4, margin: 0
  });

  slide.addText("QUARTERLY REPORT", {
    x: 0.5, y: 1.95, w: 9, h: 0.55,
    fontSize: 26, bold: true, color: C.white, fontFace: "Arial",
    charSpacing: 6, margin: 0
  });

  slide.addText("Q4 2025 · October — November — December · FULL DETAIL", {
    x: 0.5, y: 2.6, w: 9, h: 0.4,
    fontSize: 14, color: C.midGray, fontFace: "Calibri", margin: 0
  });

  // Divider line
  slide.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 3.15, w: 9, h: 0.04, fill: { color: C.orange }, line: { color: C.orange } });

  // Stats row
  const stats = [
    { val: "301", label: "Total Submissions" },
    { val: "349", label: "Total Quotes" },
    { val: "140", label: "Total Sales" },
    { val: "49", label: "Total Declined" },
  ];
  const sw = 2.1, sx = 0.5;
  stats.forEach((s, i) => {
    const x = sx + i * sw;
    slide.addShape(pres.shapes.RECTANGLE, { x, y: 3.4, w: 1.9, h: 1.3, fill: { color: C.cardBg }, line: { color: C.orange, width: 1 }, shadow: makeShadow() });
    slide.addText(s.val, { x, y: 3.5, w: 1.9, h: 0.6, fontSize: 32, bold: true, color: C.orange, align: "center", fontFace: "Arial Black", margin: 0 });
    slide.addText(s.label, { x, y: 4.1, w: 1.9, h: 0.4, fontSize: 10, color: C.midGray, align: "center", fontFace: "Calibri", margin: 0 });
  });

  // Footer info
  slide.addText("5 Sales Leaders · 19+ Producers · 5 Carriers (top) · 5 MGAs (top) · Branch: Convicción · Year: 2025", {
    x: 0.5, y: 5.1, w: 9, h: 0.3,
    fontSize: 9, color: "777777", align: "center", fontFace: "Calibri", margin: 0
  });
}

// ============================================================
// SLIDE 2 — Executive Summary
// ============================================================
{
  const slide = pres.addSlide();
  slide.background = { color: C.bg };

  // Header bar
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.65, fill: { color: C.orange }, line: { color: C.orange } });
  slide.addText("Executive Summary — Sede Convicción Q4 2025 · Oct–Nov–Dec", {
    x: 0.3, y: 0.05, w: 9.4, h: 0.55, fontSize: 16, bold: true, color: C.white, fontFace: "Arial", margin: 0
  });

  // Top KPI row
  const kpis = [
    { val: "301", label: "Submission from\nProducer to QC" },
    { val: "243", label: "Submissions\nto MGA" },
    { val: "349", label: "Quotes\nCount" },
    { val: "140", label: "Sales\nCount" },
    { val: "59.74%", label: "Submitted to\nQuoted Ratio" },
    { val: "40.11%", label: "Quote to\nSold Ratio" },
  ];
  const kw = 1.52, ky = 0.78;
  kpis.forEach((k, i) => {
    const x = 0.18 + i * (kw + 0.1);
    slide.addShape(pres.shapes.RECTANGLE, { x, y: ky, w: kw, h: 0.95, fill: { color: C.cardBg }, line: { color: C.orangeDark, width: 1 } });
    slide.addText(k.val, { x, y: ky + 0.04, w: kw, h: 0.5, fontSize: 18, bold: true, color: C.orange, align: "center", fontFace: "Arial Black", margin: 0 });
    slide.addText(k.label, { x, y: ky + 0.5, w: kw, h: 0.42, fontSize: 8, color: C.midGray, align: "center", fontFace: "Calibri", margin: 0 });
  });

  // Second row KPIs
  const kpis2 = [
    { val: "49", label: "USDOT\nDeclined" },
    { val: "$5.90M", label: "Pure Premium\n(Quoted)" },
    { val: "$1.90M", label: "Pure Premium\n(New Sale)" },
    { val: "$3.94M", label: "Combined\nPremium" },
    { val: "32.13%", label: "Prem.\nQuote→ Sale" },
    { val: "20.16%", label: "Declination\nRatio" },
  ];
  const ky2 = 1.82;
  kpis2.forEach((k, i) => {
    const x = 0.18 + i * (kw + 0.1);
    slide.addShape(pres.shapes.RECTANGLE, { x, y: ky2, w: kw, h: 0.95, fill: { color: C.cardBg }, line: { color: C.orangeDark, width: 1 } });
    slide.addText(k.val, { x, y: ky2 + 0.04, w: kw, h: 0.5, fontSize: 18, bold: true, color: C.orange, align: "center", fontFace: "Arial Black", margin: 0 });
    slide.addText(k.label, { x, y: ky2 + 0.5, w: kw, h: 0.42, fontSize: 8, color: C.midGray, align: "center", fontFace: "Calibri", margin: 0 });
  });

  // Left column — Goal & Funnel
  // Goal card
  slide.addShape(pres.shapes.RECTANGLE, { x: 0.18, y: 2.92, w: 4.7, h: 0.55, fill: { color: C.cardBg }, line: { color: C.orange, width: 1 } });
  slide.addText([
    { text: "$3.94M  ", options: { bold: true, color: C.orange, fontSize: 13 } },
    { text: "Q4 Combined Premium  ", options: { color: C.white, fontSize: 11 } },
    { text: "(Goal: $6.45M → 61.2%)", options: { color: C.midGray, fontSize: 10 } },
  ], { x: 0.25, y: 2.93, w: 4.6, h: 0.5, fontFace: "Calibri", margin: 0 });

  // Highlights
  const highlights = [
    { label: "Dec: Best", detail: "December top month · $718K New Sale · 53 sales · 41.73% Q2S" },
    { label: "Pegaso #1", detail: "Top Carrier · 79.31% Q2S" },
    { label: "Nexus #1", detail: "MGA · 56.79% Q2S" },
  ];
  highlights.forEach((h, i) => {
    const y = 3.6 + i * 0.45;
    slide.addShape(pres.shapes.RECTANGLE, { x: 0.18, y, w: 0.06, h: 0.3, fill: { color: C.orange }, line: { color: C.orange } });
    slide.addText(h.label, { x: 0.35, y, w: 1.2, h: 0.3, fontSize: 10, bold: true, color: C.orange, fontFace: "Calibri", margin: 0 });
    slide.addText(h.detail, { x: 1.6, y, w: 3.3, h: 0.3, fontSize: 10, color: C.midGray, fontFace: "Calibri", margin: 0 });
  });

  // Right column — Funnel
  slide.addText("Operational Funnel — Q4 Cumulative", {
    x: 5.1, y: 2.92, w: 4.7, h: 0.35, fontSize: 11, bold: true, color: C.orange, fontFace: "Arial", margin: 0
  });

  const funnel = [
    { label: "Submission from Producer to QC", val: 301, pct: 1.0 },
    { label: "Submissions to MGA", val: 243, pct: 0.81 },
    { label: "Formal Quote", val: 349, pct: 0.87 },
    { label: "Sales", val: 140, pct: 0.47 },
  ];
  const maxW = 4.5;
  funnel.forEach((f, i) => {
    const y = 3.35 + i * 0.46;
    const barW = maxW * f.pct;
    slide.addShape(pres.shapes.RECTANGLE, { x: 5.1, y, w: barW, h: 0.32, fill: { color: i === 3 ? C.orange : C.orangeDark }, line: { color: C.orange } });
    slide.addText(`${f.label}  ${f.val}`, { x: 5.15, y: y + 0.04, w: barW - 0.1, h: 0.24, fontSize: 9, bold: i === 3, color: C.white, fontFace: "Calibri", margin: 0 });
  });

  // Premium overview
  slide.addShape(pres.shapes.RECTANGLE, { x: 5.1, y: 5.05, w: 4.7, h: 0.52, fill: { color: C.cardBg }, line: { color: C.orangeDark, width: 1 } });
  slide.addText("Q4 Premium Overview (USD M)", {
    x: 5.15, y: 5.06, w: 2.0, h: 0.2, fontSize: 7.5, bold: true, color: C.midGray, fontFace: "Calibri", margin: 0
  });
  const prems = [{ v: "$5.9M", l: "Quoted" }, { v: "$1.90M", l: "New Sale" }, { v: "$3.94M", l: "Combined" }];
  prems.forEach((p, i) => {
    const x = 5.15 + i * 1.55;
    slide.addText(p.v, { x, y: 5.24, w: 1.5, h: 0.2, fontSize: 12, bold: true, color: C.orange, fontFace: "Arial Black", align: "center", margin: 0 });
    slide.addText(p.l, { x, y: 5.44, w: 1.5, h: 0.12, fontSize: 7, color: C.midGray, fontFace: "Calibri", align: "center", margin: 0 });
  });
}

// ============================================================
// SLIDE 3 — Monthly Drill-Down
// ============================================================
{
  const slide = pres.addSlide();
  slide.background = { color: C.bg };

  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.65, fill: { color: C.orange }, line: { color: C.orange } });
  slide.addText("Monthly Sold Premium — Drill-Down Detail Q4 2025 · Oct–Nov–Dec", {
    x: 0.3, y: 0.05, w: 9.4, h: 0.55, fontSize: 16, bold: true, color: C.white, fontFace: "Arial", margin: 0
  });

  const months = [
    {
      name: "October · Month 10", bg: C.orange,
      kpis: [
        ["New Sale (Pure Prem.)", "$631,758"],
        ["%TG New Sale", "33.31%"],
        ["Sales Count", "43"],
        ["USDOT Declined", "15"],
        ["Declination Ratio", "15.96%"],
        ["Quote to Sold Ratio", "36.13%"],
      ],
      note: "Best month — highest total premium, New Sale & Q2S ratio"
    },
    {
      name: "November · Month 11", bg: C.orangeDark,
      kpis: [
        ["New Sale (Pure Prem.)", "$546,134"],
        ["%TG New Sale", "28.79%"],
        ["Sales Count", "44"],
        ["USDOT Declined", "13"],
        ["Declination Ratio", "20.31%"],
        ["Quote to Sold Ratio", "42.72%"],
      ],
      note: "Mid-quarter improvement — highest Q2S ratio (42.72%)"
    },
    {
      name: "December · Month 12", bg: C.darkGray,
      kpis: [
        ["New Sale (Pure Prem.)", "$718,867"],
        ["%TG New Sale", "37.90%"],
        ["Sales Count", "53"],
        ["USDOT Declined", "21"],
        ["Declination Ratio", "24.71%"],
        ["Quote to Sold Ratio", "41.73%"],
      ],
      note: "Strongest finish — highest New Sale ($718K) and sales count (53)"
    },
  ];

  const colW = 3.1;
  months.forEach((m, i) => {
    const x = 0.18 + i * (colW + 0.13);
    // Month header
    slide.addShape(pres.shapes.RECTANGLE, { x, y: 0.75, w: colW, h: 0.42, fill: { color: m.bg }, line: { color: m.bg } });
    slide.addText(m.name, { x, y: 0.78, w: colW, h: 0.36, fontSize: 12, bold: true, color: C.white, align: "center", fontFace: "Arial", margin: 0 });

    // Card
    slide.addShape(pres.shapes.RECTANGLE, { x, y: 1.17, w: colW, h: 3.5, fill: { color: C.cardBg }, line: { color: C.orangeDark, width: 1 } });

    m.kpis.forEach((kpi, j) => {
      const ky = 1.22 + j * 0.52;
      const isFirst = j === 0;
      slide.addText(kpi[0], { x: x + 0.12, y: ky, w: colW - 0.24, h: 0.22, fontSize: isFirst ? 9 : 8.5, color: C.midGray, fontFace: "Calibri", margin: 0 });
      slide.addText(kpi[1], { x: x + 0.12, y: ky + 0.22, w: colW - 0.24, h: 0.28, fontSize: isFirst ? 18 : 14, bold: true, color: isFirst ? C.orange : C.white, fontFace: "Arial Black", margin: 0 });
    });

    // Note
    slide.addShape(pres.shapes.RECTANGLE, { x, y: 4.67, w: colW, h: 0.55, fill: { color: "1E1E1E" }, line: { color: C.orange, width: 1 } });
    slide.addText(m.note, { x: x + 0.1, y: 4.68, w: colW - 0.2, h: 0.53, fontSize: 8.5, color: C.midGray, fontFace: "Calibri", italic: true, margin: 0 });
  });

  // Footer summary
  slide.addShape(pres.shapes.RECTANGLE, { x: 0.18, y: 5.28, w: 9.64, h: 0.27, fill: { color: C.orangeDark }, line: { color: C.orangeDark } });
  slide.addText("Q4 Total: $631K + $546K + $718K = $1,896,759 New Sale · December strongest month · Q2S improved through quarter (36% → 42% → 41%)", {
    x: 0.3, y: 5.29, w: 9.4, h: 0.24, fontSize: 8.5, color: C.white, fontFace: "Calibri", align: "center", margin: 0
  });
}

// ============================================================
// SLIDE 4 — Sales Leaders
// ============================================================
{
  const slide = pres.addSlide();
  slide.background = { color: C.bg };

  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.65, fill: { color: C.orange }, line: { color: C.orange } });
  slide.addText("Sales Leaders — Full Q4 Performance Q4 2025 · 5 Leaders", {
    x: 0.3, y: 0.05, w: 9.4, h: 0.55, fontSize: 16, bold: true, color: C.white, fontFace: "Arial", margin: 0
  });

  // Office totals
  const totals = [
    { label: "Goal (Measure)", val: "$6,450,000" },
    { label: "Combined Premium", val: "$3,944,774" },
    { label: "Producer Fees", val: "$191,349" },
    { label: "New Sale Total", val: "$1,896,759" },
    { label: "% Achievement", val: "~61.2%" },
  ];
  const tw = 1.85;
  totals.forEach((t, i) => {
    const x = 0.18 + i * (tw + 0.05);
    slide.addShape(pres.shapes.RECTANGLE, { x, y: 0.72, w: tw, h: 0.75, fill: { color: C.cardBg }, line: { color: C.orange, width: 1 } });
    slide.addText(t.val, { x, y: 0.74, w: tw, h: 0.38, fontSize: 14, bold: true, color: C.orange, align: "center", fontFace: "Arial Black", margin: 0 });
    slide.addText(t.label, { x, y: 1.1, w: tw, h: 0.3, fontSize: 8, color: C.midGray, align: "center", fontFace: "Calibri", margin: 0 });
  });

  // Table — ordered by New Sale DESC
  const headers = ["Rank", "Sales Leader", "USDOT", "Sales Ct.", "New Sale", "Producer Fees", "Combined", "%TG", "Q2S", "Declined", "Dec.Ratio"];
  const colWs = [0.38, 1.5, 0.55, 0.55, 0.85, 0.85, 0.9, 0.55, 0.55, 0.55, 0.65];
  const leaders = [
    ["#1", "Mario Andres Ledesma", "82", "52", "$668,078", "$81,256", "$1,480,638.75", "33.23%", "39.68%", "14", "21.54%"],
    ["#2", "Laura Serna", "54", "30", "$568,961", "$42,100", "$989,961.41", "30.00%", "43.48%", "3", "10.34%"],
    ["#3", "Jhon Dairon Lopez", "74", "31", "$358,314", "$39,093", "$749,243.63", "18.89%", "41.89%", "17", "26.98%"],
    ["#4", "Yessica Rendon", "50", "22", "$253,317", "$20,900", "$437,317.20", "12.95%", "50.00%", "10", "21.74%"],
    ["#5", "Yefrik Alvarez", "89", "6", "$94,124", "$11,000", "$184,124.20", "5.99%", "15.15%", "5", "12.82%"],
  ];

  const tableStartY = 1.57;
  const rowH = 0.34;
  let startX = 0.12;

  // Header
  slide.addShape(pres.shapes.RECTANGLE, { x: startX, y: tableStartY, w: 9.76, h: rowH, fill: { color: C.orange }, line: { color: C.orange } });
  let cx = startX;
  headers.forEach((h, ci) => {
    slide.addText(h, { x: cx, y: tableStartY + 0.04, w: colWs[ci], h: rowH - 0.08, fontSize: 9, bold: true, color: C.white, align: "center", fontFace: "Arial", margin: 0 });
    cx += colWs[ci];
  });

  // Rows
  leaders.forEach((row, ri) => {
    const ry = tableStartY + rowH + ri * rowH;
    const rowBg = ri % 2 === 0 ? C.tableRowBase : C.tableRowAlt;
    slide.addShape(pres.shapes.RECTANGLE, { x: startX, y: ry, w: 9.76, h: rowH, fill: { color: rowBg }, line: { color: "333333" } });
    cx = startX;
    row.forEach((cell, ci) => {
      const isNewSale = ci === 4;
      const isFees = ci === 5;
      const isCombined = ci === 6;
      const isQ2S = ci === 8;
      const color = isNewSale || isFees || isCombined ? C.orange : (isQ2S && parseFloat(row[8]) > 45 ? C.green : C.white);
      slide.addText(cell, { x: cx, y: ry + 0.05, w: colWs[ci], h: rowH - 0.1, fontSize: 8, bold: isNewSale || isFees || isCombined, color, align: "center", fontFace: "Calibri", margin: 0 });
      cx += colWs[ci];
    });
  });

  // Insights
  slide.addShape(pres.shapes.RECTANGLE, { x: 0.12, y: tableStartY + rowH * 7 + 0.08, w: 9.76, h: 0.52, fill: { color: C.cardBg }, line: { color: C.orange, width: 1 } });
  slide.addText([
    { text: "Key Insights: ", options: { bold: true, color: C.orange, fontSize: 9 } },
    { text: "Yessica Rendon leads Q2S at 50.00%. ", options: { color: C.white, fontSize: 9 } },
    { text: "Mario Andres Ledesma leads New Sale ($668K) and fees ($81K). ", options: { color: C.white, fontSize: 9 } },
    { text: "Mario Andres most sales (52). ", options: { color: C.white, fontSize: 9 } },
    { text: "Order determined by New Sale volume.", options: { color: C.midGray, fontSize: 9 } },
  ], { x: 0.25, y: tableStartY + rowH * 6 + 0.14, w: 9.5, h: 0.4, fontFace: "Calibri", margin: 0 });
}

// ============================================================
// SLIDE 5 — Top Producers
// ============================================================
{
  const slide = pres.addSlide();
  slide.background = { color: C.bg };

  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.5, fill: { color: C.orange }, line: { color: C.orange } });
  slide.addText("Top Producers — Top 5 by New Sale Q4 2025 · 19 Producers", {
    x: 0.3, y: 0.06, w: 9.4, h: 0.4, fontSize: 14, bold: true, color: C.white, fontFace: "Arial", margin: 0
  });

  // Producer data
  const producers = [
    {
      rank: "★ #1", name: "Ronald Alexander\nAzuaje Perez", newSale: "$265,132", usdot: 11, sales: 5,
      tg: "13.98%", decl: 1, q2s: "35.71%", decRatio: "14.29%",
      fees: "$17,250", combined: "$377,632",
      note: "Top clients: J & I Multiservices ($211K), ATD LOGISTICS ($62K total)"
    },
    {
      rank: "★ #2", name: "Jose David\nBedoya", newSale: "$250,768", usdot: 20, sales: 20,
      tg: "13.22%", decl: 3, q2s: "33.90%", decRatio: "21.43%",
      fees: "$42,401", combined: "$674,778",
      note: "Consistent high volume across multiple carriers"
    },
    {
      rank: "★ #3", name: "Melisa Negrete\nCuello", newSale: "$195,885", usdot: 7, sales: 8,
      tg: "10.33%", decl: 1, q2s: "114.29%", decRatio: "50.00%",
      fees: "$12,600", combined: "$321,885",
      note: "Outstanding Q2S ratio (114.29%) - exceptional conversion"
    },
    {
      rank: "#4", name: "Juan Fernando\nVasquez Usma", newSale: "$126,179", usdot: 11, sales: 8,
      tg: "6.65%", decl: 2, q2s: "53.33%", decRatio: "25.00%",
      fees: "$6,000", combined: "$186,179",
      note: "Strong Q2S performance"
    },
    {
      rank: "★ #5", name: "Carlos Andres\nHernandez Rivera", newSale: "$103,496", usdot: 8, sales: 7,
      tg: "5.46%", decl: 1, q2s: "50.00%", decRatio: "16.67%",
      fees: "$9,700", combined: "$200,496",
      note: "Solid performance with excellent Q2S"
    },
  ];

  // Layout: 5 columns
  const colWidth = 1.9;
  const startX = 0.12;
  const gap = 0.07;

  producers.forEach((p, i) => {
    const x = startX + i * (colWidth + gap);

    // Rank badge
    const rankColor = i < 3 ? C.orange : C.orangeDark;
    slide.addShape(pres.shapes.RECTANGLE, { x, y: 0.55, w: colWidth, h: 0.3, fill: { color: rankColor }, line: { color: rankColor } });
    slide.addText(p.rank, { x, y: 0.57, w: colWidth, h: 0.26, fontSize: 11, bold: true, color: C.white, align: "center", fontFace: "Arial Black", margin: 0 });

    // Name
    slide.addShape(pres.shapes.RECTANGLE, { x, y: 0.85, w: colWidth, h: 0.5, fill: { color: C.cardBg }, line: { color: C.orangeDark, width: 1 } });
    slide.addText(p.name, { x, y: 0.87, w: colWidth, h: 0.46, fontSize: 9, bold: true, color: C.white, align: "center", fontFace: "Arial", margin: 0 });

    // New Sale big
    slide.addText(p.newSale, { x, y: 1.37, w: colWidth, h: 0.35, fontSize: 14, bold: true, color: C.orange, align: "center", fontFace: "Arial Black", margin: 0 });
    slide.addText("New Sale (Pure Prem.)", { x, y: 1.68, w: colWidth, h: 0.2, fontSize: 7, color: C.midGray, align: "center", fontFace: "Calibri", margin: 0 });

    // Stats mini
    const stats = [
      `USDOT: ${p.usdot}`, `Sales: ${p.sales}`, `%TG: ${p.tg}`,
      `Decl: ${p.decl}`, `Q2S: ${p.q2s}`, `Dec%: ${p.decRatio}`,
      `Fees: ${p.fees}`, `Comb: ${p.combined}`
    ];
    stats.forEach((s, j) => {
      const row = Math.floor(j / 2);
      const col = j % 2;
      const sx = x + col * (colWidth / 2);
      const sy = 1.9 + row * 0.22;
      slide.addText(s, { x: sx, y: sy, w: colWidth / 2, h: 0.2, fontSize: 7.5, color: C.midGray, fontFace: "Calibri", align: "center", margin: 0 });
    });

    // Note section
    const noteY = 2.8;
    slide.addShape(pres.shapes.RECTANGLE, { x, y: noteY, w: colWidth, h: 2.7, fill: { color: C.cardBg }, line: { color: C.orange, width: 1 } });
    slide.addText(p.note, { x: x + 0.1, y: noteY + 0.1, w: colWidth - 0.2, h: 2.5, fontSize: 7.5, color: C.midGray, italic: true, fontFace: "Calibri", margin: 0 });
  });
}

// ============================================================
// SLIDE 6 — Zero Sale & Bottom 5
// ============================================================
{
  const slide = pres.addSlide();
  slide.background = { color: C.bg };

  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.65, fill: { color: C.red }, line: { color: C.red } });
  slide.addText("Producers — Zero New Sale Alert & Bottom 5 · Q4 2025", {
    x: 0.3, y: 0.05, w: 9.4, h: 0.55, fontSize: 16, bold: true, color: C.white, fontFace: "Arial", margin: 0
  });

  // Zero sale section
  slide.addText("ZERO NEW SALE — Immediate Follow-Up Required", {
    x: 0.2, y: 0.75, w: 9.6, h: 0.35, fontSize: 12, bold: true, color: C.red, fontFace: "Arial", margin: 0
  });

  const zeroProducers = [
    {
      name: "Felipe Mejia\nOlarte", usdot: 2, qPrem: "$25,915", qCnt: 1,
      note: "Has pipeline activity ($25K quotes) but zero conversions"
    },
    {
      name: "Daniela Alejandra\nSalas Estrada", usdot: 9, qPrem: "$135,709", qCnt: 5,
      note: "Significant quote volume ($135K) with no sales - requires immediate intervention"
    },
  ];

  zeroProducers.forEach((p, i) => {
    const x = 0.2 + i * 4.85;
    slide.addShape(pres.shapes.RECTANGLE, { x, y: 1.14, w: 4.6, h: 1.4, fill: { color: "2A0A0A" }, line: { color: C.red, width: 2 } });
    slide.addText("$0", { x, y: 1.18, w: 4.6, h: 0.55, fontSize: 36, bold: true, color: C.red, align: "center", fontFace: "Arial Black", margin: 0 });
    slide.addText(p.name, { x, y: 1.73, w: 4.6, h: 0.35, fontSize: 10, bold: true, color: C.white, align: "center", fontFace: "Arial", margin: 0 });
    slide.addText(`USDOT: ${p.usdot}  |  Quotes Prem: ${p.qPrem}  |  Quotes Cnt: ${p.qCnt}`, {
      x, y: 2.08, w: 4.6, h: 0.22, fontSize: 8, color: C.midGray, align: "center", fontFace: "Calibri", margin: 0
    });
    slide.addText(p.note, { x: x + 0.1, y: 2.32, w: 4.4, h: 0.18, fontSize: 8, color: C.red, italic: true, fontFace: "Calibri", margin: 0 });
  });

  // Action items
  slide.addShape(pres.shapes.RECTANGLE, { x: 0.2, y: 2.62, w: 9.6, h: 0.28, fill: { color: C.orangeDark }, line: { color: C.orangeDark } });
  slide.addText("Action Required for Both: · Identify blockers in pipeline · Align with MGA appetite · Set hard closing deadline for Q1 2026", {
    x: 0.3, y: 2.64, w: 9.4, h: 0.24, fontSize: 8.5, color: C.white, fontFace: "Calibri", margin: 0
  });

  // Bottom 5
  slide.addText("BOTTOM 5 PRODUCERS — Highest Improvement Opportunity", {
    x: 0.2, y: 3.0, w: 9.6, h: 0.3, fontSize: 12, bold: true, color: C.amber, fontFace: "Arial", margin: 0
  });

  const bottom5 = [
    { rank: "#1", name: "Tatiana Velez\nBedoya", newSale: "$6,644", usdot: 4, sales: 1, tg: "0.35%", fees: "$2,100", combined: "$27,644" },
    { rank: "#2", name: "Kimberlin Roixeli\nDiaz Marquez", newSale: "$8,299", usdot: 2, sales: 2, tg: "0.44%", fees: "$1,980", combined: "$28,099" },
    { rank: "#3", name: "Nathali Carolina\nVegas Carballo", newSale: "$11,222", usdot: 6, sales: 1, tg: "0.59%", fees: "$1,000", combined: "$21,222" },
    { rank: "#4", name: "Anderson Carvajal\nOsorno", newSale: "$14,829", usdot: 16, sales: 3, tg: "0.78%", fees: "$3,300", combined: "$47,829" },
    { rank: "#5", name: "Kevin Alfredo\nMeza Noriega", newSale: "$15,041", usdot: 4, sales: 2, tg: "0.79%", fees: "$2,350", combined: "$38,541" },
  ];

  const bw = 1.85;
  bottom5.forEach((p, i) => {
    const x = 0.2 + i * (bw + 0.1);
    slide.addShape(pres.shapes.RECTANGLE, { x, y: 3.35, w: bw, h: 1.95, fill: { color: C.cardBg }, line: { color: C.amber, width: 1 } });
    slide.addShape(pres.shapes.RECTANGLE, { x, y: 3.35, w: bw, h: 0.28, fill: { color: C.amber }, line: { color: C.amber } });
    slide.addText(p.rank, { x, y: 3.37, w: bw, h: 0.24, fontSize: 10, bold: true, color: C.textDark, align: "center", fontFace: "Arial Black", margin: 0 });
    slide.addText(p.name, { x, y: 3.66, w: bw, h: 0.4, fontSize: 9, bold: true, color: C.white, align: "center", fontFace: "Arial", margin: 0 });
    slide.addText(p.newSale, { x, y: 4.1, w: bw, h: 0.38, fontSize: 18, bold: true, color: C.amber, align: "center", fontFace: "Arial Black", margin: 0 });
    slide.addText(`USDOT: ${p.usdot}  Sales: ${p.sales}  %TG: ${p.tg}`, {
      x, y: 4.52, w: bw, h: 0.18, fontSize: 7, color: C.midGray, align: "center", fontFace: "Calibri", margin: 0
    });
    slide.addText(`Fees: ${p.fees}  Combined: ${p.combined}`, {
      x, y: 4.70, w: bw, h: 0.18, fontSize: 7, color: C.orange, align: "center", fontFace: "Calibri", margin: 0
    });
    slide.addText("Coaching + pipeline alignment needed", {
      x, y: 4.90, w: bw, h: 0.35, fontSize: 7, color: "888888", align: "center", italic: true, fontFace: "Calibri", margin: 0
    });
  });

  slide.addShape(pres.shapes.RECTANGLE, { x: 0.2, y: 5.35, w: 9.6, h: 0.22, fill: { color: C.cardBg }, line: { color: C.amber } });
  slide.addText("These producers represent the highest improvement opportunity · Coaching + pipeline alignment can significantly impact Q1 2026 results", {
    x: 0.3, y: 5.37, w: 9.4, h: 0.18, fontSize: 8, color: C.midGray, align: "center", fontFace: "Calibri", margin: 0
  });
}

// ============================================================
// SLIDE 7 — Premium per Unit Type
// ============================================================
{
  const slide = pres.addSlide();
  slide.background = { color: C.bg };

  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.65, fill: { color: C.orange }, line: { color: C.orange } });
  slide.addText("Premium per Unit Type — Full Drill-Down Q4 2025", {
    x: 0.3, y: 0.05, w: 9.4, h: 0.55, fontSize: 16, bold: true, color: C.white, fontFace: "Arial", margin: 0
  });

  // Top 3 units
  const units = [
    { name: "Truck Tractor", usdot: 195, sales: 94, newSale: "$1,482,549", tg: "78.16%", decl: 34, qPrem: "$4,336,774", qCnt: 228, q2s: "41.23%", decR: "19.43%", note: "Dominant — 78.16% of New Sale" },
    { name: "Box Truck",     usdot: 49,  sales: 17, newSale: "$170,917",   tg: "9.01%",  decl: 7,  qPrem: "$549,713",   qCnt: 47,  q2s: "36.17%", decR: "19.44%", note: "Strong #2 — 9.01% of New Sale" },
    { name: "Pickup",        usdot: 40,  sales: 16, newSale: "$116,345",   tg: "6.13%",  decl: 5,  qPrem: "$652,886",   qCnt: 45,  q2s: "35.56%", decR: "20.83%", note: "Moderate volume — 6.13% of New Sale" },
  ];

  const unitW = 3.1;
  units.forEach((u, i) => {
    const x = 0.18 + i * (unitW + 0.15);
    slide.addShape(pres.shapes.RECTANGLE, { x, y: 0.72, w: unitW, h: 3.9, fill: { color: C.cardBg }, line: { color: i === 0 ? C.orange : C.orangeDark, width: i === 0 ? 2 : 1 } });

    slide.addShape(pres.shapes.RECTANGLE, { x, y: 0.72, w: unitW, h: 0.38, fill: { color: i === 0 ? C.orange : C.orangeDark }, line: { color: C.orange } });
    slide.addText(u.name, { x, y: 0.74, w: unitW, h: 0.34, fontSize: 12, bold: true, color: C.white, align: "center", fontFace: "Arial", margin: 0 });

    slide.addText(u.newSale, { x, y: 1.13, w: unitW, h: 0.45, fontSize: 17, bold: true, color: C.orange, align: "center", fontFace: "Arial Black", margin: 0 });
    slide.addText("%TG: " + u.tg, { x, y: 1.58, w: unitW, h: 0.25, fontSize: 11, bold: true, color: C.white, align: "center", fontFace: "Calibri", margin: 0 });

    const details = [
      ["USDOT Count",      String(u.usdot)],
      ["Sales Count",      String(u.sales)],
      ["USDOT Declined",   String(u.decl)],
      ["Quotes Premium",   u.qPrem],
      ["Quotes Count",     String(u.qCnt)],
      ["Q2S Ratio",        u.q2s],
      ["Declination Ratio",u.decR],
    ];
    details.forEach((d, j) => {
      const dy = 1.88 + j * 0.3;
      slide.addText(d[0], { x: x + 0.1, y: dy, w: unitW * 0.6, h: 0.25, fontSize: 8.5, color: C.midGray, fontFace: "Calibri", margin: 0 });
      slide.addText(d[1], { x: x + unitW * 0.6, y: dy, w: unitW * 0.38, h: 0.25, fontSize: 8.5, bold: true, color: C.white, align: "right", fontFace: "Calibri", margin: 0 });
    });

    slide.addShape(pres.shapes.RECTANGLE, { x, y: 4.62, w: unitW, h: 0.3, fill: { color: "1A1A1A" }, line: { color: C.orange } });
    slide.addText(u.note, { x: x + 0.05, y: 4.63, w: unitW - 0.1, h: 0.28, fontSize: 7.5, color: C.midGray, italic: true, fontFace: "Calibri", align: "center", margin: 0 });
  });

  // Additional unit types note
  slide.addText("Additional unit types with smaller volumes: Cargo Van, Dump Truck, Flatbed, and others contributing remaining 6.7%", {
    x: 0.18, y: 5.05, w: 9.64, h: 0.3, fontSize: 8, color: "777777", italic: true, fontFace: "Calibri", margin: 0
  });
}

// ============================================================
// SLIDE 8 — Carriers
// ============================================================
{
  const slide = pres.addSlide();
  slide.background = { color: C.bg };

  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.65, fill: { color: C.orange }, line: { color: C.orange } });
  slide.addText("Premium per Carrier — Top 5 Full Detail Q4 2025 · 29 Carriers", {
    x: 0.3, y: 0.05, w: 9.4, h: 0.55, fontSize: 16, bold: true, color: C.white, fontFace: "Arial", margin: 0
  });

  // Spotlight — Pegaso #1
  slide.addShape(pres.shapes.RECTANGLE, { x: 0.18, y: 0.72, w: 4.7, h: 2.1, fill: { color: C.cardBg }, line: { color: C.orange, width: 2 }, shadow: makeShadow() });
  slide.addText("★ Carrier Spotlight — Pegaso Risk Retention Group", {
    x: 0.28, y: 0.78, w: 4.5, h: 0.3, fontSize: 11, bold: true, color: C.orange, fontFace: "Arial", margin: 0
  });
  const pegStats = [
    ["79.31%", "Q2S Ratio"], ["18.18%", "Declination"], ["23", "Sales Count"], ["$403,247", "New Sale"]
  ];
  pegStats.forEach((s, i) => {
    const px = 0.28 + i * 1.15;
    slide.addText(s[0], { x: px, y: 1.12, w: 1.1, h: 0.42, fontSize: 15, bold: true, color: C.orange, align: "center", fontFace: "Arial Black", margin: 0 });
    slide.addText(s[1], { x: px, y: 1.54, w: 1.1, h: 0.25, fontSize: 7.5, color: C.midGray, align: "center", fontFace: "Calibri", margin: 0 });
  });
  slide.addText("Pegaso: highest Q2S (79.31%) and lowest declination (18.18%) of all carriers — #1 placement priority.", {
    x: 0.28, y: 1.83, w: 4.5, h: 0.35, fontSize: 8, color: C.midGray, italic: true, fontFace: "Calibri", margin: 0
  });

  // Bar chart — Top 5 by New Sale DESC
  const carriers = [
    { name: "Pegaso Risk Retention Group", val: 403247 },
    { name: "Motor Transport Mutual RRG", val: 379882 },
    { name: "Accredited Specialty Insurance", val: 217908 },
    { name: "Progressive County Mutual", val: 131302 },
    { name: "Geico County Mutual Ins Co.", val: 124962 },
  ];
  const maxValC = 403247;
  const barMaxWC = 4.0;
  slide.addText("Carrier Ranking by New Sale (USD)", {
    x: 5.08, y: 0.72, w: 4.7, h: 0.3, fontSize: 10, bold: true, color: C.orange, fontFace: "Arial", margin: 0
  });
  carriers.forEach((c, i) => {
    const by = 1.1 + i * 0.56;
    const bw = Math.max((c.val / maxValC) * barMaxWC, 0.15);
    slide.addText(c.name, { x: 5.08, y: by, w: 4.7, h: 0.2, fontSize: 8, color: C.midGray, fontFace: "Calibri", margin: 0 });
    slide.addShape(pres.shapes.RECTANGLE, { x: 5.08, y: by + 0.22, w: bw, h: 0.22, fill: { color: i === 0 ? C.orange : C.orangeDark }, line: { color: C.orangeDark } });
    slide.addText("$" + (c.val / 1000).toFixed(0) + "K", { x: 5.08 + bw + 0.05, y: by + 0.22, w: 1.0, h: 0.22, fontSize: 8, bold: true, color: C.white, fontFace: "Calibri", margin: 0 });
  });

  // Table
  const tHeaders8 = ["Carrier", "USDOT", "Quotes Ct.", "Quotes Prem.", "New Sale", "%TG"];
  const tColWs8 = [2.7, 0.7, 0.8, 1.1, 1.05, 0.75];
  const tData8 = [
    ["Pegaso Risk Retention Group",      "48", "29", "$922,193",   "$403,247", "21.76%"],
    ["Motor Transport Mutual RRG",       "13", "14", "$700,706",   "$379,882", "20.50%"],
    ["Accredited Specialty Insurance",   "14", "20", "$872,989",   "$217,908", "11.76%"],
    ["Progressive County Mutual",        "14", "14", "$187,678",   "$131,302",  "7.09%"],
    ["Geico County Mutual Ins Co.",      "10", "13", "$231,079",   "$124,962",  "6.74%"],
  ];
  const tStartX8 = 0.18, tStartY8 = 2.95, rowH8 = 0.31;
  const tTotalW8 = tColWs8.reduce((a, b) => a + b, 0);

  slide.addShape(pres.shapes.RECTANGLE, { x: tStartX8, y: tStartY8, w: tTotalW8, h: rowH8, fill: { color: C.orange }, line: { color: C.orange } });
  let cx8 = tStartX8;
  tHeaders8.forEach((h, ci) => {
    slide.addText(h, { x: cx8, y: tStartY8 + 0.05, w: tColWs8[ci], h: rowH8 - 0.1, fontSize: 8.5, bold: true, color: C.white, align: "center", fontFace: "Arial", margin: 0 });
    cx8 += tColWs8[ci];
  });
  tData8.forEach((row, ri) => {
    const ry = tStartY8 + rowH8 + ri * rowH8;
    slide.addShape(pres.shapes.RECTANGLE, { x: tStartX8, y: ry, w: tTotalW8, h: rowH8, fill: { color: ri % 2 === 0 ? C.tableRowBase : C.tableRowAlt }, line: { color: "333333" } });
    cx8 = tStartX8;
    row.forEach((cell, ci) => {
      const isNS = ci === 4;
      slide.addText(cell, { x: cx8 + 0.03, y: ry + 0.05, w: tColWs8[ci] - 0.06, h: rowH8 - 0.1, fontSize: 8.5, bold: isNS, color: isNS ? C.orange : C.white, align: ci === 0 ? "left" : "center", fontFace: "Calibri", margin: 0 });
      cx8 += tColWs8[ci];
    });
  });
  slide.addText("Additional carriers (24 more): United Financial Casualty · Drive New Jersey · Underwriters At Lloyd's · Trailblazers + more", {
    x: tStartX8, y: tStartY8 + rowH8 * 6 + 0.05, w: tTotalW8, h: 0.25, fontSize: 7.5, color: "777777", italic: true, fontFace: "Calibri", margin: 0
  });
}

// ============================================================
// SLIDE 9 — MGAs
// ============================================================
{
  const slide = pres.addSlide();
  slide.background = { color: C.bg };

  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.65, fill: { color: C.orange }, line: { color: C.orange } });
  slide.addText("Premium per MGA — Top 5 Full Detail Q4 2025 · 29 MGAs", {
    x: 0.3, y: 0.05, w: 9.4, h: 0.55, fontSize: 16, bold: true, color: C.white, fontFace: "Arial", margin: 0
  });

  // Spotlight — Nexus
  slide.addShape(pres.shapes.RECTANGLE, { x: 0.18, y: 0.72, w: 4.7, h: 2.1, fill: { color: C.cardBg }, line: { color: C.orange, width: 2 } });
  slide.addText("★ MGA Spotlight — Nexus Risk Management", {
    x: 0.28, y: 0.78, w: 4.5, h: 0.3, fontSize: 11, bold: true, color: C.orange, fontFace: "Arial", margin: 0
  });
  const nexStats = [
    ["56.79%", "Q2S Ratio"], ["10.45%", "Declination"], ["46", "Sales Count"], ["81", "Quotes Count"]
  ];
  nexStats.forEach((s, i) => {
    const px = 0.28 + i * 1.15;
    slide.addText(s[0], { x: px, y: 1.12, w: 1.1, h: 0.42, fontSize: 15, bold: true, color: C.orange, align: "center", fontFace: "Arial Black", margin: 0 });
    slide.addText(s[1], { x: px, y: 1.54, w: 1.1, h: 0.25, fontSize: 7.5, color: C.midGray, align: "center", fontFace: "Calibri", margin: 0 });
  });
  slide.addText("Nexus leads with 56.79% Q2S, highest Quotes count (81) and low declination (10.45%). Primary MGA partner.", {
    x: 0.28, y: 1.83, w: 4.5, h: 0.35, fontSize: 8, color: C.midGray, italic: true, fontFace: "Calibri", margin: 0
  });

  // Bar chart — Top 5 by New Sale DESC
  const mgas = [
    { name: "Nexus Risk Management", val: 445599 },
    { name: "Motor Transport Mutual RRG", val: 404614 },
    { name: "Cover Whale", val: 335525 },
    { name: "Progressive", val: 293256 },
    { name: "GEICO", val: 159808 },
  ];
  const maxValM = 445599;
  const barMaxWM = 4.0;
  slide.addText("MGA Ranking by New Sale (USD)", {
    x: 5.08, y: 0.72, w: 4.7, h: 0.3, fontSize: 10, bold: true, color: C.orange, fontFace: "Arial", margin: 0
  });
  mgas.forEach((m, i) => {
    const by = 1.1 + i * 0.56;
    const bw = Math.max((m.val / maxValM) * barMaxWM, 0.1);
    slide.addText(m.name, { x: 5.08, y: by, w: 4.7, h: 0.2, fontSize: 8, color: C.midGray, fontFace: "Calibri", margin: 0 });
    slide.addShape(pres.shapes.RECTANGLE, { x: 5.08, y: by + 0.22, w: bw, h: 0.22, fill: { color: i === 0 ? C.orange : C.orangeDark }, line: { color: C.orangeDark } });
    slide.addText("$" + (m.val / 1000).toFixed(0) + "K", { x: 5.08 + bw + 0.05, y: by + 0.22, w: 0.9, h: 0.22, fontSize: 8, bold: true, color: C.white, fontFace: "Calibri", margin: 0 });
  });

  // Table
  const tHeaders9 = ["MGA", "USDOT", "Quotes Ct.", "Quotes Prem.", "New Sale", "%TG"];
  const tColWs9 = [2.5, 0.7, 0.8, 1.15, 1.05, 0.72];
  const tData9 = [
    ["Nexus Risk Management",         "105", "81", "$1,076,052", "$445,599", "23.49%"],
    ["Motor Transport Mutual RRG",     "28", "17",   "$978,230", "$404,614", "21.33%"],
    ["Cover Whale",                    "18", "49", "$1,041,723", "$335,525", "12.42%"],
    ["Progressive",                    "46", "43", "$1,690,597", "$293,256", "15.46%"],
    ["GEICO",                          "24", "28",   "$513,937", "$159,808",  "8.95%"],
  ];
  const tStartX9 = 0.18, tStartY9 = 2.95, rowH9 = 0.31;
  const tTotalW9 = tColWs9.reduce((a, b) => a + b, 0);

  slide.addShape(pres.shapes.RECTANGLE, { x: tStartX9, y: tStartY9, w: tTotalW9, h: rowH9, fill: { color: C.orange }, line: { color: C.orange } });
  let cx9 = tStartX9;
  tHeaders9.forEach((h, ci) => {
    slide.addText(h, { x: cx9, y: tStartY9 + 0.05, w: tColWs9[ci], h: rowH9 - 0.1, fontSize: 8.5, bold: true, color: C.white, align: "center", fontFace: "Arial", margin: 0 });
    cx9 += tColWs9[ci];
  });
  tData9.forEach((row, ri) => {
    const ry = tStartY9 + rowH9 + ri * rowH9;
    slide.addShape(pres.shapes.RECTANGLE, { x: tStartX9, y: ry, w: tTotalW9, h: rowH9, fill: { color: ri % 2 === 0 ? C.tableRowBase : C.tableRowAlt }, line: { color: "333333" } });
    cx9 = tStartX9;
    row.forEach((cell, ci) => {
      const isNS = ci === 4;
      slide.addText(cell, { x: cx9 + 0.03, y: ry + 0.05, w: tColWs9[ci] - 0.06, h: rowH9 - 0.1, fontSize: 8.5, bold: isNS, color: isNS ? C.orange : C.white, align: ci === 0 ? "left" : "center", fontFace: "Calibri", margin: 0 });
      cx9 += tColWs9[ci];
    });
  });
  slide.addText("Additional MGAs (24 more): County Hall RRG · First Light Program Managers · AmWINS · Berkshire Hathaway + more", {
    x: tStartX9, y: tStartY9 + rowH9 * 6 + 0.05, w: tTotalW9, h: 0.25, fontSize: 7.5, color: "777777", italic: true, fontFace: "Calibri", margin: 0
  });
}

// ============================================================
// SLIDE 10 — Business Type
// ============================================================
{
  const slide = pres.addSlide();
  slide.background = { color: C.bg };

  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.65, fill: { color: C.orange }, line: { color: C.orange } });
  slide.addText("Business Type — Full Detail Q4 2025", {
    x: 0.3, y: 0.05, w: 9.4, h: 0.55, fontSize: 16, bold: true, color: C.white, fontFace: "Arial", margin: 0
  });

  // Top 5 by New Sale DESC
  const bizTypes = [
    { name: "General Freight",    newSale: "$866,328", tg: "45.67%", usdot: 187, decl: 27, q2s: "40.96%", decR: "19.01%" },
    { name: "Building Materials", newSale: "$464,830", tg: "24.51%", usdot: 36,  decl: 4,  q2s: "53.85%", decR: "17.39%" },
    { name: "Dirt Sand & Gravel", newSale: "$218,628", tg: "11.53%", usdot: 30,  decl: 4,  q2s: "63.64%", decR: "28.57%" },
    { name: "Auto Hauler",        newSale: "$65,526",  tg:  "3.45%", usdot: 25,  decl: 3,  q2s: "24.00%", decR: "20.00%" },
    { name: "Passenger",          newSale: "$60,745",  tg:  "3.20%", usdot: 5,   decl: 1,  q2s: "75.00%", decR: "33.33%" },
  ];

  // LEFT TABLE
  const tHeadersBT = ["Business Type", "New Sale", "%TG", "USDOT", "Dec.", "Q2S", "Dec.%"];
  const tColWsBT = [1.75, 1.05, 0.65, 0.6, 0.5, 0.65, 0.65];
  const tStartXBT = 0.15, tStartYBT = 0.75, rowHBT = 0.38;
  const tTotalWBT = tColWsBT.reduce((a, b) => a + b, 0);

  slide.addShape(pres.shapes.RECTANGLE, { x: tStartXBT, y: tStartYBT, w: tTotalWBT, h: rowHBT, fill: { color: C.orange }, line: { color: C.orange } });
  let cxBT = tStartXBT;
  tHeadersBT.forEach((h, ci) => {
    slide.addText(h, { x: cxBT, y: tStartYBT + 0.06, w: tColWsBT[ci], h: rowHBT - 0.12, fontSize: 9, bold: true, color: C.white, align: "center", fontFace: "Arial", margin: 0 });
    cxBT += tColWsBT[ci];
  });
  bizTypes.forEach((b, ri) => {
    const ry = tStartYBT + rowHBT + ri * rowHBT;
    slide.addShape(pres.shapes.RECTANGLE, { x: tStartXBT, y: ry, w: tTotalWBT, h: rowHBT, fill: { color: ri % 2 === 0 ? C.tableRowBase : C.tableRowAlt }, line: { color: "333333" } });
    const vals = [b.name, b.newSale, b.tg, String(b.usdot), String(b.decl), b.q2s, b.decR];
    cxBT = tStartXBT;
    vals.forEach((cell, ci) => {
      const isNS = ci === 1;
      slide.addText(cell, { x: cxBT + 0.04, y: ry + 0.08, w: tColWsBT[ci] - 0.08, h: rowHBT - 0.16, fontSize: 9, bold: isNS, color: isNS ? C.orange : C.white, align: ci === 0 ? "left" : "center", fontFace: "Calibri", margin: 0 });
      cxBT += tColWsBT[ci];
    });
  });
  slide.addText("+ Additional types: Refrigerated Food · Household Goods · Haz Mat · Towing + more", {
    x: tStartXBT, y: tStartYBT + rowHBT * 6 + 0.08, w: tTotalWBT, h: 0.22, fontSize: 8, color: "777777", italic: true, fontFace: "Calibri", margin: 0
  });

  // RIGHT BAR CHART
  const chartX = 6.1, chartW = 3.7;
  slide.addText("New Sale by Business Type", {
    x: chartX, y: 0.72, w: chartW, h: 0.3, fontSize: 10, bold: true, color: C.orange, fontFace: "Arial", margin: 0
  });
  const maxNSBT = 866328;
  bizTypes.forEach((b, i) => {
    const bNS = parseFloat(b.newSale.replace(/[$,]/g, ""));
    const bw = Math.max((bNS / maxNSBT) * 2.5, 0.1);
    const by = 0.75 + 0.38 + i * 0.38;
    slide.addText(b.name, { x: chartX, y: by + 0.02, w: chartW - 1.2, h: 0.18, fontSize: 7.5, color: C.midGray, fontFace: "Calibri", margin: 0 });
    slide.addShape(pres.shapes.RECTANGLE, { x: chartX, y: by + 0.21, w: bw, h: 0.14, fill: { color: i === 0 ? C.orange : C.orangeDark }, line: { color: C.orangeDark } });
    slide.addText(b.newSale, { x: chartX + bw + 0.06, y: by + 0.18, w: 1.1, h: 0.2, fontSize: 7.5, bold: true, color: C.white, fontFace: "Calibri", margin: 0 });
  });
}

// ============================================================
// SLIDE 11 — Premium per Miles
// ============================================================
{
  const slide = pres.addSlide();
  slide.background = { color: C.bg };

  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.65, fill: { color: C.orange }, line: { color: C.orange } });
  slide.addText("Premium per Miles — Full Detail Q4 2025", {
    x: 0.3, y: 0.05, w: 9.4, h: 0.55, fontSize: 16, bold: true, color: C.white, fontFace: "Arial", margin: 0
  });

  slide.addText("All Mileage Bands — ordered by New Sale (DESC)", {
    x: 0.18, y: 0.72, w: 9.64, h: 0.28, fontSize: 10, bold: true, color: C.orange, fontFace: "Arial", margin: 0
  });

  // Miles data ordered by New Sale descending
  const milesData = [
    { miles: "500",      newSale: "$770,167", tg: "40.60%", usdot: 85,  declined: 11, qCnt: 118, qPrem: "$2,729,336", q2s: "34.75%", decR: "22.92%" },
    { miles: "Unlimited",newSale: "$561,412", tg: "29.60%", usdot: 151, declined: 24, qCnt: 134, qPrem: "$1,980,274", q2s: "44.03%", decR: "18.46%" },
    { miles: "100",      newSale: "$209,205", tg: "13.88%", usdot: 26,  declined: 5,  qCnt: 23,  qPrem: "$286,811",   q2s: "43.48%", decR: "27.78%" },
    { miles: "200",      newSale: "$165,610", tg: "10.98%", usdot: 34,  declined: 6,  qCnt: 40,  qPrem: "$482,813",   q2s: "35.00%", decR: "21.43%" },
    { miles: "50",       newSale: "$154,454", tg:  "8.14%", usdot: 17,  declined: 1,  qCnt: 17,  qPrem: "$255,369",   q2s: "70.59%", decR: "100.00%" },
    { miles: "300",      newSale: "$35,912",  tg:  "2.38%", usdot: 26,  declined: 1,  qCnt: 15,  qPrem: "$151,855",   q2s: "26.67%", decR: "6.67%" },
  ];

  const mTHeaders = ["Miles", "New Sale", "%TG", "USDOT", "Decl.", "Quotes Ct.", "Quotes Prem.", "Q2S", "Dec.%"];
  const mTColWs  = [0.7, 1.05, 0.62, 0.62, 0.55, 0.75, 1.2, 0.7, 0.62];
  const mTotalW  = mTColWs.reduce((a, b) => a + b, 0);
  const mStartX  = 0.18, mStartY = 1.03, mRowH = 0.43;

  slide.addShape(pres.shapes.RECTANGLE, { x: mStartX, y: mStartY, w: mTotalW, h: mRowH, fill: { color: C.orange }, line: { color: C.orange } });
  let mcx = mStartX;
  mTHeaders.forEach((h, ci) => {
    slide.addText(h, { x: mcx, y: mStartY + 0.07, w: mTColWs[ci], h: mRowH - 0.14, fontSize: 8.5, bold: true, color: C.white, align: "center", fontFace: "Arial", margin: 0 });
    mcx += mTColWs[ci];
  });
  milesData.forEach((m, ri) => {
    const ry = mStartY + mRowH + ri * mRowH;
    const isTop = ri === 0;
    slide.addShape(pres.shapes.RECTANGLE, { x: mStartX, y: ry, w: mTotalW, h: mRowH, fill: { color: ri % 2 === 0 ? C.tableRowBase : C.tableRowAlt }, line: { color: "333333" } });
    const vals = [m.miles, m.newSale, m.tg, String(m.usdot), String(m.declined), String(m.qCnt), m.qPrem, m.q2s, m.decR];
    mcx = mStartX;
    vals.forEach((cell, ci) => {
      const isNS = ci === 1;
      slide.addText(cell, {
        x: mcx, y: ry + 0.1, w: mTColWs[ci], h: mRowH - 0.2,
        fontSize: 8.5, bold: isNS || ci === 0,
        color: isNS ? C.orange : (ci === 0 && isTop ? C.orangeLight : C.white),
        align: "center", fontFace: "Calibri", margin: 0
      });
      mcx += mTColWs[ci];
    });
  });

  // Key insight below table
  const insightY = mStartY + mRowH * 7 + 0.1;
  slide.addShape(pres.shapes.RECTANGLE, { x: mStartX, y: insightY, w: mTotalW, h: 0.38, fill: { color: C.cardBg }, line: { color: C.orange, width: 1 } });
  slide.addText([
    { text: "Key Pattern: ", options: { bold: true, color: C.orange, fontSize: 9 } },
    { text: "500 miles + Unlimited = 70.2% of New Sale. 50 miles has best Q2S (70.59%). Focus submissions on 500-mile and Unlimited accounts for maximum revenue impact.", options: { color: C.white, fontSize: 9 } }
  ], { x: mStartX + 0.1, y: insightY + 0.05, w: mTotalW - 0.2, h: 0.3, fontFace: "Calibri", margin: 0 });
}

// ============================================================
// SLIDE 12 — Premium per # of Units (Part 1) - NEW
// ============================================================
{
  const slide = pres.addSlide();
  slide.background = { color: C.bg };

  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.65, fill: { color: C.orange }, line: { color: C.orange } });
  slide.addText("Premium Analysis by Number of Units — Part 1 of 2 · Units 1-5", {
    x: 0.3, y: 0.05, w: 9.4, h: 0.55, fontSize: 16, bold: true, color: C.white, fontFace: "Arial", margin: 0
  });

  // Data ordered by Premium New Sale (descending)
  const unitsData1 = [
    { units: "1", purePrem: "$5,342,211.76", newSale: "$811,582.06", tgPct: "42.79%", usdotCnt: "248", usdotDecl: "28", quotesCnt: "231", quotesPrem: "$2,197,513.59", q2s: "41.56%", declRatio: "18.06%" },
    { units: "3", purePrem: "$2,697,385.78", newSale: "$376,873.28", tgPct: "19.87%", usdotCnt: "22", usdotDecl: "4", quotesCnt: "45", quotesPrem: "$1,016,149.11", q2s: "37.78%", declRatio: "17.39%" },
    { units: "2", purePrem: "$2,012,050.28", newSale: "$267,832.39", tgPct: "14.12%", usdotCnt: "45", usdotDecl: "5", quotesCnt: "41", quotesPrem: "$906,415.82", q2s: "41.46%", declRatio: "13.16%" },
    { units: "4", purePrem: "$547,648.52", newSale: "$114,901.80", tgPct: "6.06%", usdotCnt: "7", usdotDecl: "1", quotesCnt: "8", quotesPrem: "$213,737.92", q2s: "87.50%", declRatio: "33.33%" },
    { units: "5", purePrem: "$896,821.60", newSale: "$114,375", tgPct: "6.03%", usdotCnt: "9", usdotDecl: "3", quotesCnt: "17", quotesPrem: "$517,677.10", q2s: "11.76%", declRatio: "42.86%" },
  ];

  const tHeaders = ["Units", "Pure Premium", "New Sale", "%TG", "USDOT Cnt", "Declined", "Quotes Cnt", "Quotes Prem", "Q2S", "Decl.%"];
  const tColWs = [0.5, 1.15, 1.05, 0.62, 0.75, 0.65, 0.75, 1.15, 0.65, 0.62];
  const tStartX = 0.15, tStartY = 0.75, rowH = 0.42;
  const tTotalW = tColWs.reduce((a, b) => a + b, 0);

  // Header
  slide.addShape(pres.shapes.RECTANGLE, { x: tStartX, y: tStartY, w: tTotalW, h: rowH, fill: { color: C.orange }, line: { color: C.orange } });
  let cx = tStartX;
  tHeaders.forEach((h, ci) => {
    slide.addText(h, { x: cx, y: tStartY + 0.07, w: tColWs[ci], h: rowH - 0.14, fontSize: 9, bold: true, color: C.white, align: "center", fontFace: "Arial", margin: 0 });
    cx += tColWs[ci];
  });

  // Rows
  unitsData1.forEach((row, ri) => {
    const ry = tStartY + rowH + ri * rowH;
    slide.addShape(pres.shapes.RECTANGLE, { x: tStartX, y: ry, w: tTotalW, h: rowH, fill: { color: ri % 2 === 0 ? C.tableRowBase : C.tableRowAlt }, line: { color: "333333" } });
    const vals = [row.units, row.purePrem, row.newSale, row.tgPct, row.usdotCnt, row.usdotDecl, row.quotesCnt, row.quotesPrem, row.q2s, row.declRatio];
    cx = tStartX;
    vals.forEach((v, ci) => {
      const isNewSale = ci === 2;
      const isHighQ2S = ci === 8 && parseFloat(row.q2s) > 40;
      slide.addText(v, { x: cx, y: ry + 0.08, w: tColWs[ci], h: rowH - 0.16, fontSize: 8.5, bold: isNewSale, color: isNewSale ? C.orange : (isHighQ2S ? C.green : C.white), align: "center", fontFace: "Calibri", margin: 0 });
      cx += tColWs[ci];
    });
  });

  // Insights
  slide.addShape(pres.shapes.RECTANGLE, { x: tStartX, y: tStartY + rowH * 6 + 0.1, w: tTotalW, h: 0.52, fill: { color: C.cardBg }, line: { color: C.orange, width: 1 } });
  slide.addText([
    { text: "Key Insights: ", options: { bold: true, color: C.orange, fontSize: 9 } },
    { text: "1 Unit accounts for 42.79% of new sales ($811K). ", options: { color: C.white, fontSize: 9 } },
    { text: "4 Units shows exceptional Q2S at 87.50%. ", options: { color: C.green, fontSize: 9 } },
    { text: "Units 1-3 represent 76.78% of total new sale premium. ", options: { color: C.white, fontSize: 9 } },
    { text: "Focus on single-unit operations for volume, multi-unit for conversion quality.", options: { color: C.midGray, fontSize: 9 } },
  ], { x: tStartX + 0.1, y: tStartY + rowH * 6 + 0.16, w: tTotalW - 0.2, h: 0.4, fontFace: "Calibri", margin: 0 });
}

// ============================================================
// SLIDE 13 — Premium per # of Units (Part 2) - NEW
// ============================================================
{
  const slide = pres.addSlide();
  slide.background = { color: C.bg };

  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.65, fill: { color: C.orange }, line: { color: C.orange } });
  slide.addText("Premium Analysis by Number of Units — Part 2 of 2 · Units 6-10", {
    x: 0.3, y: 0.05, w: 9.4, h: 0.55, fontSize: 16, bold: true, color: C.white, fontFace: "Arial", margin: 0
  });

  const unitsData2 = [
    { units: "10", purePrem: "$658,432.28", newSale: "$211,195.26", tgPct: "11.13%", usdotCnt: "1", usdotDecl: "-", quotesCnt: "2", quotesPrem: "$134,237.02", q2s: "50.00%", declRatio: "-" },
    { units: "6", purePrem: "$929,980.45", newSale: "-", tgPct: "-", usdotCnt: "5", usdotDecl: "4", quotesCnt: "4", quotesPrem: "$390,548.45", q2s: "-", declRatio: "50.00%" },
    { units: "8", purePrem: "$699,055.06", newSale: "-", tgPct: "-", usdotCnt: "3", usdotDecl: "2", quotesCnt: "-", quotesPrem: "-", q2s: "-", declRatio: "66.67%" },
    { units: "9", purePrem: "$429,048.00", newSale: "-", tgPct: "-", usdotCnt: "1", usdotDecl: "-", quotesCnt: "1", quotesPrem: "$214,524.00", q2s: "-", declRatio: "-" },
    { units: "7", purePrem: "$207,694.00", newSale: "-", tgPct: "-", usdotCnt: "3", usdotDecl: "2", quotesCnt: "-", quotesPrem: "-", q2s: "-", declRatio: "33.33%" },
  ];

  const tHeaders = ["Units", "Pure Premium", "New Sale", "%TG", "USDOT Cnt", "Declined", "Quotes Cnt", "Quotes Prem", "Q2S", "Decl.%"];
  const tColWs = [0.5, 1.15, 1.05, 0.62, 0.75, 0.65, 0.75, 1.15, 0.65, 0.62];
  const tStartX = 0.15, tStartY = 0.75, rowH = 0.42;
  const tTotalW = tColWs.reduce((a, b) => a + b, 0);

  // Header
  slide.addShape(pres.shapes.RECTANGLE, { x: tStartX, y: tStartY, w: tTotalW, h: rowH, fill: { color: C.orange }, line: { color: C.orange } });
  let cx = tStartX;
  tHeaders.forEach((h, ci) => {
    slide.addText(h, { x: cx, y: tStartY + 0.07, w: tColWs[ci], h: rowH - 0.14, fontSize: 9, bold: true, color: C.white, align: "center", fontFace: "Arial", margin: 0 });
    cx += tColWs[ci];
  });

  // Rows
  unitsData2.forEach((row, ri) => {
    const ry = tStartY + rowH + ri * rowH;
    slide.addShape(pres.shapes.RECTANGLE, { x: tStartX, y: ry, w: tTotalW, h: rowH, fill: { color: ri % 2 === 0 ? C.tableRowBase : C.tableRowAlt }, line: { color: "333333" } });
    const vals = [row.units, row.purePrem, row.newSale, row.tgPct, row.usdotCnt, row.usdotDecl, row.quotesCnt, row.quotesPrem, row.q2s, row.declRatio];
    cx = tStartX;
    vals.forEach((v, ci) => {
      const isNewSale = ci === 2;
      slide.addText(v, { x: cx, y: ry + 0.08, w: tColWs[ci], h: rowH - 0.16, fontSize: 8.5, bold: isNewSale, color: isNewSale && v !== "-" ? C.orange : C.white, align: "center", fontFace: "Calibri", margin: 0 });
      cx += tColWs[ci];
    });
  });

  // Insights
  slide.addShape(pres.shapes.RECTANGLE, { x: tStartX, y: tStartY + rowH * 6 + 0.1, w: tTotalW, h: 0.52, fill: { color: C.cardBg }, line: { color: C.orange, width: 1 } });
  slide.addText([
    { text: "Key Insights: ", options: { bold: true, color: C.orange, fontSize: 9 } },
    { text: "Only 10 Units shows new sale activity ($211K) with 50% Q2S. ", options: { color: C.white, fontSize: 9 } },
    { text: "Higher unit categories (6-9) show quote activity but no conversions. ", options: { color: C.red, fontSize: 9 } },
    { text: "Declination ratios increase significantly: 6+ units range from 33-67%. ", options: { color: C.amber, fontSize: 9 } },
    { text: "Opportunity exists to improve conversion in 6-9 unit segment with targeted underwriting.", options: { color: C.midGray, fontSize: 9 } },
  ], { x: tStartX + 0.1, y: tStartY + rowH * 6 + 0.16, w: tTotalW - 0.2, h: 0.4, fontFace: "Calibri", margin: 0 });
}

// ============================================================
// SLIDE 14 — Premium per Years in Business (Part 1) - NEW
// ============================================================
{
  const slide = pres.addSlide();
  slide.background = { color: C.bg };

  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.65, fill: { color: C.orange }, line: { color: C.orange } });
  slide.addText("Premium Analysis by Years in Business — Part 1 of 3 · Years 0-4", {
    x: 0.3, y: 0.05, w: 9.4, h: 0.55, fontSize: 16, bold: true, color: C.white, fontFace: "Arial", margin: 0
  });

  const yearsData1 = [
    { years: "0", purePrem: "$6,546,349.32", newSale: "$1,115,139.34", tgPct: "58.79%", usdotCnt: "178", usdotDecl: "18", quotesCnt: "199", quotesPrem: "$3,183,827.75", pureQ2S: "35.03%", q2s: "43.72%", declRatio: "16.98%" },
    { years: "1", purePrem: "$2,100,690.85", newSale: "$239,654.88", tgPct: "12.63%", usdotCnt: "46", usdotDecl: "5", quotesCnt: "34", quotesPrem: "$541,330.22", pureQ2S: "44.27%", q2s: "47.06%", declRatio: "14.29%" },
    { years: "3", purePrem: "$1,303,659.71", newSale: "$166,589.49", tgPct: "8.78%", usdotCnt: "26", usdotDecl: "5", quotesCnt: "45", quotesPrem: "$590,114.82", pureQ2S: "28.23%", q2s: "28.89%", declRatio: "27.78%" },
    { years: "2", purePrem: "$723,600.02", newSale: "$139,961.88", tgPct: "7.38%", usdotCnt: "27", usdotDecl: "5", quotesCnt: "25", quotesPrem: "$273,003.41", pureQ2S: "51.27%", q2s: "56.00%", declRatio: "22.73%" },
    { years: "4", purePrem: "$673,764.81", newSale: "$18,843.00", tgPct: "0.99%", usdotCnt: "22", usdotDecl: "4", quotesCnt: "19", quotesPrem: "$217,477.16", pureQ2S: "8.66%", q2s: "22.22%", declRatio: "20.00%" },
  ];

  const tHeaders = ["Years", "Pure Premium", "New Sale", "%TG", "USDOT", "Decl.", "Quotes", "Q. Prem", "Pure Q→S", "Q2S", "Dec.%"];
  const tColWs = [0.45, 1.05, 1.0, 0.6, 0.6, 0.5, 0.6, 1.05, 0.72, 0.6, 0.6];
  const tStartX = 0.12, tStartY = 0.75, rowH = 0.42;
  const tTotalW = tColWs.reduce((a, b) => a + b, 0);

  // Header
  slide.addShape(pres.shapes.RECTANGLE, { x: tStartX, y: tStartY, w: tTotalW, h: rowH, fill: { color: C.orange }, line: { color: C.orange } });
  let cx = tStartX;
  tHeaders.forEach((h, ci) => {
    slide.addText(h, { x: cx, y: tStartY + 0.06, w: tColWs[ci], h: rowH - 0.12, fontSize: 8.5, bold: true, color: C.white, align: "center", fontFace: "Arial", margin: 0 });
    cx += tColWs[ci];
  });

  // Rows
  yearsData1.forEach((row, ri) => {
    const ry = tStartY + rowH + ri * rowH;
    slide.addShape(pres.shapes.RECTANGLE, { x: tStartX, y: ry, w: tTotalW, h: rowH, fill: { color: ri % 2 === 0 ? C.tableRowBase : C.tableRowAlt }, line: { color: "333333" } });
    const vals = [row.years, row.purePrem, row.newSale, row.tgPct, row.usdotCnt, row.usdotDecl, row.quotesCnt, row.quotesPrem, row.pureQ2S, row.q2s, row.declRatio];
    cx = tStartX;
    vals.forEach((v, ci) => {
      const isNewSale = ci === 2;
      const isHighQ2S = ci === 9 && parseFloat(row.q2s) > 45;
      slide.addText(v, { x: cx, y: ry + 0.08, w: tColWs[ci], h: rowH - 0.16, fontSize: 8, bold: isNewSale, color: isNewSale ? C.orange : (isHighQ2S ? C.green : C.white), align: "center", fontFace: "Calibri", margin: 0 });
      cx += tColWs[ci];
    });
  });

  // Insights
  slide.addShape(pres.shapes.RECTANGLE, { x: tStartX, y: tStartY + rowH * 6 + 0.1, w: tTotalW, h: 0.52, fill: { color: C.cardBg }, line: { color: C.orange, width: 1 } });
  slide.addText([
    { text: "Key Insights: ", options: { bold: true, color: C.orange, fontSize: 9 } },
    { text: "0 Years dominates with 58.79% of new sales ($1.12M) - strong new business acquisition. ", options: { color: C.white, fontSize: 9 } },
    { text: "2 Years shows best Q2S at 56.00%. ", options: { color: C.green, fontSize: 9 } },
    { text: "Years 0-3 account for 87.58% of total new sale premium. ", options: { color: C.white, fontSize: 9 } },
    { text: "4 Years underperforms significantly (0.99% %TG) - requires investigation.", options: { color: C.red, fontSize: 9 } },
  ], { x: tStartX + 0.1, y: tStartY + rowH * 6 + 0.16, w: tTotalW - 0.2, h: 0.4, fontFace: "Calibri", margin: 0 });
}

// ============================================================
// SLIDE 15 — Premium per Years in Business (Part 2) - NEW
// ============================================================
{
  const slide = pres.addSlide();
  slide.background = { color: C.bg };

  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.65, fill: { color: C.orange }, line: { color: C.orange } });
  slide.addText("Premium Analysis by Years in Business — Part 2 of 3 · Years 5-10", {
    x: 0.3, y: 0.05, w: 9.4, h: 0.55, fontSize: 16, bold: true, color: C.white, fontFace: "Arial", margin: 0
  });

  const yearsData2 = [
    { years: "10", purePrem: "$358,146.80", newSale: "$87,244.20", tgPct: "4.60%", usdotCnt: "5", usdotDecl: "2", quotesCnt: "3", quotesPrem: "$119,400.20", pureQ2S: "73.07%", q2s: "66.67%", declRatio: "33.33%" },
    { years: "7", purePrem: "$151,786.00", newSale: "$23,524.00", tgPct: "1.24%", usdotCnt: "5", usdotDecl: "1", quotesCnt: "2", quotesPrem: "$23,524.00", pureQ2S: "100.00%", q2s: "100.00%", declRatio: "33.33%" },
    { years: "8", purePrem: "$126,078.60", newSale: "$22,000.00", tgPct: "1.16%", usdotCnt: "4", usdotDecl: "1", quotesCnt: "5", quotesPrem: "$65,201.60", pureQ2S: "33.74%", q2s: "20.00%", declRatio: "25.00%" },
    { years: "5", purePrem: "$834,685.67", newSale: "-", tgPct: "-", usdotCnt: "12", usdotDecl: "3", quotesCnt: "17", quotesPrem: "$413,013.10", pureQ2S: "-", q2s: "-", declRatio: "30.00%" },
    { years: "6", purePrem: "$16,379.00", newSale: "-", tgPct: "-", usdotCnt: "5", usdotDecl: "1", quotesCnt: "1", quotesPrem: "$16,379.00", pureQ2S: "-", q2s: "-", declRatio: "25.00%" },
    { years: "9", purePrem: "$446,758.95", newSale: "-", tgPct: "-", usdotCnt: "4", usdotDecl: "2", quotesCnt: "2", quotesPrem: "$75,825.75", pureQ2S: "-", q2s: "-", declRatio: "28.57%" },
  ];

  const tHeaders = ["Years", "Pure Premium", "New Sale", "%TG", "USDOT", "Decl.", "Quotes", "Q. Prem", "Pure Q→S", "Q2S", "Dec.%"];
  const tColWs = [0.45, 1.05, 1.0, 0.6, 0.6, 0.5, 0.6, 1.05, 0.72, 0.6, 0.6];
  const tStartX = 0.12, tStartY = 0.75, rowH = 0.42;
  const tTotalW = tColWs.reduce((a, b) => a + b, 0);

  // Header
  slide.addShape(pres.shapes.RECTANGLE, { x: tStartX, y: tStartY, w: tTotalW, h: rowH, fill: { color: C.orange }, line: { color: C.orange } });
  let cx = tStartX;
  tHeaders.forEach((h, ci) => {
    slide.addText(h, { x: cx, y: tStartY + 0.06, w: tColWs[ci], h: rowH - 0.12, fontSize: 8.5, bold: true, color: C.white, align: "center", fontFace: "Arial", margin: 0 });
    cx += tColWs[ci];
  });

  // Rows
  yearsData2.forEach((row, ri) => {
    const ry = tStartY + rowH + ri * rowH;
    slide.addShape(pres.shapes.RECTANGLE, { x: tStartX, y: ry, w: tTotalW, h: rowH, fill: { color: ri % 2 === 0 ? C.tableRowBase : C.tableRowAlt }, line: { color: "333333" } });
    const vals = [row.years, row.purePrem, row.newSale, row.tgPct, row.usdotCnt, row.usdotDecl, row.quotesCnt, row.quotesPrem, row.pureQ2S, row.q2s, row.declRatio];
    cx = tStartX;
    vals.forEach((v, ci) => {
      const isNewSale = ci === 2;
      const isHighQ2S = ci === 9 && v !== "-" && parseFloat(row.q2s) > 60;
      slide.addText(v, { x: cx, y: ry + 0.08, w: tColWs[ci], h: rowH - 0.16, fontSize: 8, bold: isNewSale, color: isNewSale && v !== "-" ? C.orange : (isHighQ2S ? C.green : C.white), align: "center", fontFace: "Calibri", margin: 0 });
      cx += tColWs[ci];
    });
  });

  // Insights
  slide.addShape(pres.shapes.RECTANGLE, { x: tStartX, y: tStartY + rowH * 7 + 0.1, w: tTotalW, h: 0.52, fill: { color: C.cardBg }, line: { color: C.orange, width: 1 } });
  slide.addText([
    { text: "Key Insights: ", options: { bold: true, color: C.orange, fontSize: 9 } },
    { text: "7 Years shows perfect conversion: 100% Q2S. ", options: { color: C.green, fontSize: 9 } },
    { text: "10 Years demonstrates strong performance with 66.67% Q2S. ", options: { color: C.green, fontSize: 9 } },
    { text: "Years 5, 6, and 9 show quote activity but no new sales - conversion opportunity. ", options: { color: C.amber, fontSize: 9 } },
    { text: "Mature businesses (7-10 years) show better conversion rates when they convert.", options: { color: C.midGray, fontSize: 9 } },
  ], { x: tStartX + 0.1, y: tStartY + rowH * 7 + 0.16, w: tTotalW - 0.2, h: 0.4, fontFace: "Calibri", margin: 0 });
}

// ============================================================
// SLIDE 16 — Premium per Years in Business (Part 3) - NEW
// ============================================================
{
  const slide = pres.addSlide();
  slide.background = { color: C.bg };

  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.65, fill: { color: C.orange }, line: { color: C.orange } });
  slide.addText("Premium Analysis by Years in Business — Part 3 of 3 · Years 11-28", {
    x: 0.3, y: 0.05, w: 9.4, h: 0.55, fontSize: 16, bold: true, color: C.white, fontFace: "Arial", margin: 0
  });

  const yearsData3 = [
    { years: "11", purePrem: "$187,275.00", newSale: "$51,075.00", tgPct: "2.69%", usdotCnt: "2", usdotDecl: "1", quotesCnt: "1", quotesPrem: "$136,200.00", pureQ2S: "37.50%", q2s: "50.00%", declRatio: "50.00%" },
    { years: "12", purePrem: "$98,379.00", newSale: "$32,728.00", tgPct: "1.73%", usdotCnt: "2", usdotDecl: "-", quotesCnt: "1", quotesPrem: "$32,728.00", pureQ2S: "100.00%", q2s: "100.00%", declRatio: "-" },
    { years: "14", purePrem: "$22,324.00", newSale: "-", tgPct: "-", usdotCnt: "1", usdotDecl: "-", quotesCnt: "1", quotesPrem: "$1,254.00", pureQ2S: "-", q2s: "-", declRatio: "-" },
    { years: "20", purePrem: "$438,002.00", newSale: "-", tgPct: "-", usdotCnt: "1", usdotDecl: "-", quotesCnt: "-", quotesPrem: "-", pureQ2S: "-", q2s: "-", declRatio: "-" },
    { years: "28", purePrem: "$429,048.00", newSale: "-", tgPct: "-", usdotCnt: "1", usdotDecl: "-", quotesCnt: "1", quotesPrem: "$214,524.00", pureQ2S: "-", q2s: "-", declRatio: "-" },
  ];

  const tHeaders = ["Years", "Pure Premium", "New Sale", "%TG", "USDOT", "Decl.", "Quotes", "Q. Prem", "Pure Q→S", "Q2S", "Dec.%"];
  const tColWs = [0.45, 1.05, 1.0, 0.6, 0.6, 0.5, 0.6, 1.05, 0.72, 0.6, 0.6];
  const tStartX = 0.12, tStartY = 0.75, rowH = 0.42;
  const tTotalW = tColWs.reduce((a, b) => a + b, 0);

  // Header
  slide.addShape(pres.shapes.RECTANGLE, { x: tStartX, y: tStartY, w: tTotalW, h: rowH, fill: { color: C.orange }, line: { color: C.orange } });
  let cx = tStartX;
  tHeaders.forEach((h, ci) => {
    slide.addText(h, { x: cx, y: tStartY + 0.06, w: tColWs[ci], h: rowH - 0.12, fontSize: 8.5, bold: true, color: C.white, align: "center", fontFace: "Arial", margin: 0 });
    cx += tColWs[ci];
  });

  // Rows
  yearsData3.forEach((row, ri) => {
    const ry = tStartY + rowH + ri * rowH;
    slide.addShape(pres.shapes.RECTANGLE, { x: tStartX, y: ry, w: tTotalW, h: rowH, fill: { color: ri % 2 === 0 ? C.tableRowBase : C.tableRowAlt }, line: { color: "333333" } });
    const vals = [row.years, row.purePrem, row.newSale, row.tgPct, row.usdotCnt, row.usdotDecl, row.quotesCnt, row.quotesPrem, row.pureQ2S, row.q2s, row.declRatio];
    cx = tStartX;
    vals.forEach((v, ci) => {
      const isNewSale = ci === 2;
      const isHighQ2S = ci === 9 && v !== "-" && parseFloat(row.q2s) === 100;
      slide.addText(v, { x: cx, y: ry + 0.08, w: tColWs[ci], h: rowH - 0.16, fontSize: 8, bold: isNewSale, color: isNewSale && v !== "-" ? C.orange : (isHighQ2S ? C.green : C.white), align: "center", fontFace: "Calibri", margin: 0 });
      cx += tColWs[ci];
    });
  });

  // Insights
  slide.addShape(pres.shapes.RECTANGLE, { x: tStartX, y: tStartY + rowH * 6 + 0.1, w: tTotalW, h: 0.52, fill: { color: C.cardBg }, line: { color: C.orange, width: 1 } });
  slide.addText([
    { text: "Key Insights: ", options: { bold: true, color: C.orange, fontSize: 9 } },
    { text: "12 Years shows perfect 100% Q2S conversion. ", options: { color: C.green, fontSize: 9 } },
    { text: "Very limited volume in 11+ year segment - only 2 conversions total. ", options: { color: C.white, fontSize: 9 } },
    { text: "Years 14, 20, and 28 show pure premium but no new sales - niche/specialty cases. ", options: { color: C.amber, fontSize: 9 } },
    { text: "Established businesses (11-12 years) convert well but represent small market share (4.42% combined).", options: { color: C.midGray, fontSize: 9 } },
  ], { x: tStartX + 0.1, y: tStartY + rowH * 6 + 0.16, w: tTotalW - 0.2, h: 0.4, fontFace: "Calibri", margin: 0 });
}

// ============================================================
// SLIDE 17 — Action Plan & Winning Patterns (UPDATED)
// ============================================================
{
  const slide = pres.addSlide();
  slide.background = { color: C.bg };

  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.65, fill: { color: C.orange }, line: { color: C.orange } });
  slide.addText("Q1 2026 Action Plan — Winning Patterns · Sede Convicción", {
    x: 0.3, y: 0.05, w: 9.4, h: 0.55, fontSize: 15, bold: true, color: C.white, fontFace: "Arial", margin: 0
  });

  // LEFT — Critical Alerts
  const alertItems = [
    "2 producers with $0 New Sale (Felipe Mejia + Daniela Salas).",
    "Office at 29.4% of Q4 goal ($6.45M) — critical gap.",
    "Declination rate varied: 15% → 20% → 24%.",
    "Daniela Salas: $135K in quotes not closed.",
    "Yefrik Alvarez only 2.54% %TG — needs intervention.",
  ];
  slide.addShape(pres.shapes.RECTANGLE, { x: 0.15, y: 0.72, w: 3.1, h: 0.3, fill: { color: C.red }, line: { color: C.red } });
  slide.addText("✕ Q4 Critical Alerts", { x: 0.15, y: 0.74, w: 3.1, h: 0.26, fontSize: 9.5, bold: true, color: C.white, align: "center", fontFace: "Arial", margin: 0 });
  slide.addShape(pres.shapes.RECTANGLE, { x: 0.15, y: 1.02, w: 3.1, h: alertItems.length * 0.3 + 0.1, fill: { color: C.cardBg }, line: { color: C.red, width: 1 } });
  alertItems.forEach((item, j) => {
    slide.addText([
      { text: "› ", options: { bold: true, color: C.red } },
      { text: item, options: { color: C.white } }
    ], { x: 0.25, y: 1.06 + j * 0.3, w: 2.9, h: 0.28, fontSize: 7.8, fontFace: "Calibri", margin: 0 });
  });

  // CENTER — Winning Patterns (UPDATED WITH UNITS & YEARS)
  const patternY = 0.72;
  slide.addShape(pres.shapes.RECTANGLE, { x: 3.4, y: patternY, w: 6.45, h: 0.3, fill: { color: C.orangeDark }, line: { color: C.orangeDark } });
  slide.addText("🎯 Winning Patterns — What the Office Must Replicate to Sell More", {
    x: 3.4, y: patternY + 0.03, w: 6.45, h: 0.26, fontSize: 10, bold: true, color: C.white, align: "center", fontFace: "Arial", margin: 0
  });

  const patterns = [
    {
      icon: "🚛", title: "Units: Truck Tractor + 1-3 Units",
      detail: "Truck Tractor = $1.48M (78%). Units 1-3 = 76.78% of revenue. Unit 4 has 87.50% Q2S - model for optimization."
    },
    {
      icon: "📦", title: "Business: General Freight + Building",
      detail: "General Freight: $866K (51.69%). Building Materials: $465K with excellent 53.85% Q2S. Dirt/Sand: $219K with 63.64% Q2S."
    },
    {
      icon: "👶", title: "Years: New Business (0-1 Years)",
      detail: "0 Years = 58.79% of New Sale ($1.12M). Years 0-3 combined = 87.58%. Strong new business acquisition engine."
    },
    {
      icon: "📍", title: "Miles: 500 mi + Unlimited",
      detail: "500 miles + Unlimited = 70.2% of total. 50 miles has best Q2S (70.59%). Focus on these mileage bands."
    },
    {
      icon: "🤝", title: "MGA: Nexus (56.79%) + Motor Transport",
      detail: "Nexus: best Q2S (56.79%) and low declination (10.45%). Motor Transport: second highest New Sale ($405K)."
    },
    {
      icon: "🏷️", title: "Carrier: Pegaso + Years 2,7,10,12",
      detail: "Pegaso Q2S 79.31%. Year 2: 56% Q2S. Year 7: 100% Q2S. Year 10: 66.67%. Year 12: 100%. Mature converts reliably."
    },
  ];

  const pCols = 3, pRows = 2;
  const pW = 2.08, pH = 1.0, pGap = 0.06;
  patterns.forEach((p, i) => {
    const col = i % pCols;
    const row = Math.floor(i / pCols);
    const px = 3.4 + col * (pW + pGap);
    const py = patternY + 0.35 + row * (pH + 0.06);
    slide.addShape(pres.shapes.RECTANGLE, { x: px, y: py, w: pW, h: pH, fill: { color: C.cardBg }, line: { color: C.orange, width: 1 } });
    slide.addText(`${p.icon} ${p.title}`, { x: px + 0.06, y: py + 0.04, w: pW - 0.12, h: 0.24, fontSize: 7.5, bold: true, color: C.orange, fontFace: "Arial", margin: 0 });
    slide.addText(p.detail, { x: px + 0.06, y: py + 0.28, w: pW - 0.12, h: pH - 0.32, fontSize: 6.8, color: C.midGray, fontFace: "Calibri", margin: 0 });
  });

  // Bottom — Priority Action Plan
  const apY = patternY + 0.35 + 2 * (1.0 + 0.06) + 0.08;
  slide.addShape(pres.shapes.RECTANGLE, { x: 0.15, y: apY, w: 9.7, h: 0.25, fill: { color: C.orange }, line: { color: C.orange } });
  slide.addText("Q1 2026 — Priority Actions", {
    x: 0.15, y: apY + 0.03, w: 9.7, h: 0.2, fontSize: 10, bold: true, color: C.white, align: "center", fontFace: "Arial", margin: 0
  });

  const actions = [
    ["1", "Close Daniela Salas $135K idle pipeline", "Immediate action"],
    ["2", "Activate Felipe Mejia", "$25K quotes pending"],
    ["3", "Optimize 1-3 Units segment", "76.78% of revenue - improve speed"],
    ["4", "Prioritize Pegaso + Nexus on every account", "79% and 57% Q2S"],
    ["5", "Focus on Truck Tractor + 500mi/Unlimited", "Highest revenue combo"],
    ["6", "Maintain 0-1 Years focus", "71.42% combined - strong acquisition"],
  ];
  const aStartY = apY + 0.27;
  const aRowH = 0.2;
  actions.forEach((a, i) => {
    const col = i < 3 ? 0 : 1;
    const row = i < 3 ? i : i - 3;
    const ax = 0.15 + col * 4.9;
    const ay = aStartY + row * aRowH;
    slide.addShape(pres.shapes.RECTANGLE, { x: ax, y: ay, w: 0.22, h: 0.17, fill: { color: C.orange }, line: { color: C.orange } });
    slide.addText(a[0], { x: ax, y: ay, w: 0.22, h: 0.17, fontSize: 7, bold: true, color: C.white, align: "center", fontFace: "Arial Black", margin: 0 });
    slide.addText(`${a[1]}: `, { x: ax + 0.25, y: ay, w: 2.2, h: 0.17, fontSize: 7, bold: true, color: C.orange, fontFace: "Calibri", margin: 0 });
    slide.addText(a[2], { x: ax + 2.45, y: ay, w: 2.2, h: 0.17, fontSize: 7, color: C.midGray, fontFace: "Calibri", margin: 0 });
  });
}
// ============================================================
// SLIDE 18 — Top 10 Producer Fees (NEW)
// ============================================================
{
  const slide = pres.addSlide();
  slide.background = { color: C.bg };

  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.65, fill: { color: C.orange }, line: { color: C.orange } });
  slide.addText("Producer Fees Ranking — Top 10 Q4 2025", {
    x: 0.3, y: 0.05, w: 9.4, h: 0.55, fontSize: 16, bold: true, color: C.white, fontFace: "Arial", margin: 0
  });

  // Top 10 data from screenshot
  const feesData = [
    ["1", "Ronald Alexander Azuaje Perez", "$17,250.00"],
    ["2", "Anderson Blanco", "$13,300.00"],
    ["3", "Melisa Negrete Cuello", "$12,600.00"],
    ["4", "Carlos Andres Hernandez Rivera", "$9,700.00"],
    ["5", "Camila Alvarez", "$9,482.00"],
    ["6", "Juan Jose Restrepo", "$8,061.00"],
    ["7", "Yuliana Andrea Valez Zapata", "$7,650.00"],
    ["8", "Paulina Restrepo Ospina", "$7,000.00"],
    ["9", "Frank Derley David Betancur", "$7,000.00"],
    ["10", "Edison Alberto Rodriguez Muñoz", "$6,100.00"],
  ];

  const headers = ["Rank", "Producer Name", "Producer Fees"];
  const colWidths = [0.6, 5.5, 1.5];
  const startX = 1.5, startY = 0.85, rowHeight = 0.38;
  const totalW = colWidths.reduce((a,b) => a+b, 0);

  // Header
  slide.addShape(pres.shapes.RECTANGLE, { x: startX, y: startY, w: totalW, h: rowHeight, fill: { color: C.orange }, line: { color: C.orange } });
  let cx = startX;
  headers.forEach((h, i) => {
    slide.addText(h, { x: cx, y: startY + 0.06, w: colWidths[i], h: rowHeight - 0.12, fontSize: 11, bold: true, color: C.white, align: "center", fontFace: "Arial", margin: 0 });
    cx += colWidths[i];
  });

  // Rows
  feesData.forEach((row, ri) => {
    const ry = startY + rowHeight + ri * rowHeight;
    const rowBg = ri % 2 === 0 ? C.tableRowBase : C.tableRowAlt;
    slide.addShape(pres.shapes.RECTANGLE, { x: startX, y: ry, w: totalW, h: rowHeight, fill: { color: rowBg }, line: { color: "333333" } });
    
    cx = startX;
    row.forEach((cell, ci) => {
      const isFee = ci === 2;
      const isRank = ci === 0;
      slide.addText(cell, { 
        x: cx, y: ry + 0.08, w: colWidths[ci], h: rowHeight - 0.16, 
        fontSize: 10, bold: isFee || isRank, 
        color: isFee ? C.orange : C.white, 
        align: ci === 1 ? "left" : "center", 
        fontFace: "Calibri", margin: 0 
      });
      cx += colWidths[ci];
    });
  });

  // Total
  const totalY = startY + rowHeight * 11 + 0.1;
  slide.addShape(pres.shapes.RECTANGLE, { x: startX, y: totalY, w: totalW, h: 0.45, fill: { color: C.orangeDark }, line: { color: C.orangeDark } });
  slide.addText("TOTAL PRODUCER FEES (TOP 10): $98,143.00", { 
    x: startX, y: totalY + 0.08, w: totalW, h: 0.3, 
    fontSize: 12, bold: true, color: C.white, align: "center", 
    fontFace: "Arial Black", margin: 0 
  });

  // Note
  slide.addShape(pres.shapes.RECTANGLE, { x: startX, y: totalY + 0.55, w: totalW, h: 0.4, fill: { color: C.cardBg }, line: { color: C.orange, width: 1 } });
  slide.addText([
    { text: "Note: ", options: { bold: true, color: C.orange, fontSize: 9 } },
    { text: "Complete office total including all producers: $191,349.04. ", options: { color: C.white, fontSize: 9 } },
    { text: "Top 10 producers account for 51.3% of total fees.", options: { color: C.midGray, fontSize: 9 } },
  ], { x: startX + 0.15, y: totalY + 0.65, w: totalW - 0.3, h: 0.25, fontFace: "Calibri", margin: 0 });
}

console.log("✓ Slide 18: Top 10 Producer Fees");

// ============================================================
// SLIDE 19 — Nexus vs Trinity Comparison (NEW)
// ============================================================
{
  const slide = pres.addSlide();
  slide.background = { color: C.bg };

  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.65, fill: { color: C.orange }, line: { color: C.orange } });
  slide.addText("MGA Performance Comparison — Nexus vs Trinity Q4 2025", {
    x: 0.3, y: 0.05, w: 9.4, h: 0.55, fontSize: 16, bold: true, color: C.white, fontFace: "Arial", margin: 0
  });

  // Nexus data
  const nexusData = {
    name: "Nexus Risk Management",
    purePrem: "$522,592.13",
    salesCount: "23",
    newSale: "$42,351.60",
    usdot: "55",
    declined: "5",
    quotesPrem: "$145,879.26",
    quotesCount: "49",
    q2s: "46.94%",
    declRatio: "9.43%"
  };

  // Trinity data
  const trinityData = {
    name: "Trinity Underwriters",
    purePrem: "$264,946.76",
    salesCount: "19",
    newSale: "$46,098.38",
    usdot: "58",
    declined: "1",
    quotesPrem: "$206,123.38",
    quotesCount: "71",
    q2s: "26.76%",
    declRatio: "50.00%"
  };

  // Left side - Nexus
  const leftX = 0.3, cardW = 4.6, cardY = 0.8;
  slide.addShape(pres.shapes.RECTANGLE, { x: leftX, y: cardY, w: cardW, h: 0.5, fill: { color: C.orange }, line: { color: C.orange } });
  slide.addText("NEXUS RISK MANAGEMENT", { 
    x: leftX, y: cardY + 0.05, w: cardW, h: 0.4, 
    fontSize: 13, bold: true, color: C.white, align: "center", 
    fontFace: "Arial Black", margin: 0 
  });

  slide.addShape(pres.shapes.RECTANGLE, { x: leftX, y: cardY + 0.5, w: cardW, h: 3.8, fill: { color: C.cardBg }, line: { color: C.orange, width: 2 } });

  const nexusStats = [
    ["New Sale", nexusData.newSale, true],
    ["Sales Count", nexusData.salesCount, false],
    ["Pure Premium", nexusData.purePrem, false],
    ["USDOT Count", nexusData.usdot, false],
    ["USDOT Declined", nexusData.declined, false],
    ["Quotes Premium", nexusData.quotesPrem, false],
    ["Quotes Count", nexusData.quotesCount, false],
    ["Quote to Sold Ratio", nexusData.q2s, true],
    ["Declination Ratio", nexusData.declRatio, true],
  ];

  nexusStats.forEach((stat, i) => {
    const sy = cardY + 0.65 + i * 0.4;
    slide.addText(stat[0], { x: leftX + 0.15, y: sy, w: cardW - 0.3, h: 0.18, fontSize: 9, color: C.midGray, fontFace: "Calibri", margin: 0 });
    slide.addText(stat[1], { 
      x: leftX + 0.15, y: sy + 0.18, w: cardW - 0.3, h: 0.2, 
      fontSize: stat[2] ? 16 : 13, bold: true, 
      color: stat[2] ? C.orange : C.white, 
      fontFace: "Arial Black", margin: 0 
    });
  });

  // Right side - Trinity
  const rightX = 5.1;
  slide.addShape(pres.shapes.RECTANGLE, { x: rightX, y: cardY, w: cardW, h: 0.5, fill: { color: C.orangeDark }, line: { color: C.orangeDark } });
  slide.addText("TRINITY UNDERWRITERS", { 
    x: rightX, y: cardY + 0.05, w: cardW, h: 0.4, 
    fontSize: 13, bold: true, color: C.white, align: "center", 
    fontFace: "Arial Black", margin: 0 
  });

  slide.addShape(pres.shapes.RECTANGLE, { x: rightX, y: cardY + 0.5, w: cardW, h: 3.8, fill: { color: C.cardBg }, line: { color: C.orangeDark, width: 2 } });

  const trinityStats = [
    ["New Sale", trinityData.newSale, true],
    ["Sales Count", trinityData.salesCount, false],
    ["Pure Premium", trinityData.purePrem, false],
    ["USDOT Count", trinityData.usdot, false],
    ["USDOT Declined", trinityData.declined, false],
    ["Quotes Premium", trinityData.quotesPrem, false],
    ["Quotes Count", trinityData.quotesCount, false],
    ["Quote to Sold Ratio", trinityData.q2s, false],
    ["Declination Ratio", trinityData.declRatio, true],
  ];

  trinityStats.forEach((stat, i) => {
    const sy = cardY + 0.65 + i * 0.4;
    slide.addText(stat[0], { x: rightX + 0.15, y: sy, w: cardW - 0.3, h: 0.18, fontSize: 9, color: C.midGray, fontFace: "Calibri", margin: 0 });
    slide.addText(stat[1], { 
      x: rightX + 0.15, y: sy + 0.18, w: cardW - 0.3, h: 0.2, 
      fontSize: stat[2] ? 16 : 13, bold: true, 
      color: stat[2] ? C.orange : C.white, 
      fontFace: "Arial Black", margin: 0 
    });
  });

  // Key insights
  const insightY = 4.6;
  slide.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: insightY, w: 9.4, h: 0.85, fill: { color: C.cardBg }, line: { color: C.orange, width: 1 } });
  slide.addText([
    { text: "Key Insights: ", options: { bold: true, color: C.orange, fontSize: 11 } },
    { text: "Trinity leads in New Sale ($46K vs $42K) but has significant conversion challenges. ", options: { color: C.white, fontSize: 10 } },
    { text: "Nexus Q2S: 46.94% vs Trinity 26.76% (+20.18pp advantage). ", options: { color: C.white, fontSize: 10 } },
    { text: "Nexus Declination: 9.43% vs Trinity 50.00% (-40.57pp advantage). ", options: { color: C.white, fontSize: 10 } },
    { text: "Trinity has higher quote volume (71 vs 49) but struggles with conversion. ", options: { color: C.amber, fontSize: 10 } },
    { text: "Recommendation: Investigate Trinity's 50% declination pattern to unlock potential.", options: { color: C.midGray, fontSize: 10 } },
  ], { x: 0.45, y: insightY + 0.15, w: 9.1, h: 0.65, fontFace: "Calibri", margin: 0 });
}

console.log("✓ Slide 19: Nexus vs Trinity Comparison");


// ============================================================
// SLIDE 20 — Action Plan & Strategic Commitments (UPDATED)
// ============================================================
{
  const slide = pres.addSlide();
  slide.background = { color: C.bg };

  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.65, fill: { color: C.orange }, line: { color: C.orange } });
  slide.addText("Q1 2026 Action Plan — Winning Patterns & Strategic Commitments", {
    x: 0.3, y: 0.05, w: 9.4, h: 0.55, fontSize: 15, bold: true, color: C.white, fontFace: "Arial", margin: 0
  });

  // LEFT — Critical Alerts
  const alertItems = [
    "2 producers with $0 New Sale require leader intervention.",
    "Office at 61.2% of goal via Combined Premium.",
    "Declination rate varied: 15% → 20% → 24%.",
    "Daniela Salas: $135K in quotes not closed.",
    "Trinity: 50% declination ratio requires investigation.",
  ];
  slide.addShape(pres.shapes.RECTANGLE, { x: 0.15, y: 0.72, w: 3.1, h: 0.3, fill: { color: C.red }, line: { color: C.red } });
  slide.addText("✕ Q4 Critical Alerts", { x: 0.15, y: 0.74, w: 3.1, h: 0.26, fontSize: 9.5, bold: true, color: C.white, align: "center", fontFace: "Arial", margin: 0 });
  slide.addShape(pres.shapes.RECTANGLE, { x: 0.15, y: 1.02, w: 3.1, h: alertItems.length * 0.3 + 0.1, fill: { color: C.cardBg }, line: { color: C.red, width: 1 } });
  alertItems.forEach((item, j) => {
    slide.addText([
      { text: "› ", options: { bold: true, color: C.red } },
      { text: item, options: { color: C.white } }
    ], { x: 0.25, y: 1.06 + j * 0.3, w: 2.9, h: 0.28, fontSize: 7.8, fontFace: "Calibri", margin: 0 });
  });

  // RIGHT — Strategic Commitments (NEW)
  const commitY = 0.72;
  slide.addShape(pres.shapes.RECTANGLE, { x: 3.4, y: commitY, w: 3.1, h: 0.3, fill: { color: C.amber }, line: { color: C.amber } });
  slide.addText("⚡ Strategic Commitments", { x: 3.4, y: commitY + 0.02, w: 3.1, h: 0.26, fontSize: 9.5, bold: true, color: C.textDark, align: "center", fontFace: "Arial", margin: 0 });
  
  const commitItems = [
    "MGA Diversification: Meet goals across ALL MGAs.",
    "Progressive activation: Not selling — needs immediate focus.",
    "Balanced MGA portfolio: Reduce over-reliance on top 2.",
    "Trinity improvement plan: Address 50% declination.",
    "Weekly MGA performance review with leaders.",
  ];
  slide.addShape(pres.shapes.RECTANGLE, { x: 3.4, y: commitY + 0.3, w: 3.1, h: commitItems.length * 0.3 + 0.1, fill: { color: C.cardBg }, line: { color: C.amber, width: 1 } });
  commitItems.forEach((item, j) => {
    slide.addText([
      { text: "▸ ", options: { bold: true, color: C.amber } },
      { text: item, options: { color: C.white } }
    ], { x: 3.5, y: commitY + 0.36 + j * 0.3, w: 2.9, h: 0.28, fontSize: 7.8, fontFace: "Calibri", margin: 0 });
  });

  // Winning Patterns (UPDATED)
  const patternY = 2.6;
  slide.addShape(pres.shapes.RECTANGLE, { x: 0.15, y: patternY, w: 9.7, h: 0.3, fill: { color: C.orangeDark }, line: { color: C.orangeDark } });
  slide.addText("🎯 Winning Patterns — What the Office Must Replicate", {
    x: 0.15, y: patternY + 0.03, w: 9.7, h: 0.26, fontSize: 10, bold: true, color: C.white, align: "center", fontFace: "Arial", margin: 0
  });

  const patterns = [
    { icon: "🚛", title: "Units: 1-3 Units Dominance", detail: "Units 1-3 = 76.78% of revenue. Unit 4 has 87.50% Q2S - model for optimization." },
    { icon: "📦", title: "Business: General Freight Focus", detail: "General Freight: $866K (45.67%). Building Materials: $465K with 53.85% Q2S." },
    { icon: "👶", title: "Years: New Business Engine", detail: "0 Years = 58.79% ($1.12M). Years 0-3 = 87.58%. Strong new business acquisition." },
    { icon: "📍", title: "Miles: 500 + Unlimited", detail: "500 miles + Unlimited = 70.2% of total. Focus on these mileage bands." },
    { icon: "🤝", title: "MGA: Nexus Priority", detail: "Nexus: 46.94% Q2S and 9.43% declination. Significantly outperforms Trinity." },
    { icon: "🏷️", title: "Carrier: Pegaso Excellence", detail: "Pegaso Q2S 79.31%. Year 2: 56% Q2S. Year 7,12: 100%. Mature converts reliably." },
  ];

  const pCols = 3, pRows = 2, pW = 3.15, pH = 1.0, pGap = 0.06;
  patterns.forEach((p, i) => {
    const col = i % pCols;
    const row = Math.floor(i / pCols);
    const px = 0.15 + col * (pW + pGap);
    const py = patternY + 0.35 + row * (pH + 0.06);
    slide.addShape(pres.shapes.RECTANGLE, { x: px, y: py, w: pW, h: pH, fill: { color: C.cardBg }, line: { color: C.orange, width: 1 } });
    slide.addText(`${p.icon} ${p.title}`, { x: px + 0.06, y: py + 0.04, w: pW - 0.12, h: 0.24, fontSize: 7.5, bold: true, color: C.orange, fontFace: "Arial", margin: 0 });
    slide.addText(p.detail, { x: px + 0.06, y: py + 0.28, w: pW - 0.12, h: pH - 0.32, fontSize: 6.8, color: C.midGray, fontFace: "Calibri", margin: 0 });
  });

  // Priority Actions (UPDATED FOR LEADERS NOT PRODUCERS)
  const apY = patternY + 0.35 + 2 * (pH + 0.06) + 0.08;
  slide.addShape(pres.shapes.RECTANGLE, { x: 0.15, y: apY, w: 9.7, h: 0.25, fill: { color: C.orange }, line: { color: C.orange } });
  slide.addText("Q1 2026 — Priority Actions for Sales Leaders", {
    x: 0.15, y: apY + 0.03, w: 9.7, h: 0.2, fontSize: 10, bold: true, color: C.white, align: "center", fontFace: "Arial", margin: 0
  });

  const actions = [
    ["1", "Leaders: Close zero-sale producers pipeline", "Felipe & Daniela intervention"],
    ["2", "Leaders: Activate Progressive placements", "Not selling - immediate action"],
    ["3", "Leaders: Optimize 1-3 Units segment speed", "76.78% of revenue"],
    ["4", "Leaders: Prioritize Nexus + Pegaso on every account", "Top conversion rates"],
    ["5", "Leaders: Focus Truck Tractor + 500mi/Unlimited", "Highest revenue combo"],
    ["6", "Leaders: Investigate Trinity declination pattern", "50% vs 9.43% industry"],
  ];
  const aStartY = apY + 0.27;
  const aRowH = 0.2;
  actions.forEach((a, i) => {
    const col = i < 3 ? 0 : 1;
    const row = i < 3 ? i : i - 3;
    const ax = 0.15 + col * 4.9;
    const ay = aStartY + row * aRowH;
    slide.addShape(pres.shapes.RECTANGLE, { x: ax, y: ay, w: 0.22, h: 0.17, fill: { color: C.orange }, line: { color: C.orange } });
    slide.addText(a[0], { x: ax, y: ay, w: 0.22, h: 0.17, fontSize: 7, bold: true, color: C.white, align: "center", fontFace: "Arial Black", margin: 0 });
    slide.addText(`${a[1]}: `, { x: ax + 0.25, y: ay, w: 2.2, h: 0.17, fontSize: 7, bold: true, color: C.orange, fontFace: "Calibri", margin: 0 });
    slide.addText(a[2], { x: ax + 2.45, y: ay, w: 2.2, h: 0.17, fontSize: 7, color: C.midGray, fontFace: "Calibri", margin: 0 });
  });
}

console.log("✓ Slide 20: Action Plan with Strategic Commitments");

pres.writeFile({
 fileName: `Reporte_${branch}_Q4_2025.pptx`
})
  .then(() => {
    console.log("\n" + "=".repeat(70));
    console.log("✓✓✓ PRESENTATION COMPLETED SUCCESSFULLY! ✓✓✓");
    console.log("=".repeat(70));
    console.log("\nFile: SedeConviccion_Q4_2025_COMPLETE_FINAL.pptx");
    console.log("Total Slides: 20");
    console.log("\nALL CORRECTIONS APPLIED:");
    console.log("  ✓ Slides 1–20 generated");
    console.log("  ✓ Compatible with GitHub Actions");
    console.log("\n" + "=".repeat(70));
  })
  .catch(err => {
    console.error("Error:", err);
    process.exit(1);
  });
