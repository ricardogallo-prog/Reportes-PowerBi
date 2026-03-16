const PptxGenJS = require("pptxgenjs");
const fs = require("fs");

// ------------------------------------
// BRANCH
// ------------------------------------

const branch = process.env.BRANCH_OFFICE
  .normalize("NFD")
  .replace(/[\u0300-\u036f]/g,"")
  .replace(/\s/g,"_");

console.log("Generating report for:", branch);

// ------------------------------------
// LOAD DATA
// ------------------------------------

const data = JSON.parse(
  fs.readFileSync(`report-clean-${branch}.json`)
);

// ------------------------------------
// PRESENTATION
// ------------------------------------

const pres = new PptxGenJS();
pres.layout = "LAYOUT_16x9";

// ------------------------------------
// COLOR PALETTE
// ------------------------------------

const C = {
  bg: "1A1A1A",
  orange: "D4621A",
  orangeLight: "E87B35",
  white: "FFFFFF",
  midGray: "CCCCCC",
  tableRowAlt: "2A2A2A",
  tableRowBase: "222222",
  cardBg: "252525"
};

// ------------------------------------
// COVER SLIDE
// ------------------------------------

{
  const slide = pres.addSlide();

  slide.background = { color: C.bg };

  slide.addText("MARKETSRATER", {
    x:0.5,
    y:1,
    fontSize:36,
    bold:true,
    color:C.orange
  });

  slide.addText("Quarterly Report",{
    x:0.5,
    y:1.8,
    fontSize:22,
    color:C.white
  });

  slide.addText(branch.replace(/_/g," "),{
    x:0.5,
    y:2.5,
    fontSize:18,
    color:C.midGray
  });

  slide.addText(`Generated ${new Date().toLocaleDateString()}`,{
    x:0.5,
    y:3,
    fontSize:14,
    color:C.midGray
  });

}

// ------------------------------------
// FUNCTION TO BUILD TABLE SLIDES
// ------------------------------------

function createTableSlide(title, rows){

  const slide = pres.addSlide();

  slide.background = { color: C.bg };

  slide.addText(title,{
    x:0.5,
    y:0.5,
    fontSize:22,
    bold:true,
    color:C.orange
  });

  const tableRows = rows.map(r => [r.label, r.value]);

  slide.addTable(tableRows,{
    x:0.5,
    y:1.5,
    w:9,
    fontSize:12,
    color:C.white,
    border:{pt:1,color:"444444"}
  });

}

// ------------------------------------
// MAP VISUALS TO SLIDES
// ------------------------------------

for(const visual in data){

  const rows = data[visual];

  if(!rows || rows.length === 0) continue;

  const title = visual
    .replace(/_/g," ")
    .replace("visual","Visual");

  createTableSlide(title, rows);

}

// ------------------------------------
// SUMMARY SLIDE
// ------------------------------------

{
  const slide = pres.addSlide();

  slide.background = { color: C.bg };

  slide.addText("Report Summary",{
    x:0.5,
    y:1,
    fontSize:24,
    bold:true,
    color:C.orange
  });

  slide.addText(`Branch: ${branch.replace(/_/g," ")}`,{
    x:0.5,
    y:2,
    fontSize:16,
    color:C.white
  });

  slide.addText(`Visuals analyzed: ${Object.keys(data).length}`,{
    x:0.5,
    y:2.6,
    fontSize:16,
    color:C.white
  });

}

// ------------------------------------
// SAVE TEMP PPTX
// ------------------------------------

(async () => {

  await pres.writeFile({
    fileName: `temp_${branch}.pptx`
  });

  console.log("Presentation generated");

})();
