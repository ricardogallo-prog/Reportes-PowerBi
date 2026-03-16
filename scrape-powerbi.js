const fs = require("fs");
const { chromium } = require("playwright");

(async () => {
  const baseUrl = process.env.POWERBI_URL;
  const branch = process.env.BRANCH_OFFICE;

  // Ajusta la tabla/campo si en tu modelo se llaman distinto
  const filter = `markets_01/Branch_Office eq '${branch}'`;
  const url = `${baseUrl}&filter=${encodeURIComponent(filter)}`;

  const branchSafe = branch.replace(/\s/g, "_");

  console.log("Opening dashboard for:", branch);
  console.log(url);

  const browser = await chromium.launch({ headless: true });
  const page = await browser.newPage({ viewport: { width: 1920, height: 1080 } });

  await page.goto(url, { waitUntil: "networkidle" });
  await page.waitForTimeout(15000);

  const visuals = [
    "Premium per Unit Type",
    "Monthly sold premium",
    "Leads per Sales Leader",
    "Premium per Producer",
    "Premium per Carrier",
    "Premium per MGA",
    "Premium per Type of Business",
    "Premium per # of Units",
    "Premium per Years in Business"
  ];

  const extracted = {};

  for (const title of visuals) {
    console.log("Processing visual:", title);

    try {
      const visual = page.locator(`text=${title}`).first();

      await visual.scrollIntoViewIfNeeded();
      await visual.click();

      await page.waitForTimeout(2000);

      // abrir menú del visual
      const moreOptions = page.locator("button[aria-label='More options']").first();
      await moreOptions.click();

      await page.waitForTimeout(2000);

      const showTable = page.locator("text=Show as table");

      if (await showTable.count() > 0) {
        await showTable.click();

        await page.waitForTimeout(4000);

        const rows = await page.locator("table tr").allTextContents();

        extracted[title] = rows;

        // cerrar tabla
        await page.keyboard.press("Escape");
      } else {
        console.log("No table option for:", title);
        extracted[title] = "No table view available";
      }

    } catch (err) {
      console.log("Error with visual:", title);
      console.log(err.message);
      extracted[title] = "Extraction failed";
    }
  }

  const screenshotName = `dashboard_${branchSafe}.png`;

  await page.screenshot({
    path: screenshotName,
    fullPage: true
  });

  const data = {
    branch: branch,
    generated: new Date().toISOString(),
    visuals: extracted,
    screenshot: screenshotName
  };

  fs.writeFileSync(
    `report-data-${branchSafe}.json`,
    JSON.stringify(data, null, 2)
  );

  await browser.close();
})();
