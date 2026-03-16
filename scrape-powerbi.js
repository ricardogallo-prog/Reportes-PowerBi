const fs = require("fs");
const { chromium } = require("playwright");

(async () => {

  const baseUrl = process.env.POWERBI_URL;
  const branch = process.env.BRANCH_OFFICE;

  const filter = `markets_01/Branch_Office eq '${branch}'`;
  const url = `${baseUrl}&filter=${encodeURIComponent(filter)}`;

  console.log("Opening:", branch);

  const browser = await chromium.launch({ headless: true });

  const page = await browser.newPage({
    viewport: { width: 1920, height: 1080 }
  });

  await page.goto(url, { waitUntil: "networkidle" });

  await page.waitForTimeout(15000);

  // =============================
  // abrir visual Premium per Producer
  // =============================

  const visual = page.locator("text=Premium per Producer");

  await visual.click();

  await page.waitForTimeout(2000);

  // abrir menú del visual
  await page.locator("button[aria-label='More options']").first().click();

  await page.waitForTimeout(2000);

  await page.locator("text=Show as table").click();

  await page.waitForTimeout(4000);

  const table = await page.locator("table").innerText();

  const screenshotName = `dashboard_${branch.replace(/\s/g,"_")}.png`;

  await page.screenshot({
    path: screenshotName,
    fullPage: true
  });

  const data = {
    branch,
    premiumPerProducerTable: table,
    screenshot: screenshotName
  };

  fs.writeFileSync(
    `report-data-${branch.replace(/\s/g,"_")}.json`,
    JSON.stringify(data, null, 2)
  );

  await browser.close();

})();
