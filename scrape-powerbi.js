const fs = require("fs");
const { chromium } = require("playwright");

(async () => {

  const baseUrl = process.env.POWERBI_URL;
  const branch = process.env.BRANCH_OFFICE;

  const filter = `markets_01/Branch_Office eq '${branch}'`;

  const url = `${baseUrl}&filter=${encodeURIComponent(filter)}`;

  console.log("Opening dashboard for:", branch);
  console.log(url);

  const browser = await chromium.launch({ headless: true });

  const page = await browser.newPage({
    viewport: { width: 1920, height: 1080 }
  });

  await page.goto(url, { waitUntil: "networkidle" });

  await page.waitForTimeout(15000);

  const screenshotName = `dashboard_${branch.replace(/\s/g,"_")}.png`;

  await page.screenshot({
    path: screenshotName,
    fullPage: true
  });

  const data = {
    branch: branch,
    screenshot: screenshotName,
    generated: new Date().toISOString()
  };

  fs.writeFileSync(
    `report-data-${branch.replace(/\s/g,"_")}.json`,
    JSON.stringify(data, null, 2)
  );

  await browser.close();

})();
