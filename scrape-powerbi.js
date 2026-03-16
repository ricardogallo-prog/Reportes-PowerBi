const fs = require("fs");
const { chromium } = require("playwright");

(async () => {

  const baseUrl = process.env.POWERBI_URL;
  const branch = process.env.BRANCH_OFFICE;

  const filter = `Sales/branch_office eq '${branch}'`;
  const url = `${baseUrl}&filter=${encodeURIComponent(filter)}`;

  console.log("Opening:", url);

  const browser = await chromium.launch({ headless: true });

  const page = await browser.newPage({
    viewport: { width: 1920, height: 1080 }
  });

  await page.goto(url, { waitUntil: "networkidle" });

  await page.waitForTimeout(15000);

  await page.screenshot({
    path: `dashboard_${branch}.png`,
    fullPage: true
  });

  const data = {
    branch: branch,
    date: new Date().toISOString(),
    screenshot: `dashboard_${branch}.png`
  };

  fs.writeFileSync(
    `report-data-${branch}.json`,
    JSON.stringify(data, null, 2)
  );

  await browser.close();

})();
  fs.writeFileSync("report-data.json", JSON.stringify(reportData, null, 2), "utf8");

  await browser.close();
})();
