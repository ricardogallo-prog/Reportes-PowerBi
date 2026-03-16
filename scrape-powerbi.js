const fs = require("fs");
const { chromium } = require("playwright");

(async () => {
  const url = process.env.POWERBI_URL;

  if (!url) {
    throw new Error("Falta la variable POWERBI_URL");
  }

  const browser = await chromium.launch({ headless: true });
  const page = await browser.newPage({ viewport: { width: 1600, height: 900 } });

  await page.goto(url, { waitUntil: "domcontentloaded", timeout: 120000 });
  await page.waitForTimeout(15000);

  await page.screenshot({ path: "dashboard.png", fullPage: true });

  const reportData = {
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

  fs.writeFileSync("report-data.json", JSON.stringify(reportData, null, 2), "utf8");

  await browser.close();
})();
