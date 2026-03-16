const fs = require("fs");
const { chromium } = require("playwright");

(async () => {

  const baseUrl = process.env.POWERBI_URL;
  const branch = process.env.BRANCH_OFFICE;

  const filter = `markets_01/Branch_Office eq '${branch}'`;
  const url = `${baseUrl}&filter=${encodeURIComponent(filter)}`;

  const browser = await chromium.launch({ headless: true });

  const page = await browser.newPage({
    viewport: { width: 1920, height: 1080 }
  });

  console.log("Opening dashboard:", branch);

  await page.goto(url, { waitUntil: "networkidle" });

  await page.waitForTimeout(15000);

  const visuals = page.locator(".visual-container");

  const count = await visuals.count();

  console.log("Visuals found:", count);

  const extracted = {};

  for (let i = 0; i < count; i++) {

    const visual = visuals.nth(i);

    try {

      await visual.scrollIntoViewIfNeeded();

      await visual.hover();

      const more = visual.locator("button[aria-label='More options']");

      if (await more.count() === 0) continue;

      await more.first().click();

      await page.waitForTimeout(1500);

      const showTable = page.locator("text=Show as table");

      if (await showTable.count() > 0) {

        await showTable.first().click();

        await page.waitForTimeout(4000);

        const rows = await page.locator("table tr").allTextContents();

        extracted[`visual_${i}`] = rows;

        await page.keyboard.press("Escape");

      }

    } catch (err) {

      console.log("visual skipped:", i);

    }

  }

  const screenshot = `dashboard_${branch.replace(/\s/g,"_")}.png`;

  await page.screenshot({
    path: screenshot,
    fullPage: true
  });

  const data = {
    branch,
    generated: new Date().toISOString(),
    visuals: extracted,
    screenshot
  };

  fs.writeFileSync(
    `report-data-${branch.replace(/\s/g,"_")}.json`,
    JSON.stringify(data, null, 2)
  );

  await browser.close();

})();
