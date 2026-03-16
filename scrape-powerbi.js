const fs = require("fs");
const { chromium } = require("playwright");

(async () => {

  const baseUrl = process.env.POWERBI_URL;
  const branch = process.env.BRANCH_OFFICE;

  const branchSafe = branch
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s/g, "_");

  const browser = await chromium.launch({ headless: true });

  const page = await browser.newPage({
    viewport: { width: 1920, height: 1080 }
  });

  console.log("Opening dashboard for:", branch);

  await page.goto(baseUrl, { waitUntil: "networkidle" });

  await page.waitForTimeout(15000);

  // -------------------------------
  // APPLY BRANCH FILTER (SLICER)
  // -------------------------------

  try {

    console.log("Applying Branch_Office filter...");

    const slicer = page.locator("text=Branch_Office").first();

    await slicer.click();

    await page.waitForTimeout(2000);

    const option = page.locator(`text=${branch}`).first();

    await option.click();

    console.log("Branch selected:", branch);

    await page.waitForTimeout(8000);

  } catch (err) {

    console.log("Branch filter not applied:", err.message);

  }

  // -------------------------------
  // SCRAPE VISUALS
  // -------------------------------

  const visuals = page.locator(".visual-container");

  const count = await visuals.count();

  console.log("Visuals detected:", count);

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

        console.log("Extracted visual:", i);

      }

    } catch (err) {

      console.log("Visual skipped:", i);

    }

  }

  // -------------------------------
  // SCREENSHOT
  // -------------------------------

  const screenshot = `dashboard_${branchSafe}.png`;

  await page.screenshot({
    path: screenshot,
    fullPage: true
  });

  // -------------------------------
  // SAVE DATA
  // -------------------------------

  const data = {
    branch,
    generated: new Date().toISOString(),
    visuals: extracted,
    screenshot
  };

  fs.writeFileSync(
    `report-data-${branchSafe}.json`,
    JSON.stringify(data, null, 2)
  );

  console.log("Data saved:", `report-data-${branchSafe}.json`);

  await browser.close();

})();
