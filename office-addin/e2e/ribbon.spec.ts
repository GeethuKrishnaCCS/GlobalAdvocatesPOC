import { expect, test, type Response } from "@playwright/test";
import { captureConsole, docUrlFor, findRibbonFrame, saveRunScreenshot, sideloadUrl } from "./sideload";

const ICON = /\/assets\/icon-\d+\.png/;

test.afterEach(async ({ page }, testInfo) => {
  if (testInfo.status !== "skipped") await saveRunScreenshot(page, testInfo);
});

// Ribbon icons load cross-site from the dev server (no-cors <img>); webpack-dev-server's cross-origin check once 403'd them.
test("word: Global Advocate tab shows its icons", async ({ page }, testInfo) => {
  const docUrl = docUrlFor("word");
  test.skip(!docUrl, "no test document configured for word (e2e/.env)");
  captureConsole(page, testInfo);
  const icons: Response[] = [];
  page.on("response", (r) => ICON.test(r.url()) && icons.push(r));

  await page.goto(sideloadUrl(docUrl!));
  const ribbon = await findRibbonFrame(page, "Global Advocate");
  await ribbon.getByRole("tab", { name: "Global Advocate" }).first().click();
  await expect(ribbon.getByRole("button", { name: "Save to Matter" }).first()).toBeVisible();

  await expect.poll(() => icons.length, { message: "no ribbon icon requested", timeout: 15_000 }).toBeGreaterThan(0);
  expect(icons.map((r) => `${r.status()} ${r.url()}`).filter((s) => !s.startsWith("200"))).toEqual([]);
});
