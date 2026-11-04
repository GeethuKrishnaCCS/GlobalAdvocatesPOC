import { expect, test, type Response } from "@playwright/test";
import { DEV_TAB, captureConsole, docUrlFor, findRibbonFrame, saveRunScreenshot } from "./sideload";

const ICON = /\/assets\/icon-\d+\.png/;

test.afterEach(async ({ page }, testInfo) => {
  if (testInfo.status !== "skipped") await saveRunScreenshot(page, testInfo);
});

// Ribbon icons load cross-site from the dev server (no-cors <img>); webpack-dev-server's cross-origin check once 403'd them.
test("word: Global Advocate DEV tab shows its icons", async ({ page }, testInfo) => {
  const docUrl = docUrlFor("word");
  test.skip(!docUrl, "no test document configured for word (e2e/.env)");
  captureConsole(page, testInfo);
  const icons: Response[] = [];
  page.on("response", (r) => ICON.test(r.url()) && icons.push(r));

  await page.goto(docUrl!);
  const ribbon = await findRibbonFrame(page, DEV_TAB);
  await ribbon.getByRole("tab", { name: DEV_TAB, exact: true }).click();
  await expect(ribbon.getByRole("button", { name: "Save to Matter" }).first()).toBeVisible();

  await expect.poll(() => icons.length, { message: "no ribbon icon requested", timeout: 15_000 }).toBeGreaterThan(0);
  expect(icons.map((r) => `${r.status()} ${r.url()}`).filter((s) => !s.startsWith("200"))).toEqual([]);
});
