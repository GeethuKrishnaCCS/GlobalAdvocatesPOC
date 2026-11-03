import { expect, test, type Frame } from "@playwright/test";
import { captureConsole, docUrlFor, findRibbonFrame, saveRunScreenshot, signIn } from "./sideload";

/* global process, console */

/**
 * Centrally deployed (production) add-in, not the dev sideload. Opt-in: set ADV_E2E_HOSTED_ORIGIN to the hosting
 * origin (e.g. https://stadvtestaddin01.z29.web.core.windows.net); the signed-in test user must be assigned the add-in.
 * Word on the web keeps a user's URL-param sideloads (Developer add-ins), and dev and prod share the ribbon labels,
 * so every "Global Advocate" tab is tried and only a pane served from the hosting origin counts.
 */
const HOSTED_ORIGIN = process.env.ADV_E2E_HOSTED_ORIGIN;
const TAB = "Global Advocate";
const PANE_WAIT_MS = 20_000;

test.afterEach(async ({ page }, testInfo) => {
  if (testInfo.status !== "skipped") await saveRunScreenshot(page, testInfo);
});

test("word: centrally deployed add-in opens from hosting and signs in", async ({ page }, testInfo) => {
  const docUrl = docUrlFor("word");
  test.skip(!docUrl || !HOSTED_ORIGIN, "set ADV_E2E_HOSTED_ORIGIN and ADV_E2E_DOC_WORD to check the deployed add-in");
  test.setTimeout(4 * 60_000);
  captureConsole(page, testInfo);

  await page.goto(docUrl!);
  const ribbon = await findRibbonFrame(page, TAB);
  const tabs = ribbon.getByRole("tab", { name: TAB });
  const tabCount = await tabs.count();

  const hostedPane = (): Frame | undefined => page.frames().find((f) => f.url().startsWith(HOSTED_ORIGIN!));
  for (let i = 0; i < tabCount && !hostedPane(); i++) {
    await tabs.nth(i).click();
    await ribbon.getByRole("button", { name: "Save to Matter" }).filter({ visible: true }).first().click();
    await expect
      .poll(() => !!hostedPane(), { timeout: PANE_WAIT_MS })
      .toBe(true)
      .catch(() => undefined); // not this tab's add-in (or not loaded yet); try the next tab
  }

  const pane = hostedPane();
  const origins = [...new Set(page.frames().map((f) => new URL(f.url()).origin))].filter(
    (o) => /localhost|web\.core\.windows\.net/.test(o)
  );
  expect(
    pane,
    `No pane from ${HOSTED_ORIGIN} in ${tabCount} "${TAB}" tab(s); add-in panes seen: ${origins.join(", ") || "none"}. ` +
      "If only localhost (dev) shows, the deployed add-in is not assigned to this user yet or not propagated (up to 24 h)."
  ).toBeDefined();

  await expect(pane!.getByTestId("app-title")).toHaveText("Save to Matter");
  console.log(`[hosted] ${await pane!.getByTestId("app-host").textContent()}`);
  await signIn(page, pane!);
  await expect(pane!.getByTestId("app-user")).toContainText("@");
  await expect(pane!.getByTestId("matter-step")).toBeVisible({ timeout: 60_000 });
});
