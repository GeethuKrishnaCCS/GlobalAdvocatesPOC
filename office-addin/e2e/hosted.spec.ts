import { expect, test } from "@playwright/test";
import {
  captureConsole,
  docUrlFor,
  hostedOrigin,
  openPaneFrom,
  saveRunScreenshot,
  signIn,
  type OfficeWebHost,
} from "./sideload";

/* global console */

/**
 * Centrally deployed (production) add-in, not the dev sideload: the document opens without sideload parameters and only
 * a pane served from ADV_E2E_HOSTED_ORIGIN counts. Opt-in (skipped without the variable); the signed-in test user must be
 * assigned the add-in.
 */
test.afterEach(async ({ page }, testInfo) => {
  if (testInfo.status !== "skipped") await saveRunScreenshot(page, testInfo);
});

for (const host of ["word", "excel", "powerpoint"] as OfficeWebHost[]) {
  test(`${host}: centrally deployed add-in opens from hosting and signs in`, async ({ page }, testInfo) => {
    const docUrl = docUrlFor(host);
    const origin = hostedOrigin();
    test.skip(!docUrl || !origin, `set ADV_E2E_HOSTED_ORIGIN and a ${host} test document to check the deployed add-in`);
    test.setTimeout(4 * 60_000);
    captureConsole(page, testInfo);

    await page.goto(docUrl!);
    const pane = await openPaneFrom(page, "Save to Matter", origin!);
    await expect(pane.getByTestId("app-title")).toHaveText("Save to Matter");
    console.log(`[hosted] ${await pane.getByTestId("app-host").textContent()}`);
    await signIn(page, pane);
    await expect(pane.getByTestId("app-user")).toContainText("@");
    await expect(pane.getByTestId("matter-step")).toBeVisible({ timeout: 60_000 });
  });
}
