import { expect, test } from "@playwright/test";
import { readFileSync } from "node:fs";

/**
 * Uploads the DEV copy (release/advocate-addin-dev.zip from `npm run package:dev`, manifest.json with a bumped
 * version) as the E2E account's personal custom app in Teams on the web. Run: npm run e2e:dev-app.
 * Needed when the "Global Advocate DEV" tab / Outlook tile disappears for the E2E account (seen twice on 2026-11-05
 * while Teams still listed the app) or after manifest.json changes.
 */
const ZIP = "release/advocate-addin-dev.zip";

test("upload the DEV copy as a personal custom app", async ({ page }) => {
  test.setTimeout(3 * 60_000);
  const version = (JSON.parse(readFileSync("manifest.json", "utf8")) as { version: string })
    .version;
  await page.setViewportSize({ width: 1400, height: 900 });
  await page.goto("https://teams.microsoft.com/v2/");
  // shortcut: Teams rail/store icons have no stable names; positions hold at 1400x900. Upgrade if Teams moves them.
  await page.waitForTimeout(20_000);
  await page.mouse.click(34, 588); // rail: Apps
  await page.waitForTimeout(8_000);
  await page.mouse.click(1368, 114); // store: settings → Manage your apps
  await page.getByText("Upload an app").click({ timeout: 60_000 });
  const chooser = page.waitForEvent("filechooser");
  await page
    .getByText(/Upload a custom app/i)
    .first()
    .click();
  await (await chooser).setFiles(ZIP);
  await page
    .getByRole("button", { name: "Add Global Advocate DEV" })
    .last()
    .click({ timeout: 60_000 });
  await expect(page.getByText("Added successfully!").first()).toBeVisible({ timeout: 60_000 });
  console.log(`Uploaded Global Advocate DEV ${version}`);
});
