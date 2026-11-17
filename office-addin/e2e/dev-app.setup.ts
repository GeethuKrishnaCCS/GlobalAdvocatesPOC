import { expect, test } from "@playwright/test";
import { readFileSync } from "node:fs";

/**
 * Uploads the DEV copies (from `npm run package:dev`, versions bumped in their manifests) as the E2E account's personal
 * custom apps in Teams on the web. Run: npm run e2e:dev-app.
 * Two apps on purpose: one DEV app with Word/Excel/PowerPoint and Outlook kept vanishing for the account while Teams
 * still listed it (2026-11-05..09); before Outlook was added it was stable. Re-upload also fixes a vanished copy.
 */
const APPS = [
  { name: "Global Advocate DEV", zip: "release/advocate-addin-dev.zip", manifest: "manifest.json" },
  {
    name: "Global Advocate DEV (Outlook)",
    zip: "release/advocate-addin-outlook-dev.zip",
    manifest: "manifest.outlook-dev.json",
  },
];

for (const app of APPS) {
  test(`upload ${app.name} as a personal custom app`, async ({ page }) => {
    test.setTimeout(3 * 60_000);
    const version = (JSON.parse(readFileSync(app.manifest, "utf8")) as { version: string }).version;
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
    await (await chooser).setFiles(app.zip);
    await page
      .getByRole("button", { name: `Add ${app.name}`, exact: true })
      .last()
      .click({ timeout: 60_000 });
    await expect(page.getByText("Added successfully!").first()).toBeVisible({ timeout: 60_000 });
    console.log(`Uploaded ${app.name} ${version}`);
  });
}
