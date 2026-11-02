import { test } from "@playwright/test";
import { STORAGE_STATE } from "../playwright.config";

/** User signs in by hand (dev06@ccsdev01, MFA if asked); the session is saved for headless runs. */
const SIGN_IN_WAIT_MS = 5 * 60_000;

test("sign in to Microsoft 365 and save session", async ({ page }) => {
  test.setTimeout(SIGN_IN_WAIT_MS + 30_000);
  await page.goto("https://www.office.com/login");
  // Signed in once the browser leaves login.microsoftonline.com for the Microsoft 365 home page.
  // `/login` itself is on office.com, so the path check stops the wait resolving before the redirect.
  await page.waitForURL(
    (url) =>
      /(m365\.cloud\.microsoft|office\.com)$/.test(url.host) && !url.pathname.toLowerCase().startsWith("/login"),
    { timeout: SIGN_IN_WAIT_MS }
  );
  await page.waitForLoadState("networkidle");
  await page.context().storageState({ path: STORAGE_STATE });
});
