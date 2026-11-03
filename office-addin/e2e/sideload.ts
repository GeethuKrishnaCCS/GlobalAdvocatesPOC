import { expect, type Frame, type Page, type TestInfo } from "@playwright/test";
import { readFileSync, appendFileSync } from "node:fs";
import { join } from "node:path";

/* global process */

export type OfficeWebHost = "word" | "excel" | "powerpoint";

const DEV_SERVER_ORIGIN = "https://localhost:3000";
const MANIFEST_FILE = "manifest.xml";
const DOC_URL_ENV: Record<OfficeWebHost, string> = {
  word: "ADV_E2E_DOC_WORD",
  excel: "ADV_E2E_DOC_EXCEL",
  powerpoint: "ADV_E2E_DOC_POWERPOINT",
};

/** Test document URL for a host from e2e/.env, or undefined when not configured. */
export function docUrlFor(host: OfficeWebHost): string | undefined {
  return process.env[DOC_URL_ENV[host]] || undefined;
}

/**
 * Same URL-param dev sideload office-addin-debugging uses for Office on the web
 * (office-addin-dev-settings generateSideloadUrl): no catalog upload, loads instantly, this session only.
 * `wdaddintest=true` suppresses Office's add-in trust dialogs.
 */
export function sideloadUrl(docUrl: string): string {
  const id = /<Id>([^<]+)<\/Id>/.exec(readFileSync(MANIFEST_FILE, "utf8"))?.[1];
  if (!id) throw new Error(`No <Id> in ${MANIFEST_FILE}`);
  const port = new URL(DEV_SERVER_ORIGIN).port;
  return `${docUrl}&wdaddindevserverport=${port}&wdaddinmanifestfile=${MANIFEST_FILE}&wdaddinmanifestguid=${id}${process.env.ADV_E2E_NO_TESTMODE ? "" : "&wdaddintest=true"}`;
}

/**
 * Full-page screenshot of every run (pass or fail) into ADV_E2E_SCREENSHOT_DIR, kept across runs
 * (outputDir is wiped each run). No-op when unset.
 */
export async function saveRunScreenshot(page: Page, testInfo: TestInfo): Promise<void> {
  const dir = process.env.ADV_E2E_SCREENSHOT_DIR;
  if (!dir) return;
  const stamp = new Date().toISOString().replace(/[-:]/g, "").replace("T", "-").slice(0, 15);
  const name = testInfo.title.replace(/[^\w-]+/g, "-").replace(/^-+|-+$/g, "");
  await page.screenshot({ path: join(dir, `${stamp}-${name}-${testInfo.status}.png`), fullPage: true });
}

/** Pipes page + iframe console output and page errors to <test output>/console.log (grep it; never dump DOM). */
export function captureConsole(page: Page, testInfo: TestInfo): void {
  const file = testInfo.outputPath("console.log");
  const write = (line: string): void => appendFileSync(file, `${new Date().toISOString()} ${line}\n`);
  const attach = (p: Page, tag: string): void => {
    p.on("console", (m) => write(`[${m.type()}] ${tag}${m.location().url?.startsWith(DEV_SERVER_ORIGIN) ? "PANE " : ""}${m.text()}`));
    p.on("pageerror", (e) => write(`[pageerror] ${tag}${e.message}`));
    p.on("requestfailed", (r) => write(`[requestfailed] ${tag}${r.url().slice(0, 200)} ${r.failure()?.errorText}`));
    p.on("response", (r) => r.status() >= 400 && write(`[http ${r.status()}] ${tag}${r.url().slice(0, 200)}`));
  };
  attach(page, "");
  // The Office sign-in dialog (and the IdP pages it visits) is a separate page.
  page.context().on("page", (p) => {
    write(`[popup] opened ${p.url().slice(0, 120)}`);
    attach(p, "POPUP ");
  });
}

/** Office on the web asks once per browser profile before URL-param sideload works; the modal hides the ribbon. */
async function acceptDeveloperMode(frame: Frame): Promise<void> {
  const dialog = frame.getByRole("dialog", { name: "Enable Developer Mode" });
  if (await dialog.isVisible().catch(() => false)) {
    await dialog.getByRole("button", { name: "OK" }).click();
  }
}

/** First frame (any depth) where the ribbon tab is visible; Office on the web hosts the ribbon in a nested frame. */
export async function findRibbonFrame(page: Page, tabName: string): Promise<Frame> {
  let found: Frame | undefined;
  await expect
    .poll(
      async () => {
        for (const frame of page.frames()) {
          await acceptDeveloperMode(frame);
          if (await frame.getByRole("tab", { name: tabName }).first().isVisible().catch(() => false)) {
            found = frame;
            return true;
          }
        }
        return false;
      },
      { message: `ribbon tab "${tabName}" not visible in any frame`, timeout: 90_000 }
    )
    .toBe(true);
  return found!;
}

/** Task pane frame served by the dev server, once it has loaded. */
export async function findPaneFrame(page: Page): Promise<Frame> {
  let pane: Frame | undefined;
  await expect
    .poll(() => (pane = page.frames().find((f) => f.url().startsWith(DEV_SERVER_ORIGIN))) !== undefined, {
      message: "task pane frame from localhost:3000 never loaded",
      timeout: 60_000,
    })
    .toBe(true);
  return pane!;
}

/** Opens the doc with the add-in sideloaded, clicks Global Advocate › <button>, returns the pane frame. */
export async function openPane(page: Page, docUrl: string, buttonName: string): Promise<Frame> {
  await page.goto(sideloadUrl(docUrl));
  const ribbon = await findRibbonFrame(page, "Global Advocate");
  await ribbon.getByRole("tab", { name: "Global Advocate" }).first().click();
  await ribbon.getByRole("button", { name: buttonName }).first().click();
  return findPaneFrame(page);
}

/**
 * Signs the pane in. NAA works silently where Office provides its bridge; the web dev sideload gets none, so the pane
 * opens the Office sign-in dialog, which completes from the saved session and closes itself.
 */
export async function signIn(page: Page, pane: Frame): Promise<void> {
  const user = pane.getByTestId("app-user");
  const button = pane.getByTestId("app-signin-button");
  await expect(user.or(button)).toBeVisible({ timeout: 60_000 });
  if (await user.isVisible()) return;

  const dialogPromise = page.context().waitForEvent("page");
  await button.click();
  const dialog = await dialogPromise;
  await dialog.waitForEvent("close", { timeout: 60_000 });
  await expect(user).toBeVisible({ timeout: 30_000 });
}
