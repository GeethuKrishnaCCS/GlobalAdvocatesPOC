import { expect, test, type Frame, type Page, type TestInfo } from "@playwright/test";
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


const TAB = "Global Advocate";
/** How long one tab's pane gets to appear before the next "Global Advocate" tab is tried. */
const PANE_WAIT_MS = 20_000;

/** Hosting origin of the centrally deployed add-in (e2e/.env or the shell); unset = hosted checks are skipped. */
export function hostedOrigin(): string | undefined {
  return process.env.ADV_E2E_HOSTED_ORIGIN || undefined;
}

/**
 * Clicks Global Advocate › <button> and returns the pane served from `origin`. A user can have the dev sideload and the
 * centrally deployed add-in at once, with identical ribbon labels, so each "Global Advocate" tab is tried in turn.
 */
export async function openPaneFrom(page: Page, buttonName: string, origin: string): Promise<Frame> {
  const ribbon = await findRibbonFrame(page, TAB);
  const tabs = ribbon.getByRole("tab", { name: TAB });
  const count = await tabs.count();
  const pane = (): Frame | undefined => page.frames().find((f) => f.url().startsWith(origin));
  for (let i = 0; i < count && !pane(); i++) {
    await tabs.nth(i).click();
    await ribbon.getByRole("button", { name: buttonName }).filter({ visible: true }).first().click();
    await expect
      .poll(() => !!pane(), { timeout: PANE_WAIT_MS })
      .toBe(true)
      .catch(() => undefined); // another add-in's tab (or not loaded yet): try the next one
  }
  const seen = [...new Set(page.frames().map((f) => f.url().split("/").slice(0, 3).join("/")))].filter((o) =>
    /localhost|web.core.windows.net/.test(o)
  );
  expect(
    pane(),
    `No task pane from ${origin} in ${count} "${TAB}" tab(s); add-in panes seen: ${seen.join(", ") || "none"}. ` +
      "Hosted: the add-in may not be assigned to this user or not propagated yet (up to 24 h)."
  ).toBeDefined();
  return pane()!;
}

/** Opens the doc with the add-in sideloaded from the dev server, clicks Global Advocate › <button>, returns the pane. */
export async function openPane(page: Page, docUrl: string, buttonName: string): Promise<Frame> {
  await page.goto(sideloadUrl(docUrl));
  return openPaneFrom(page, buttonName, DEV_SERVER_ORIGIN);
}

/**
 * Pane for a host: Word from the dev server (URL sideload); Excel and PowerPoint only load the add-in through
 * Centralized Deployment on the web, so they use the hosted add-in, or are skipped when no hosting origin is set.
 */
export async function openPaneFor(page: Page, host: OfficeWebHost, docUrl: string, buttonName: string): Promise<Frame> {
  if (host === "word") return openPane(page, docUrl, buttonName);
  const origin = hostedOrigin();
  test.skip(!origin, `${host} on the web only loads the deployed add-in: set ADV_E2E_HOSTED_ORIGIN`);
  await page.goto(docUrl);
  return openPaneFrom(page, buttonName, origin!);
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

/**
 * Calls Graph from inside the pane with the pane's own (dialog) token: the test checks what the signed-in user
 * created, with exactly the user's delegated access. Returns the parsed JSON body (undefined for 204).
 */
export async function graphFromPane<T>(pane: Frame, method: string, path: string): Promise<T> {
  return pane.evaluate(
    async ({ method, path }) => {
      const token = JSON.parse(localStorage.getItem("adv.dialogToken") ?? "{}").accessToken;
      const response = await fetch(`https://graph.microsoft.com/v1.0${path}`, {
        method,
        headers: { Authorization: `Bearer ${token}` },
      });
      if (!response.ok) throw new Error(`Graph ${method} ${path}: HTTP ${response.status}`);
      return response.status === 204 ? undefined : response.json();
    },
    { method, path }
  ) as Promise<T>;
}
