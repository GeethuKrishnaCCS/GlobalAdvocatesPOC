import { expect, test, type Frame, type Page, type TestInfo } from "@playwright/test";
import { appendFileSync } from "node:fs";
import { join } from "node:path";

/* global process */

export type OfficeWebHost = "word" | "excel" | "powerpoint";

const DEV_SERVER_ORIGIN = "https://localhost:3000";
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
 * Full-page screenshot of every run (pass or fail) into ADV_E2E_SCREENSHOT_DIR, kept across runs
 * (outputDir is wiped each run). No-op when unset.
 */
export async function saveRunScreenshot(page: Page, testInfo: TestInfo): Promise<void> {
  const dir = process.env.ADV_E2E_SCREENSHOT_DIR;
  if (!dir) return;
  const stamp = new Date().toISOString().replace(/[-:]/g, "").replace("T", "-").slice(0, 15);
  const name = testInfo.title.replace(/[^\w-]+/g, "-").replace(/^-+|-+$/g, "");
  await page.screenshot({
    path: join(dir, `${stamp}-${name}-${testInfo.status}.png`),
    fullPage: true,
  });
}

/** Pipes page + iframe console output and page errors to <test output>/console.log (grep it; never dump DOM). */
export function captureConsole(page: Page, testInfo: TestInfo): void {
  const file = testInfo.outputPath("console.log");
  const write = (line: string): void =>
    appendFileSync(file, `${new Date().toISOString()} ${line}\n`);
  const attach = (p: Page, tag: string): void => {
    p.on("console", (m) =>
      write(
        `[${m.type()}] ${tag}${m.location().url?.startsWith(DEV_SERVER_ORIGIN) ? "PANE " : ""}${m.text()}`
      )
    );
    p.on("pageerror", (e) => write(`[pageerror] ${tag}${e.message}`));
    p.on("requestfailed", (r) =>
      write(`[requestfailed] ${tag}${r.url().slice(0, 200)} ${r.failure()?.errorText}`)
    );
    p.on(
      "response",
      (r) => r.status() >= 400 && write(`[http ${r.status()}] ${tag}${r.url().slice(0, 200)}`)
    );
  };
  attach(page, "");
  // The Office sign-in dialog (and the IdP pages it visits) is a separate page.
  page.context().on("page", (p) => {
    write(`[popup] opened ${p.url().slice(0, 120)}`);
    attach(p, "POPUP ");
  });
}

/** First frame (any depth) where the ribbon tab is visible; Office on the web hosts the ribbon in a nested frame. */
export async function findRibbonFrame(page: Page, tabName: string): Promise<Frame> {
  let found: Frame | undefined;
  await expect
    .poll(
      async () => {
        for (const frame of page.frames()) {
          if (
            await frame
              .getByRole("tab", { name: tabName })
              .first()
              .isVisible()
              .catch(() => false)
          ) {
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

/** Ribbon tab of the deployed add-in (users) and of the DEV copy loaded from the dev server (E2E account only:
 * personal upload of release/advocate-addin-dev.zip, see README). */
export const PROD_TAB = "Global Advocate";
export const DEV_TAB = "Global Advocate DEV";
/** How long the pane gets to appear after the ribbon button is clicked. */
const PANE_WAIT_MS = 20_000;

/** Hosting origin of the centrally deployed add-in (e2e/.env or the shell); unset = hosted checks are skipped. */
export function hostedOrigin(): string | undefined {
  return process.env.ADV_E2E_HOSTED_ORIGIN || undefined;
}

/**
 * Clicks the add-in's tab › <button> and returns the pane served from `origin`: the DEV copy's tab for the dev server,
 * the deployed add-in's tab otherwise (the E2E account has both; their labels differ).
 */
export async function openPaneFrom(page: Page, buttonName: string, origin: string): Promise<Frame> {
  const tabName = origin === DEV_SERVER_ORIGIN ? DEV_TAB : PROD_TAB;
  const ribbon = await findRibbonFrame(page, tabName);
  await ribbon.getByRole("tab", { name: tabName, exact: true }).click();
  await ribbon.getByRole("button", { name: buttonName }).filter({ visible: true }).first().click();
  const pane = (): Frame | undefined => page.frames().find((f) => f.url().startsWith(origin));
  await expect
    .poll(() => !!pane(), {
      message:
        `No task pane from ${origin} after "${tabName}" › ${buttonName}. Hosted: the add-in may not be assigned to ` +
        "this user or not propagated yet (up to 24 h).",
      timeout: PANE_WAIT_MS,
    })
    .toBe(true);
  return pane()!;
}

/** Opens the doc, clicks the DEV copy's tab › <button>, returns the pane served by the dev server. */
export async function openPane(page: Page, docUrl: string, buttonName: string): Promise<Frame> {
  await page.goto(docUrl);
  return openPaneFrom(page, buttonName, DEV_SERVER_ORIGIN);
}

/**
 * Pane for a host: the DEV copy from the dev server (E2E account), or with ADV_E2E_NO_DEV_SERVER the deployed
 * add-in from ADV_E2E_HOSTED_ORIGIN (skipped when unset).
 */
export async function openPaneFor(
  page: Page,
  host: OfficeWebHost,
  docUrl: string,
  buttonName: string
): Promise<Frame> {
  if (!process.env.ADV_E2E_NO_DEV_SERVER) return openPane(page, docUrl, buttonName);
  const origin = hostedOrigin();
  test.skip(!origin, `${host}: production-only run needs ADV_E2E_HOSTED_ORIGIN`);
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
export async function graphFromPane<T>(
  pane: Frame,
  method: string,
  path: string,
  body?: unknown
): Promise<T> {
  return pane.evaluate(
    async ({ method, path, body }) => {
      const token = await (
        window as unknown as { advAuth: { graphToken: () => Promise<string> } }
      ).advAuth.graphToken();
      const response = await fetch(`https://graph.microsoft.com/v1.0${path}`, {
        method,
        headers: {
          Authorization: `Bearer ${token}`,
          ...(body === undefined ? {} : { "Content-Type": "application/json" }),
        },
        body: body === undefined ? undefined : JSON.stringify(body),
      });
      if (!response.ok) throw new Error(`Graph ${method} ${path}: HTTP ${response.status}`);
      return response.status === 204 ? undefined : response.json();
    },
    { method, path, body }
  ) as Promise<T>;
}

/** Outlook on the web; the DEV copy (manifest.json, Teams custom app) carries the Outlook commands for the E2E account. */
export const OUTLOOK_URL = "https://outlook.cloud.microsoft/mail/";
export const OUTLOOK_DEV_ADDIN = "Global Advocate DEV";

/**
 * Opens the DEV Outlook add-in from the Apps flyout of the current surface (compose: "Apps", read: "More apps")
 * and returns its pane (served by the dev server). One button per surface, so the app tile opens the pane directly.
 */
export async function openOutlookPane(
  page: Page,
  appsButton: "Apps" | "More apps"
): Promise<Frame> {
  // taskpane.html, not commands.html: in compose Outlook also loads the banner runtime from the dev server.
  const pane = (): Frame | undefined =>
    page.frames().find((f) => f.url().startsWith(`${DEV_SERVER_ORIGIN}/taskpane.html`));
  // Outlook toggles the add-in pane: when it still thinks an earlier pane is open, the first click only closes it.
  for (let attempt = 1; ; attempt++) {
    await page
      .getByRole("button", { name: appsButton, exact: true })
      .last()
      .click({ timeout: 60_000 });
    await page
      .getByRole("listitem", { name: OUTLOOK_DEV_ADDIN, exact: true })
      .first()
      .click({ timeout: 30_000 });
    const opened = await expect
      .poll(() => !!pane(), { timeout: PANE_WAIT_MS / 2 })
      .toBe(true)
      .then(
        () => true,
        () => false
      );
    if (opened) return pane()!;
    if (attempt === 2)
      throw new Error(
        `no ${OUTLOOK_DEV_ADDIN} pane (DEV copy without Outlook commands? upload manifest.json v1.0.4+, see README)`
      );
  }
}
