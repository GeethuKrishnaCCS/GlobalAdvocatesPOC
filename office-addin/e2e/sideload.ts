import { expect, type Frame, type Page, type TestInfo } from "@playwright/test";
import { readFileSync, appendFileSync } from "node:fs";

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
  return `${docUrl}&wdaddindevserverport=${port}&wdaddinmanifestfile=${MANIFEST_FILE}&wdaddinmanifestguid=${id}&wdaddintest=true`;
}

/** Pipes page + iframe console output and page errors to <test output>/console.log (grep it; never dump DOM). */
export function captureConsole(page: Page, testInfo: TestInfo): void {
  const file = testInfo.outputPath("console.log");
  const write = (line: string): void => appendFileSync(file, `${new Date().toISOString()} ${line}\n`);
  page.on("console", (m) => write(`[${m.type()}] ${m.location().url?.startsWith(DEV_SERVER_ORIGIN) ? "PANE " : ""}${m.text()}`));
  page.on("pageerror", (e) => write(`[pageerror] ${e.message}`));
  page.on("requestfailed", (r) => r.url().startsWith(DEV_SERVER_ORIGIN) && write(`[requestfailed] ${r.url()} ${r.failure()?.errorText}`));
}

/** First frame (any depth) where the ribbon tab is visible; Office on the web hosts the ribbon in a nested frame. */
async function findRibbonFrame(page: Page, tabName: string): Promise<Frame> {
  let found: Frame | undefined;
  await expect
    .poll(
      async () => {
        for (const frame of page.frames()) {
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
