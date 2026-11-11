import { expect, test, type Frame, type Page, type Route } from "@playwright/test";
import {
  captureConsole,
  graphFromPane,
  openOutlookPane,
  saveRunScreenshot,
  signIn,
} from "./sideload";
import {
  FILING_FOLDER,
  FOLDER,
  MATTER,
  createTestEmail,
  deleteFiled,
  folderId,
  matterDriveId,
  openEmail,
  setupPane,
  type IItem,
} from "./outlookHelpers";

/* File to Matter › "My matters" tab: drag-and-drop filing, filed emails, pinned matters (Outlook on the web, ccsdev01). */

const MATTER_ID = MATTER.toLowerCase();
const FOLDER_ROW = `folder-step-row-${FOLDER.toLowerCase().replace(/_/g, "-")}`;
const FOLDER_ITEM = `folder-step-item-${FOLDER.toLowerCase().replace(/_/g, "-")}`;

test.skip(
  !!process.env.ADV_E2E_NO_DEV_SERVER,
  "Outlook runs the DEV add-in from the dev server only"
);

test.afterEach(async ({ page }, testInfo) => {
  if (testInfo.status !== "skipped") await saveRunScreenshot(page, testInfo);
});

/** Opens the pane on the email and switches to "My matters", matter list loaded. */
async function openMyMatters(page: Page, subject: string): Promise<Frame> {
  await openEmail(page, subject);
  const pane = await openOutlookPane(page, "More apps");
  await signIn(page, pane);
  await pane.getByTestId("file-to-matter-tab-matters").click();
  await expect(
    pane.getByTestId("matter-browser").getByTestId(`matter-step-row-${MATTER_ID}`)
  ).toBeVisible({ timeout: 60_000 });
  return pane;
}

/**
 * Drops files onto an element of the pane the way classic Outlook and Mac do (HTML drag-and-drop with files).
 * Outlook on the web's own drag is covered by the real mouse drag test below.
 */
async function dropFiles(
  pane: Frame,
  testId: string,
  files: { name: string; text: string }[]
): Promise<void> {
  await pane
    .getByTestId("matter-browser")
    .getByTestId(testId)
    .evaluate((element, dropped) => {
      const dataTransfer = new DataTransfer();
      for (const f of dropped)
        dataTransfer.items.add(new File([f.text], f.name, { type: "message/rfc822" }));
      const fire = (type: string): boolean =>
        element.dispatchEvent(
          new DragEvent(type, { dataTransfer, bubbles: true, cancelable: true })
        );
      fire("dragover");
      fire("drop");
    }, files);
}

async function filedNames(
  pane: Frame,
  driveId: string,
  parentId: string,
  stamp: string
): Promise<string[]> {
  const children = await graphFromPane<{ value: IItem[] }>(
    pane,
    "GET",
    `/drives/${driveId}/items/${parentId}/children?$select=id,name,size&$top=999`
  );
  return children.value
    .filter((c) => c.name.includes(stamp))
    .map((c) => c.name)
    .sort();
}

const eml = (subject: string): string =>
  `From: advtest@example.com\r\nTo: advtest@example.com\r\nSubject: ${subject}\r\nMIME-Version: 1.0\r\nContent-Type: text/plain\r\n\r\nADVTest E2E dropped email.\r\n`;

test("my matters: emails dropped on a matter, a folder row and the open folder are filed and listed", async ({
  page,
}, testInfo) => {
  test.setTimeout(8 * 60_000);
  captureConsole(page, testInfo);
  const stamp = `${Date.now()}`;
  const subject = `ADVTest E2E Drop ${stamp}`;
  const fileName = `${subject}.eml`;
  let setup = await setupPane(page);
  const driveId = await matterDriveId(setup);
  const filingId = await folderId(setup, driveId, FILING_FOLDER);
  const folderAId = await folderId(setup, driveId, FOLDER);
  const messageId = await createTestEmail(setup, subject, stamp);

  try {
    const pane = await openMyMatters(page, subject);
    setup = pane;
    await expect(pane.getByTestId("matter-browser-hint")).toContainText("Drag emails");
    await expect(pane.getByTestId("drop-queue")).toHaveCount(0);

    // 1. Onto a matter row: goes to the matter's filing folder. Two files in one drag.
    await dropFiles(pane, `matter-step-row-${MATTER_ID}`, [
      { name: fileName, text: eml(subject) },
      { name: `ADVTest E2E Second ${stamp}.eml`, text: eml("second") },
    ]);
    await expect(pane.getByTestId("drop-queue-summary")).toHaveText(
      "2 filed. The emails stay in your mailbox.",
      { timeout: 120_000 }
    );
    await expect(pane.locator('[data-testid^="drop-queue-target-"]').first()).toHaveText(
      `→ ${MATTER} › Correspondence › Matter`
    );
    expect(await filedNames(pane, driveId, filingId, stamp)).toEqual([
      `ADVTest E2E Drop ${stamp}.eml`,
      `ADVTest E2E Second ${stamp}.eml`,
    ]);
    await pane.getByTestId("drop-queue-clear").click();
    await expect(pane.getByTestId("drop-queue")).toHaveCount(0);

    // 2. Onto a folder row.
    await pane.getByTestId("matter-browser").getByTestId(`matter-step-item-${MATTER_ID}`).click();
    await expect(pane.getByTestId(FOLDER_ROW)).toBeVisible({ timeout: 60_000 });
    await dropFiles(pane, FOLDER_ROW, [{ name: fileName, text: eml(subject) }]);
    await expect(pane.getByTestId("drop-queue-summary")).toHaveText(
      "1 filed. The emails stay in your mailbox.",
      { timeout: 120_000 }
    );
    await expect(pane.locator('[data-testid^="drop-queue-target-"]').first()).toHaveText(
      `→ ${MATTER} › ${FOLDER}`
    );

    // 3. Open that folder: the filed email is listed; dropping the same name again keeps both (no overwrite).
    await pane.getByTestId(FOLDER_ITEM).click();
    const rows = pane.locator('[data-testid^="filed-emails-row-"]').filter({ hasText: stamp });
    await expect(rows).toHaveCount(1, { timeout: 60_000 });
    await expect(rows.first().getByRole("link")).toHaveAttribute(
      "href",
      /^https:\/\/ccsdev01\.sharepoint\.com\//
    );
    await dropFiles(pane, "matter-browser-folder", [{ name: fileName, text: eml(subject) }]);
    await expect(rows).toHaveCount(2, { timeout: 120_000 });
    expect(await filedNames(pane, driveId, folderAId, stamp)).toEqual([
      `ADVTest E2E Drop ${stamp} 1.eml`,
      `ADVTest E2E Drop ${stamp}.eml`,
    ]);
    await expect(pane.getByTestId("filed-emails-heading")).toHaveText(/^Filed emails \(\d+\)$/);

    // 4. Not on a target: nothing is filed, the pane says where to drop.
    await dropFiles(pane, "matter-browser-hint", [{ name: fileName, text: eml(subject) }]);
    await expect(pane.getByTestId("matter-browser-notice")).toHaveText(
      "Drop emails onto a matter or a folder."
    );
    expect(await filedNames(pane, driveId, folderAId, stamp)).toHaveLength(2);

    // Back to the matter list; the wizard tab kept its own state meanwhile.
    await pane.getByTestId("matter-browser-back").click();
    await expect(
      pane.getByTestId("matter-browser").getByTestId(`matter-step-row-${MATTER_ID}`)
    ).toBeVisible();
    await pane.getByTestId("file-to-matter-tab-email").click();
    await expect(pane.getByTestId("save-step-matter")).toBeVisible();
    await expect(pane.getByTestId("matter-browser")).toBeHidden();
  } finally {
    await deleteFiled(setup, driveId, filingId, stamp).catch(() => undefined);
    await deleteFiled(setup, driveId, folderAId, stamp).catch(() => undefined);
    await graphFromPane(setup, "DELETE", `/me/messages/${messageId}`).catch(() => undefined);
  }
});

test("my matters: an email dragged from the Outlook message list onto a matter is filed", async ({
  page,
}, testInfo) => {
  test.setTimeout(8 * 60_000);
  captureConsole(page, testInfo);
  const stamp = `${Date.now()}`;
  const subject = `ADVTest E2E Drag ${stamp}`;
  let setup = await setupPane(page);
  const driveId = await matterDriveId(setup);
  const filingId = await folderId(setup, driveId, FILING_FOLDER);
  const messageId = await createTestEmail(setup, subject, stamp);

  try {
    const pane = await openMyMatters(page, subject);
    setup = pane;
    const from = (await page
      .getByRole("option", { name: new RegExp(subject) })
      .first()
      .boundingBox())!;
    const to = (await pane
      .getByTestId("matter-browser")
      .getByTestId(`matter-step-row-${MATTER_ID}`)
      .boundingBox())!;
    const target = { x: to.x + to.width / 2, y: to.y + to.height / 2 };
    await page.mouse.move(from.x + from.width / 2, from.y + from.height / 2);
    await page.mouse.down();
    // Several moves: Outlook starts its drag only after the pointer travelled, and reports dragover as it moves.
    await page.mouse.move(from.x + from.width / 2 + 30, from.y + from.height / 2 + 10, {
      steps: 5,
    });
    await page.mouse.move(target.x, target.y, { steps: 25 });
    await page.mouse.move(target.x + 2, target.y + 2, { steps: 3 });
    await expect(
      pane.getByTestId("matter-browser").getByTestId(`matter-step-row-${MATTER_ID}`)
    ).toHaveCSS("outline-style", "dashed");
    await page.screenshot({ path: testInfo.outputPath("drag-hover.png") });
    await page.mouse.up();

    await expect(pane.getByTestId("drop-queue-summary")).toHaveText(
      "1 filed. The emails stay in your mailbox.",
      { timeout: 120_000 }
    );
    const names = await filedNames(pane, driveId, filingId, stamp);
    expect(names).toEqual([`${subject}.eml`]);
  } finally {
    await deleteFiled(setup, driveId, filingId, stamp).catch(() => undefined);
    await graphFromPane(setup, "DELETE", `/me/messages/${messageId}`).catch(() => undefined);
  }
});

const PIN = (id: string): string => `matter-step-pin-${id}`;
const MATTER_ROWS = '[data-testid^="matter-step-row-"]';

/** Clicks the pin of a matter in the open tab until it is in the wanted state. */
async function setPinned(pane: Frame, id: string, wanted: boolean): Promise<void> {
  const pin = pane.getByTestId("matter-browser").getByTestId(PIN(id));
  if ((await pin.getAttribute("aria-pressed")) !== String(wanted)) await pin.click();
  await expect(pin).toHaveAttribute("aria-pressed", String(wanted));
}

test("my matters: a pinned matter is listed first and stays pinned after a reload", async ({
  page,
}, testInfo) => {
  test.setTimeout(6 * 60_000);
  captureConsole(page, testInfo);
  const stamp = `${Date.now()}`;
  const subject = `ADVTest E2E Pin ${stamp}`;
  const setup = await setupPane(page);
  const messageId = await createTestEmail(setup, subject, stamp);
  let pane = setup;
  const pinned: string[] = [];

  try {
    pane = await openMyMatters(page, subject);
    const pin = pane.getByTestId("matter-browser").getByTestId(PIN(MATTER_ID));
    // Left pinned by a failed earlier run: start from unpinned.
    await setPinned(pane, MATTER_ID, false);

    pinned.push(MATTER_ID);
    await pin.click();
    await expect(pin).toHaveAttribute("aria-pressed", "true");
    await expect(pin).toHaveAccessibleName(`Unpin ${MATTER}`);

    // One pin state for both tabs.
    await pane.getByTestId("file-to-matter-tab-email").click();
    await expect(
      pane.getByTestId("file-to-matter-panel-email").getByTestId(PIN(MATTER_ID))
    ).toHaveAttribute("aria-pressed", "true");

    // Stored in Outlook roaming settings: survives a reload.
    pane = await openMyMatters(page, subject);
    await expect(pane.getByTestId("matter-browser").getByTestId(PIN(MATTER_ID))).toHaveAttribute(
      "aria-pressed",
      "true"
    );

    // Pinned first: needs a second matter, and ADVTestMatter01 unpinned so its list position cannot decide the order.
    await setPinned(pane, MATTER_ID, false);
    pinned.length = 0;
    const rows = pane.getByTestId("matter-browser").locator(MATTER_ROWS);
    if ((await rows.count()) >= 2) {
      const lastRow = (await rows.last().getAttribute("data-testid"))!;
      const lastId = lastRow.replace("matter-step-row-", "");
      const lastPin = pane.getByTestId("matter-browser").getByTestId(PIN(lastId));
      // Already pinned by something else: the order proves nothing and the pin is not ours to remove.
      if ((await lastPin.getAttribute("aria-pressed")) === "false") {
        pinned.push(lastId);
        await lastPin.click();
        await expect(lastPin).toHaveAttribute("aria-pressed", "true");
        await expect.poll(() => rows.first().getAttribute("data-testid")).toBe(lastRow);
      }
    }
  } finally {
    for (const id of pinned) await setPinned(pane, id, false).catch(() => undefined);
    await graphFromPane(pane, "DELETE", `/me/messages/${messageId}`).catch(() => undefined);
  }
});

test("my matters: a failed upload shows as failed and Retry files it", async ({
  page,
}, testInfo) => {
  test.setTimeout(6 * 60_000);
  captureConsole(page, testInfo);
  const stamp = `${Date.now()}`;
  const subject = `ADVTest E2E Retry ${stamp}`;
  const fileName = `${subject}.eml`;
  let setup = await setupPane(page);
  const driveId = await matterDriveId(setup);
  const filingId = await folderId(setup, driveId, FILING_FOLDER);
  const messageId = await createTestEmail(setup, subject, stamp);
  // The Graph SDK retries some failures itself: keep aborting every simple upload until unrouted.
  const uploadUrl = /graph\.microsoft\.com\/.*\/content/;
  const abortUploads = (route: Route): Promise<void> =>
    route.request().method() === "PUT" ? route.abort("failed") : route.continue();

  try {
    const pane = await openMyMatters(page, subject);
    setup = pane;
    await page.route(uploadUrl, abortUploads);
    await dropFiles(pane, `matter-step-row-${MATTER_ID}`, [{ name: fileName, text: eml(subject) }]);

    await expect(pane.getByTestId("drop-queue-summary")).toHaveText("0 filed, 1 failed", {
      timeout: 120_000,
    });
    await expect(pane.locator('[data-testid^="drop-queue-status-"]')).toHaveAttribute(
      "data-status",
      "failed"
    );
    await expect(pane.locator('[data-testid^="drop-queue-error-"]')).toBeVisible();
    const retry = pane.locator('[data-testid^="drop-queue-retry-"]');
    await expect(retry).toBeVisible();
    expect(await filedNames(pane, driveId, filingId, stamp)).toEqual([]);

    await page.unroute(uploadUrl, abortUploads);
    await retry.click();
    await expect(pane.getByTestId("drop-queue-summary")).toHaveText(
      "1 filed. The emails stay in your mailbox.",
      { timeout: 120_000 }
    );
    expect(await filedNames(pane, driveId, filingId, stamp)).toEqual([fileName]);
  } finally {
    await page.unroute(uploadUrl, abortUploads).catch(() => undefined);
    await deleteFiled(setup, driveId, filingId, stamp).catch(() => undefined);
    await graphFromPane(setup, "DELETE", `/me/messages/${messageId}`).catch(() => undefined);
  }
});

test("my matters: a matter list that fails to load shows an error and Retry recovers", async ({
  page,
}, testInfo) => {
  test.setTimeout(6 * 60_000);
  captureConsole(page, testInfo);
  const stamp = `${Date.now()}`;
  const subject = `ADVTest E2E List ${stamp}`;
  let setup = await setupPane(page);
  const messageId = await createTestEmail(setup, subject, stamp);
  // listMatters is GET /sites?search=*; the drive lookups (/sites/<id>/...) are not matched.
  const listUrl = /graph\.microsoft\.com\/v1\.0\/sites\?/;
  const abortList = (route: Route): Promise<void> => route.abort("failed");

  try {
    await openEmail(page, subject);
    const pane = await openOutlookPane(page, "More apps");
    await signIn(page, pane);
    setup = pane; // the setup pane went away with the navigation; clean up through this one
    // The "This email" tab loads matters on mount: route only after that, before the other tab loads its own.
    await expect(
      pane.getByTestId("file-to-matter").getByTestId(`matter-step-row-${MATTER_ID}`)
    ).toBeVisible({ timeout: 60_000 });
    await page.route(listUrl, abortList);
    await pane.getByTestId("file-to-matter-tab-matters").click();

    const browser = pane.getByTestId("matter-browser");
    await expect(browser.getByTestId("matter-browser-error")).toBeVisible({ timeout: 60_000 });
    const retry = browser.getByTestId("matter-step-retry");
    await expect(retry).toBeVisible();

    await page.unroute(listUrl, abortList);
    await retry.click();
    await expect(browser.getByTestId(`matter-step-row-${MATTER_ID}`)).toBeVisible({
      timeout: 60_000,
    });
    await expect(browser.getByTestId("matter-browser-error")).toHaveCount(0);
  } finally {
    await page.unroute(listUrl, abortList).catch(() => undefined);
    await graphFromPane(setup, "DELETE", `/me/messages/${messageId}`).catch(() => undefined);
  }
});

/** Tabs, button names, horizontal overflow and icon sizes of the "My matters" tab as it is on screen. */
async function expectTabLooksRight(pane: Frame): Promise<void> {
  await expect(pane.getByTestId("file-to-matter-tablist").getByRole("tab")).toHaveCount(2);
  await expect(pane.getByTestId("file-to-matter-tab-matters")).toHaveAttribute(
    "aria-selected",
    "true"
  );
  await expect(pane.getByTestId("file-to-matter-tab-email")).toHaveAttribute(
    "aria-selected",
    "false"
  );

  const unnamed = await pane.getByTestId("matter-browser").evaluate((root) =>
    Array.from(root.querySelectorAll("button"))
      .filter((b) => !(b.getAttribute("aria-label") || b.textContent || "").trim())
      .map((b) => b.getAttribute("data-testid") ?? b.outerHTML.slice(0, 80))
  );
  expect(unnamed, "buttons without an accessible name").toEqual([]);

  const overflow = await pane.evaluate(() => ({
    scroll: document.documentElement.scrollWidth,
    view: window.innerWidth,
  }));
  expect(overflow.scroll, "no horizontal scroll at pane width").toBeLessThanOrEqual(overflow.view);

  const collapsed = await pane.getByTestId("matter-browser").evaluate((root) =>
    Array.from(root.querySelectorAll("svg"))
      .filter((s) => {
        const r = s.getBoundingClientRect();
        return r.width <= 0 || r.height <= 0;
      })
      .map((s) => s.parentElement?.getAttribute("data-testid") ?? "svg")
  );
  expect(collapsed, "icons with no rendered size").toEqual([]);
}

test("my matters: tabs, buttons, icons and width are accessible in the list and the folder view", async ({
  page,
}, testInfo) => {
  test.setTimeout(6 * 60_000);
  captureConsole(page, testInfo);
  const stamp = `${Date.now()}`;
  const subject = `ADVTest E2E Layout ${stamp}`;
  let setup = await setupPane(page);
  const messageId = await createTestEmail(setup, subject, stamp);

  try {
    const pane = await openMyMatters(page, subject);
    setup = pane;
    await expectTabLooksRight(pane);

    await pane.getByTestId("matter-browser").getByTestId(`matter-step-item-${MATTER_ID}`).click();
    await expect(pane.getByTestId("filed-emails-heading")).toBeVisible({ timeout: 60_000 });
    await expect(pane.getByTestId("matter-browser-back")).toHaveText("All matters");
    await expectTabLooksRight(pane);
    await page.screenshot({ path: testInfo.outputPath("my-matters-folder.png") });
  } finally {
    await graphFromPane(setup, "DELETE", `/me/messages/${messageId}`).catch(() => undefined);
  }
});
