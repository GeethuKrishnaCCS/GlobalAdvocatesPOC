import { expect, test, type Frame, type Page } from "@playwright/test";
import { FILE_TO_FOLDER_PROPERTY_ID, MATTER_LIBRARY_NAME } from "../src/shared/config";
import {
  OUTLOOK_URL,
  captureConsole,
  graphFromPane,
  openOutlookPane,
  saveRunScreenshot,
  signIn,
} from "./sideload";

/** Test matter / folder on ccsdev01 and an existing ADVTest document to attach. */
const MATTER = "ADVTestMatter01";
const MATTER_SITE = "ccsdev01.sharepoint.com:/sites/ADVTestMatter01";
const FOLDER = "ADVTest_FolderA";
const FILING_FOLDER = "Correspondence/Matter";
const ATTACH_DOC = "ADVTest_Order_0001.docx";
/** PR_MESSAGE_FLAGS = read, not unsent: a Graph-created message then shows as received mail, not a draft. */
const RECEIVED_FLAGS = { id: "Integer 0x0E07", value: "1" };

test.skip(!!process.env.ADV_E2E_NO_DEV_SERVER, "Outlook runs the DEV add-in from the dev server only");

test.afterEach(async ({ page }, testInfo) => {
  if (testInfo.status !== "skipped") await saveRunScreenshot(page, testInfo);
});

interface IItem {
  id: string;
  name: string;
  size: number;
}

async function matterDriveId(pane: Frame): Promise<string> {
  const site = await graphFromPane<{ id: string }>(pane, "GET", `/sites/${MATTER_SITE}?$select=id`);
  const lists = await graphFromPane<{ value: { id: string; name: string }[] }>(pane, "GET", `/sites/${site.id}/lists?$select=id,name`);
  const list = lists.value.find((l) => l.name === MATTER_LIBRARY_NAME)!;
  return (await graphFromPane<{ id: string }>(pane, "GET", `/sites/${site.id}/lists/${list.id}/drive?$select=id`)).id;
}

async function folderId(pane: Frame, driveId: string, path: string): Promise<string> {
  return (await graphFromPane<{ id: string }>(pane, "GET", `/drives/${driveId}/root:/${path}?$select=id`)).id;
}

/** Received-looking test email in the Inbox with two small attachments; returns its Graph id. */
async function createTestEmail(
  pane: Frame,
  subject: string,
  stamp: string,
  { props = [], ...extra }: { props?: object[]; categories?: string[] } = {}
): Promise<string> {
  const attachment = (name: string, text: string): object => ({
    "@odata.type": "#microsoft.graph.fileAttachment",
    name,
    contentBytes: Buffer.from(text).toString("base64"),
  });
  const message = await graphFromPane<{ id: string }>(pane, "POST", "/me/mailFolders/inbox/messages", {
    subject,
    body: { contentType: "Text", content: "ADVTest E2E message for File to Matter." },
    inferenceClassification: "focused",
    attachments: [
      attachment(`ADVTest_E2E_File_${stamp}.txt`, "hello matter"),
      attachment(`ADVTest_E2E_File_${stamp}.csv`, "a,b\n1,2"),
    ],
    ...extra,
    singleValueExtendedProperties: [RECEIVED_FLAGS, ...props],
  });
  return message.id;
}

async function openEmail(page: Page, subject: string): Promise<void> {
  await page.goto(OUTLOOK_URL);
  await page.getByRole("option", { name: new RegExp(subject) }).first().click({ timeout: 60_000 });
}

async function deleteFiled(pane: Frame, driveId: string, parentId: string, stamp: string): Promise<void> {
  const children = await graphFromPane<{ value: IItem[] }>(pane, "GET", `/drives/${driveId}/items/${parentId}/children?$select=id,name,size&$top=999`);
  for (const item of children.value.filter((c) => c.name.includes(stamp))) {
    await graphFromPane(pane, "DELETE", `/drives/${driveId}/items/${item.id}`);
  }
}

/** Pane on any Inbox email, signed in: used only to set up test data through Graph. */
async function setupPane(page: Page): Promise<Frame> {
  await page.goto(OUTLOOK_URL);
  await page.getByRole("option").first().click({ timeout: 60_000 });
  const pane = await openOutlookPane(page, "More apps");
  await signIn(page, pane);
  // Test emails a failed run could not delete (its pane was gone): ADVTest-marked, ours.
  const filter = encodeURIComponent("startswith(subject,'ADVTest E2E ')");
  const stale = await graphFromPane<{ value: { id: string }[] }>(pane, "GET", `/me/mailFolders/inbox/messages?$filter=${filter}&$select=id`);
  for (const m of stale.value) await graphFromPane(pane, "DELETE", `/me/messages/${m.id}`);
  return pane;
}

test("outlook read: file to matter files the email (.eml) and its attachments", async ({ page }, testInfo) => {
  test.setTimeout(6 * 60_000);
  captureConsole(page, testInfo);
  const stamp = `${Date.now()}`;
  const subject = `ADVTest E2E File ${stamp}`;
  let setup = await setupPane(page);
  const driveId = await matterDriveId(setup);
  const targetId = await folderId(setup, driveId, FOLDER);
  const messageId = await createTestEmail(setup, subject, stamp);

  try {
    await openEmail(page, subject);
    const pane = await openOutlookPane(page, "More apps");
    await signIn(page, pane);
    setup = pane; // the setup pane went away with the navigation; clean up through this one
    await expect(pane.getByTestId("save-step-matter")).toHaveText("1. Matter");
    await pane.getByTestId(`matter-step-item-${MATTER.toLowerCase()}`).click({ timeout: 60_000 });
    await pane.getByTestId(`folder-step-item-${FOLDER.toLowerCase().replace(/_/g, "-")}`).click();
    await expect(pane.getByTestId("folder-step-path-1")).toHaveText(FOLDER);
    await pane.getByTestId("file-to-matter-primary").click();

    // Items: email + both attachments ticked by default.
    await expect(pane.getByTestId("items-step-check-email")).toBeChecked();
    const rows = pane.locator('[data-testid^="items-step-row-"]');
    await expect(rows).toHaveCount(3);
    await expect(pane.getByTestId("file-to-matter-summary")).toHaveText(`3 item(s) → ${MATTER} › ${FOLDER}`);
    await pane.getByTestId("file-to-matter-primary").click();

    await expect(pane.getByTestId("file-to-matter-success")).toContainText("Filed 3 item(s)", { timeout: 120_000 });
    const filed = await graphFromPane<{ value: IItem[] }>(pane, "GET", `/drives/${driveId}/items/${targetId}/children?$select=id,name,size&$top=999`);
    const mine = filed.value.filter((f) => f.name.includes(stamp));
    expect(mine.map((f) => f.name).sort()).toEqual(
      // Date-prefixed .eml sorts first (digits before letters).
      [expect.stringMatching(new RegExp(`^\\d{4}-\\d{2}-\\d{2} ${subject}\\.eml$`)), `ADVTest_E2E_File_${stamp}.csv`, `ADVTest_E2E_File_${stamp}.txt`]
    );
    expect(mine.every((f) => f.size > 0)).toBe(true);
    // The email stays in the mailbox.
    expect((await graphFromPane<{ id: string }>(pane, "GET", `/me/messages/${messageId}?$select=id`)).id).toBe(messageId);
  } finally {
    await deleteFiled(setup, driveId, targetId, stamp).catch(() => undefined);
    await graphFromPane(setup, "DELETE", `/me/messages/${messageId}`).catch(() => undefined);
  }
});

test("outlook read: an email marked by Word Send opens on its filing folder", async ({ page }, testInfo) => {
  test.setTimeout(5 * 60_000);
  captureConsole(page, testInfo);
  const stamp = `${Date.now()}`;
  const subject = `ADVTest E2E Prefill ${stamp}`;
  let setup = await setupPane(page);
  const driveId = await matterDriveId(setup);
  const filingId = await folderId(setup, driveId, FILING_FOLDER);
  const messageId = await createTestEmail(setup, subject, stamp, {
    categories: [`Matter: ${MATTER}`],
    props: [{ id: FILE_TO_FOLDER_PROPERTY_ID, value: `${driveId}|${filingId}` }],
  });

  try {
    await openEmail(page, subject);
    const pane = await openOutlookPane(page, "More apps");
    await signIn(page, pane);
    setup = pane;
    await expect(pane.getByTestId("file-to-matter-summary")).toHaveText(
      `3 item(s) → ${MATTER} › Correspondence › Matter`,
      { timeout: 90_000 }
    );
    await expect(pane.getByTestId("save-step-items")).toHaveAttribute("aria-current", "step");
    // Back shows the pre-selected folder in the browser.
    await pane.getByTestId("file-to-matter-back").click();
    await expect(pane.getByTestId("folder-step-path-2")).toHaveText("Matter");
  } finally {
    await graphFromPane(setup, "DELETE", `/me/messages/${messageId}`).catch(() => undefined);
  }
});

test("outlook compose: attach from matter attaches a document once", async ({ page }, testInfo) => {
  test.setTimeout(5 * 60_000);
  captureConsole(page, testInfo);
  const subject = `ADVTest E2E Attach ${Date.now()}`;
  await page.goto(`${OUTLOOK_URL}deeplink/compose`);
  await page.getByPlaceholder("Add a subject").fill(subject, { timeout: 60_000 });
  const pane = await openOutlookPane(page, "Apps");
  await signIn(page, pane);

  try {
    await expect(pane.getByTestId("save-step-documents")).toHaveText("2. Documents");
    await pane.getByTestId(`matter-step-item-${MATTER.toLowerCase()}`).click({ timeout: 60_000 });
    // Browse shows the library's folders; search finds the document anywhere in it.
    await expect(pane.getByTestId(`folder-step-item-${FOLDER.toLowerCase().replace(/_/g, "-")}`)).toBeVisible();
    await pane.getByTestId("documents-step-search").fill(ATTACH_DOC.replace(/\.docx$/, ""));
    await pane.getByTestId("documents-step-search").press("Enter");
    const docId = "advtest-order-0001-docx";
    const check = pane.getByTestId(`documents-step-check-${docId}`);
    await expect(check).toBeEnabled({ timeout: 60_000 });
    await check.check();
    await expect(pane.getByTestId("attach-from-matter-summary")).toHaveText("1 selected");
    await pane.getByTestId("attach-from-matter-primary").click();

    await expect(pane.getByTestId("attach-from-matter-success")).toContainText(ATTACH_DOC, { timeout: 90_000 });
    // Marked as attached and locked: ticking it again is impossible (no duplicate).
    await expect(pane.getByTestId(`documents-step-note-${docId}`)).toContainText("Attached");
    await expect(check).toBeDisabled();
    await expect(page.getByText(ATTACH_DOC).first()).toBeVisible();
  } finally {
    const filter = encodeURIComponent(`subject eq '${subject}'`);
    const drafts = await graphFromPane<{ value: { id: string }[] }>(pane, "GET", `/me/mailFolders/drafts/messages?$filter=${filter}&$select=id`).catch(() => ({ value: [] }));
    for (const d of drafts.value) await graphFromPane(pane, "DELETE", `/me/messages/${d.id}`).catch(() => undefined);
  }
});
