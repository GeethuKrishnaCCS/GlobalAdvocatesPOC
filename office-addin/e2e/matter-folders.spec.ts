import { expect, test, type BrowserContext, type Frame, type Page } from "@playwright/test";
import { existsSync } from "node:fs";
import { API_BASE_URL } from "../src/shared/config";
import { captureConsole, graphFromPane, outlookHosted } from "./sideload";
import {
  FILING_FOLDER,
  FOLDER,
  MATTER,
  createTestEmail,
  deleteFiled,
  folderId,
  matterDriveId,
  setupPane,
  type IItem,
} from "./outlookHelpers";

/* Mailbox folders of a pinned matter: pin creates them, the backend (1-min timer) files what lands in them (ccsdev01). */

const MATTER_ID = MATTER.toLowerCase();
const CATEGORY = "Filed to matter";
const MARK = "ADVTest E2E Folders ";
const SWEEP_WAIT_MS = 240_000;
const POLL_EVERY_MS = 10_000;
const STORAGE_STATE = ".auth/office-web.json";

test.skip(
  !!process.env.ADV_E2E_NO_DEV_SERVER && !outlookHosted(),
  "Outlook runs the DEV add-in from the dev server only"
);
test.describe.configure({ mode: "serial" });

/** What the pane exchanged with the Function backend (method, path, status, body), for evidence. */
interface IApiCall {
  at: number;
  method: string;
  path: string;
  status: number;
  body: string;
}
interface IFolderStatus {
  matterRef: string;
  filedCount: number;
  lastSweepUtc: string | null;
  lastError: string | null;
}
interface IMailFolder {
  id: string;
  displayName: string;
}
interface IMessage {
  id: string;
  subject: string;
  categories: string[];
}

const apiCalls: IApiCall[] = [];
const stamp = `${Date.now()}`;
const subjectOf = (kind: string): string => `${MARK}${stamp} ${kind}`;
const enc = encodeURIComponent;

let context: BrowserContext;
let page: Page;
let pane: Frame;
let driveId: string;
let filingId: string;
let folderAId: string;
let mailMatterId: string;
let mailFolderAId: string;
let initiallyPinned = false;
const timings: string[] = [];

/** Latest status the pane read from GET /matter-folders for the test matter. */
function latestStatus(): (IFolderStatus & { at: number }) | undefined {
  for (const c of [...apiCalls].reverse()) {
    if (c.method !== "GET" || c.path !== "/matter-folders" || c.status !== 200) continue;
    const mine = (JSON.parse(c.body) as { matters: IFolderStatus[] }).matters.find(
      (m) => m.matterRef === MATTER
    );
    if (mine) return { ...mine, at: c.at };
  }
  return undefined;
}

async function mailChildren(parentId?: string): Promise<IMailFolder[]> {
  const path = parentId ? `/me/mailFolders/${parentId}/childFolders` : "/me/mailFolders";
  return (
    await graphFromPane<{ value: IMailFolder[] }>(
      pane,
      "GET",
      `${path}?$top=200&$select=id,displayName`
    )
  ).value;
}

/** Mail folders Matters → matter → children (undefined parts when missing). */
async function mailTree(): Promise<{
  matters?: IMailFolder;
  matter?: IMailFolder;
  children: IMailFolder[];
}> {
  const matters = (await mailChildren()).find((f) => f.displayName === "Matters");
  const matter = matters && (await mailChildren(matters.id)).find((f) => f.displayName === MATTER);
  return { matters, matter, children: matter ? await mailChildren(matter.id) : [] };
}

async function messagesIn(folder: string, filter: string): Promise<IMessage[]> {
  return (
    await graphFromPane<{ value: IMessage[] }>(
      pane,
      "GET",
      `/me/mailFolders/${folder}/messages?$filter=${enc(filter)}&$select=id,subject,categories&$top=50`
    )
  ).value;
}

const bySubject = (folder: string, subject: string): Promise<IMessage[]> =>
  messagesIn(folder, `subject eq '${subject}'`);

async function filedNames(parentId: string, part: string): Promise<string[]> {
  const children = await graphFromPane<{ value: IItem[] }>(
    pane,
    "GET",
    `/drives/${driveId}/items/${parentId}/children?$select=id,name,size&$top=999`
  );
  return children.value.map((c) => c.name).filter((n) => n.includes(part));
}

/** Deletes this spec's test emails (subject starts with `prefix`) wherever they are, and the .eml files they caused. */
async function cleanup(prefix: string): Promise<void> {
  const tree = await mailTree();
  const folders = [
    "inbox",
    ...(tree.matter ? [tree.matter.id] : []),
    ...tree.children.map((c) => c.id),
  ];
  for (const f of folders)
    for (const m of await messagesIn(f, `startswith(subject,'${prefix}')`))
      await graphFromPane(pane, "DELETE", `/me/messages/${m.id}`);
  await deleteFiled(pane, driveId, filingId, prefix);
  await deleteFiled(pane, driveId, folderAId, prefix);
}

async function setPinned(wanted: boolean): Promise<void> {
  const pin = pane.getByTestId("matter-browser").getByTestId(`matter-step-pin-${MATTER_ID}`);
  if ((await pin.getAttribute("aria-pressed")) !== String(wanted)) await pin.click();
  await expect(pin).toHaveAttribute("aria-pressed", String(wanted));
}

/** Waits for the next backend call of `method` after `since` and fails with its HTTP status and body when it was not OK. */
async function backendCall(method: string, since: number): Promise<IApiCall> {
  const matching = (): IApiCall | undefined =>
    apiCalls.slice(since).find((c) => c.method === method && c.path.startsWith("/matter-folders/"));
  await expect.poll(matching, { timeout: 120_000 }).toBeTruthy();
  const call = matching()!;
  if (call.status >= 400)
    throw new Error(`${method} ${call.path}: HTTP ${call.status} body ${call.body}`);
  return call;
}

/** Test email moved into a mail folder; returns the NEW id (a move changes it). */
async function createAndMove(kind: string, destination: string): Promise<string> {
  const id = await createTestEmail(pane, subjectOf(kind), `${stamp}${kind}`);
  const moved = await graphFromPane<{ id: string }>(pane, "POST", `/me/messages/${id}/move`, {
    destinationId: destination,
  });
  return moved.id;
}

/** Polls until the library holds the .eml and the mail has the category; on timeout throws with backend evidence. */
async function expectFiled(
  kind: string,
  mailFolder: string,
  libraryFolder: string,
  movedAt: number
): Promise<void> {
  const subject = subjectOf(kind);
  const state = async (): Promise<{ file: boolean; tagged: boolean }> => ({
    file: (await filedNames(libraryFolder, subject)).some((n) =>
      new RegExp(`^\\d{4}-\\d{2}-\\d{2} ${subject}\\.eml$`).test(n)
    ),
    tagged: (await bySubject(mailFolder, subject))[0]?.categories?.includes(CATEGORY) ?? false,
  });
  try {
    await expect
      .poll(state, { timeout: SWEEP_WAIT_MS, intervals: [POLL_EVERY_MS] })
      .toEqual({ file: true, tagged: true });
  } catch (e) {
    const msgs = await bySubject(mailFolder, subject);
    throw new Error(
      `not filed within ${SWEEP_WAIT_MS / 1000}s. status=${JSON.stringify(latestStatus())} ` +
        `inFolder=${msgs.length} categories=${JSON.stringify(msgs.map((m) => m.categories))}\n${e}`
    );
  }
  timings.push(`${kind}: filed+tagged ${Math.round((Date.now() - movedAt) / 1000)}s after move`);
}

test.beforeAll(async ({ browser }, testInfo) => {
  test.setTimeout(5 * 60_000);
  context = await browser.newContext({
    ignoreHTTPSErrors: true,
    storageState: existsSync(STORAGE_STATE) ? STORAGE_STATE : undefined,
  });
  page = await context.newPage();
  captureConsole(page, testInfo);
  page.on("response", async (r) => {
    if (!r.url().startsWith(API_BASE_URL) || r.request().method() === "OPTIONS") return;
    apiCalls.push({
      at: Date.now(),
      method: r.request().method(),
      path: new URL(r.url()).pathname.replace(new URL(API_BASE_URL).pathname, ""),
      status: r.status(),
      body: await r.text().catch(() => ""),
    });
  });
  pane = await setupPane(page).catch(() => setupPane(page)); // openOutlookPane sometimes opens nothing right after an email is selected
  driveId = await matterDriveId(pane);
  filingId = await folderId(pane, driveId, FILING_FOLDER);
  folderAId = await folderId(pane, driveId, FOLDER);
  await cleanup(MARK); // leftovers of an earlier run of this spec only
  await pane.getByTestId("file-to-matter-tab-matters").click();
  const pin = pane.getByTestId("matter-browser").getByTestId(`matter-step-pin-${MATTER_ID}`);
  await expect(pin).toBeVisible({ timeout: 60_000 });
  initiallyPinned = (await pin.getAttribute("aria-pressed")) === "true";
});

test.afterAll(async () => {
  test.setTimeout(3 * 60_000);
  if (!pane) return;
  // Pin state first, while our mail still sits in the folder (an unpin deletes only empty folders).
  await setPinned(initiallyPinned).catch((e) => console.log(`pin restore failed: ${e}`));
  await new Promise((resolve) => setTimeout(resolve, 5_000)); // let the DELETE/PUT reach the backend
  await cleanup(`${MARK}${stamp}`).catch((e) => console.log(`cleanup failed: ${e}`));
  console.log(`timings: ${timings.join(" | ")}`);
  await context.close();
});

test("1. pinning creates Matters / matter / library sub-folders", async () => {
  test.setTimeout(5 * 60_000);
  await setPinned(false);
  await new Promise((resolve) => setTimeout(resolve, 5_000)); // let the DELETE of the unpin finish
  const since = apiCalls.length;
  await setPinned(true);
  await backendCall("PUT", since);

  const status = pane.getByTestId("matter-folders-status");
  await expect(status).toBeVisible();
  await expect(pane.getByTestId("matter-folders-retry")).toHaveCount(0);
  await expect(pane.getByTestId(`matter-step-mailfolder-${MATTER_ID}`)).toBeVisible();

  const root = await graphFromPane<{ value: { name: string; folder?: object }[] }>(
    pane,
    "GET",
    `/drives/${driveId}/root/children?$select=name,folder&$top=999`
  );
  const library = root.value.filter((c) => c.folder && c.name !== "Forms").map((c) => c.name);
  await expect
    .poll(async () => (await mailTree()).children.map((c) => c.displayName), { timeout: 30_000 })
    .toEqual(expect.arrayContaining([FOLDER, "Correspondence", ...library]));
  const tree = await mailTree();
  mailMatterId = tree.matter!.id;
  mailFolderAId = tree.children.find((c) => c.displayName === FOLDER)!.id;
  console.log(
    `library folders ${library}; mail children ${tree.children.map((c) => c.displayName)}`
  );
});

test("2. an email moved into the matter folder is filed to Correspondence/Matter", async () => {
  test.setTimeout(6 * 60_000);
  const movedAt = Date.now();
  await createAndMove("Matter", mailMatterId);
  await expectFiled("Matter", mailMatterId, filingId, movedAt);
});

test("3. an email moved into a sub-folder is filed to that library folder", async () => {
  test.setTimeout(6 * 60_000);
  const movedAt = Date.now();
  await createAndMove("FolderA", mailFolderAId);
  await expectFiled("FolderA", mailFolderAId, folderAId, movedAt);
});

test("4. a filed email is not filed twice", async () => {
  test.setTimeout(6 * 60_000);
  // The sweep that filed "FolderA" (test 3) went over every folder of the matter again, the matter folder included.
  // The pane's status line (re-read every minute) must reach the two filings, with no third.
  await expect(pane.getByTestId("matter-folders-status")).toContainText("2 emails filed", {
    timeout: SWEEP_WAIT_MS,
  });
  expect(await filedNames(filingId, subjectOf("Matter"))).toHaveLength(1);
  expect(await filedNames(folderAId, subjectOf("FolderA"))).toHaveLength(1);
  expect(latestStatus()?.lastError).toBeNull();
});

test("5. unpin keeps the non-empty folder and stops filing", async () => {
  test.setTimeout(6 * 60_000);
  const since = apiCalls.length;
  await setPinned(false);
  await backendCall("DELETE", since);
  expect((await mailTree()).matter?.id, "matter mail folder still exists").toBe(mailMatterId);

  await createAndMove("Waiting", mailMatterId);
  // Negative check: nothing to poll for. Filing follows a move within seconds (change notification), so 45 s is ample.
  await page.waitForTimeout(45_000);
  const subject = subjectOf("Waiting");
  expect(await filedNames(filingId, subject)).toEqual([]);
  const msgs = await bySubject(mailMatterId, subject);
  expect(msgs).toHaveLength(1);
  expect(msgs[0].categories).not.toContain(CATEGORY);
});

test("6. re-pinning files the waiting email", async () => {
  test.setTimeout(6 * 60_000);
  const since = apiCalls.length;
  const movedAt = Date.now();
  await setPinned(true);
  await backendCall("PUT", since);
  await expectFiled("Waiting", mailMatterId, filingId, movedAt);
});
