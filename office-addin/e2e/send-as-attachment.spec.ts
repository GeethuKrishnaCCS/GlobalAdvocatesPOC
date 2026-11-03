import { expect, test, type Frame } from "@playwright/test";
import { FILE_TO_FOLDER_PROPERTY_ID as PROPERTY_ID } from "../src/shared/config";
import { captureConsole, docUrlFor, graphFromPane, openPane, saveRunScreenshot, signIn } from "./sideload";

/** Matter the Word test document lives in (pre-selected) and the draft's test-marked names. */
const MATTER_REF = "ADVTestMatter01";
const WORD_NAME = "ADVTest_E2E_Send_Word";
const PDF_NAME = "ADVTest_E2E_Send_PDF";
const CREATE_TIMEOUT = 120_000;
const SUBJECT_PREFIX = "ADVTest E2E send";

interface IDraft {
  id: string;
  subject: string;
  categories: string[];
  toRecipients: { emailAddress: { address: string } }[];
  attachments: { name: string; size: number }[];
  singleValueExtendedProperties?: { id: string; value: string }[];
}

test.afterEach(async ({ page }, testInfo) => {
  if (testInfo.status !== "skipped") await saveRunScreenshot(page, testInfo);
});

async function findDraft(pane: Frame, subject: string): Promise<IDraft | undefined> {
  const filter = encodeURIComponent(`subject eq '${subject.replace(/'/g, "''")}'`);
  const expand = encodeURIComponent(
    `attachments($select=name,size),singleValueExtendedProperties($filter=id eq '${PROPERTY_ID}')`
  );
  const result = await graphFromPane<{ value: IDraft[] }>(
    pane,
    "GET",
    `/me/mailFolders/drafts/messages?$filter=${filter}&$select=id,subject,categories,toRecipients&$expand=${expand}`
  );
  return result.value[0];
}

test("word: send as attachment creates a tagged Outlook draft with Word + PDF", async ({ page }, testInfo) => {
  const docUrl = docUrlFor("word");
  test.skip(!docUrl, "no test document configured for word (e2e/.env)");
  test.setTimeout(5 * 60_000);
  captureConsole(page, testInfo);

  const pane = await openPane(page, docUrl!, "Send as Attachment");
  await signIn(page, pane);
  const me = /\(([^)]+@[^)]+)\)/.exec((await pane.getByTestId("app-user").textContent()) ?? "")![1];

  // Matter of the open document is pre-selected; its Correspondence › Matter folder enables filing.
  await expect(pane.getByTestId("send-matter-picker")).toContainText(MATTER_REF, { timeout: 60_000 });
  await expect(pane.getByTestId("send-matter-file-email")).toBeEnabled({ timeout: 30_000 });

  await pane.getByTestId("send-attach-format-both").click();
  await pane.getByTestId("send-attach-name-native").fill(WORD_NAME);
  await pane.getByTestId("send-attach-name-pdf").fill(`${PDF_NAME}.pdf`);
  await pane.getByTestId("send-matter-file-email").check();

  await pane.getByTestId("send-message-cc").fill("not-an-address");
  await expect(pane.getByTestId("send-message-cc-field")).toContainText("is not an email address");
  await expect(pane.getByTestId("send-create")).toBeDisabled();
  await pane.getByTestId("send-message-cc").fill("");
  await pane.getByTestId("send-message-to").fill(me);

  const subject = `${SUBJECT_PREFIX} ${Date.now()}`;
  await pane.getByTestId("send-message-subject").fill(subject);
  const tagged = `[${MATTER_REF}] ${subject}`;
  await expect(pane.getByTestId("send-message-subject-field")).toContainText(`Sent as: ${tagged}`);
  await expect(pane.getByTestId("send-summary")).toHaveText(`2 attachments · to ${me}`);
  await page.screenshot({ path: testInfo.outputPath("send-form.png") });
  await pane.getByTestId("send-create").click();

  // From here a draft may exist: whatever fails, the finally removes this run's test draft.
  try {
    await expect(pane.getByTestId("send-result")).toBeVisible({ timeout: CREATE_TIMEOUT });
    await expect(pane.getByTestId("send-result-attachment-name")).toHaveText([`${WORD_NAME}.docx`, `${PDF_NAME}.pdf`]);

    const draft = await findDraft(pane, tagged);
    expect(draft, "draft in Drafts").toBeDefined();
    expect(draft!.categories).toEqual([`Matter: ${MATTER_REF}`]);
    expect(draft!.toRecipients.map((r) => r.emailAddress.address.toLowerCase())).toEqual([me.toLowerCase()]);
    expect(draft!.attachments.map((a) => a.name).sort()).toEqual([`${PDF_NAME}.pdf`, `${WORD_NAME}.docx`]);
    expect(draft!.attachments.every((a) => a.size > 0)).toBe(true);
    expect(draft!.singleValueExtendedProperties?.[0]?.value).toMatch(/^b!.+\|.+$/);

    // Opening the draft: a real link, so Outlook on the web opens in a new tab from the click.
    const popup = page.context().waitForEvent("page", { timeout: 30_000 });
    await pane.getByTestId("send-result-open").click();
    const outlook = await popup;
    await expect.poll(() => outlook.url(), { timeout: 30_000 }).toMatch(/outlook\.(office|cloud\.microsoft)/);
    await outlook.close();
  } finally {
    // This run's draft, plus any an earlier failed run left behind (ADVTest subject only); they move to Deleted Items.
    const filter = encodeURIComponent(`startswith(subject,'[${MATTER_REF}] ${SUBJECT_PREFIX}')`);
    const leftovers = await graphFromPane<{ value: { id: string }[] }>(
      pane,
      "GET",
      `/me/mailFolders/drafts/messages?$filter=${filter}&$select=id`
    ).catch(() => ({ value: [] }));
    for (const d of leftovers.value) await graphFromPane(pane, "DELETE", `/me/messages/${d.id}`);
  }
});
