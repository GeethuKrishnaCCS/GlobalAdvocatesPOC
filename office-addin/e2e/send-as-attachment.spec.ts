import { expect, test, type Frame } from "@playwright/test";
import { FILE_TO_FOLDER_PROPERTY_ID as PROPERTY_ID } from "../src/shared/config";
import {
  captureConsole,
  expectFooterVisible,
  docUrlFor,
  graphFromPane,
  openPaneFor,
  saveRunScreenshot,
  signIn,
  type OfficeWebHost,
} from "./sideload";

/** Matter the Word test document lives in (pre-selected) and the draft's test-marked names. */
const MATTER_REF = "ADVTestMatter01";
const PDF_NAME = "ADVTest_E2E_Send_PDF";
const CREATE_TIMEOUT = 120_000;
/** Word runs from the dev server; Excel and PowerPoint through the deployed add-in (ADV_E2E_HOSTED_ORIGIN). */
const HOSTS: OfficeWebHost[] = ["word", "excel", "powerpoint"];
const EXTENSIONS: Record<OfficeWebHost, string> = {
  word: ".docx",
  excel: ".xlsx",
  powerpoint: ".pptx",
};

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

/** Test drafts with this exact (ADVTest document) subject; they move to Deleted Items. */
async function deleteDrafts(pane: Frame, subject: string): Promise<void> {
  const filter = encodeURIComponent(`subject eq '${subject.replace(/'/g, "''")}'`);
  const drafts = await graphFromPane<{ value: { id: string }[] }>(
    pane,
    "GET",
    `/me/mailFolders/drafts/messages?$filter=${filter}&$select=id`
  ).catch(() => ({ value: [] }));
  for (const d of drafts.value) await graphFromPane(pane, "DELETE", `/me/messages/${d.id}`);
}

for (const host of HOSTS) {
  const ext = EXTENSIONS[host];
  test(`${host}: send as attachment opens a tagged Outlook draft with the document + PDF`, async ({
    page,
  }, testInfo) => {
    const docUrl = docUrlFor(host);
    test.skip(!docUrl, `no test document configured for ${host} (e2e/.env)`);
    const NATIVE_NAME = `ADVTest_E2E_Send_${host}`;
    test.setTimeout(5 * 60_000);
    captureConsole(page, testInfo);

    const pane = await openPaneFor(page, host, docUrl!, "Send as Attachment");
    await signIn(page, pane);

    // Matter of the open document is pre-selected; its Correspondence › Matter folder enables filing.
    await expect(pane.getByTestId("send-matter-picker")).toContainText(MATTER_REF, {
      timeout: 60_000,
    });
    await expect(pane.getByTestId("send-matter-file-email")).toBeEnabled({ timeout: 30_000 });
    // Subject is the tagged document name; recipients and message are left to Outlook (no fields in the pane).
    const docBase = ((await pane.getByTestId("send-file-name").textContent()) ?? "").replace(
      /.[^.]+$/,
      ""
    );
    const tagged = `[${MATTER_REF}] ${docBase}`;
    expect(docBase, "ADVTest document (cleanup deletes drafts by this subject)").toMatch(
      /^ADVTest/
    );
    await expect(pane.locator('[data-testid^="send-message"]')).toHaveCount(0);
    await deleteDrafts(pane, tagged);

    await pane.getByTestId("send-attach-format-both").click();
    await pane.getByTestId("send-attach-name-native").fill(NATIVE_NAME);
    await pane.getByTestId("send-attach-name-pdf").fill(`${PDF_NAME}.pdf`);
    await pane.getByTestId("send-matter-file-email").check();
    await expect(pane.getByTestId("send-summary")).toHaveText(`2 attachments · ${tagged}`);
    await expectFooterVisible(pane, "send-footer");
    await page.screenshot({ path: testInfo.outputPath("send-form.png") });

    // From here a draft may exist: whatever fails, the finally removes this run's test draft.
    try {
      // The window opens inside the click (popup blocker on) and lands on the draft once it exists.
      const popup = page.context().waitForEvent("page", { timeout: 30_000 });
      await pane.getByTestId("send-create").click();
      const outlook = await popup;
      await expect
        .poll(() => outlook.url(), { timeout: CREATE_TIMEOUT })
        .toMatch(/^https:\/\/outlook\.cloud\.microsoft\/mail\/deeplink\/compose\//);
      await expect(pane.getByTestId("send-result-message")).toContainText(
        `Opened in Outlook: ${tagged}`
      );
      await expect(pane.getByTestId("send-result-attachment-name")).toHaveText([
        `${NATIVE_NAME}${ext}`,
        `${PDF_NAME}.pdf`,
      ]);
      // Compose form of the draft (Send button) with the tagged subject: the user only adds recipients and message.
      await expect(outlook.getByRole("button", { name: "Send", exact: true }).first()).toBeVisible({
        timeout: 90_000,
      });
      await expect(outlook.getByRole("textbox", { name: /subject/i }).first()).toHaveValue(tagged);
      await outlook.screenshot({ path: testInfo.outputPath("outlook.png") });
      await outlook.close();

      const draft = await findDraft(pane, tagged);
      expect(draft, "draft in Drafts").toBeDefined();
      expect(draft!.categories).toEqual([`Matter: ${MATTER_REF}`]);
      expect(draft!.toRecipients).toEqual([]);
      expect(draft!.attachments.map((a) => a.name).sort()).toEqual(
        [`${NATIVE_NAME}${ext}`, `${PDF_NAME}.pdf`].sort()
      );
      expect(draft!.attachments.every((a) => a.size > 0)).toBe(true);
      expect(draft!.singleValueExtendedProperties?.[0]?.value).toMatch(/^b!.+|.+$/);
    } finally {
      await deleteDrafts(pane, tagged);
    }
  });
}
