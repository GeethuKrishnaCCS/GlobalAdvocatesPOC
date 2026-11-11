import { type Frame, type Page } from "@playwright/test";
import { MATTER_LIBRARY_NAME } from "../src/shared/config";
import { OUTLOOK_URL, graphFromPane, openOutlookPane, signIn } from "./sideload";

/* Test data helpers shared by the Outlook specs (ccsdev01, ADVTest-marked objects only). */

/** Test matter / folder on ccsdev01 and an existing ADVTest document to attach. */
export const MATTER = "ADVTestMatter01";
export const MATTER_SITE = "ccsdev01.sharepoint.com:/sites/ADVTestMatter01";
export const FOLDER = "ADVTest_FolderA";
export const FILING_FOLDER = "Correspondence/Matter";
export const ATTACH_DOC = "ADVTest_Order_0001.docx";
/** PR_MESSAGE_FLAGS = read, not unsent: a Graph-created message then shows as received mail, not a draft. */
export const RECEIVED_FLAGS = { id: "Integer 0x0E07", value: "1" };

export interface IItem {
  id: string;
  name: string;
  size: number;
}

export async function matterDriveId(pane: Frame): Promise<string> {
  const site = await graphFromPane<{ id: string }>(pane, "GET", `/sites/${MATTER_SITE}?$select=id`);
  const lists = await graphFromPane<{ value: { id: string; name: string }[] }>(
    pane,
    "GET",
    `/sites/${site.id}/lists?$select=id,name`
  );
  const list = lists.value.find((l) => l.name === MATTER_LIBRARY_NAME)!;
  return (
    await graphFromPane<{ id: string }>(
      pane,
      "GET",
      `/sites/${site.id}/lists/${list.id}/drive?$select=id`
    )
  ).id;
}

export async function folderId(pane: Frame, driveId: string, path: string): Promise<string> {
  return (
    await graphFromPane<{ id: string }>(pane, "GET", `/drives/${driveId}/root:/${path}?$select=id`)
  ).id;
}

/** Received-looking test email in the Inbox with two small attachments; returns its Graph id. */
export async function createTestEmail(
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
  const message = await graphFromPane<{ id: string }>(
    pane,
    "POST",
    "/me/mailFolders/inbox/messages",
    {
      subject,
      body: { contentType: "Text", content: "ADVTest E2E message for File to Matter." },
      inferenceClassification: "focused",
      attachments: [
        attachment(`ADVTest_E2E_File_${stamp}.txt`, "hello matter"),
        attachment(`ADVTest_E2E_File_${stamp}.csv`, "a,b\n1,2"),
      ],
      ...extra,
      singleValueExtendedProperties: [RECEIVED_FLAGS, ...props],
    }
  );
  return message.id;
}

export async function openEmail(page: Page, subject: string): Promise<void> {
  await page.goto(OUTLOOK_URL);
  await page
    .getByRole("option", { name: new RegExp(subject) })
    .first()
    .click({ timeout: 60_000 });
}

export async function deleteFiled(
  pane: Frame,
  driveId: string,
  parentId: string,
  stamp: string
): Promise<void> {
  const children = await graphFromPane<{ value: IItem[] }>(
    pane,
    "GET",
    `/drives/${driveId}/items/${parentId}/children?$select=id,name,size&$top=999`
  );
  for (const item of children.value.filter((c) => c.name.includes(stamp))) {
    await graphFromPane(pane, "DELETE", `/drives/${driveId}/items/${item.id}`);
  }
}

/** Pane on any Inbox email, signed in: used only to set up test data through Graph. */
export async function setupPane(page: Page): Promise<Frame> {
  await page.goto(OUTLOOK_URL);
  await page.getByRole("option").first().click({ timeout: 60_000 });
  const pane = await openOutlookPane(page, "More apps");
  await signIn(page, pane);
  // Test emails a failed run could not delete (its pane was gone): ADVTest-marked, ours.
  const filter = encodeURIComponent("startswith(subject,'ADVTest E2E ')");
  const stale = await graphFromPane<{ value: { id: string }[] }>(
    pane,
    "GET",
    `/me/mailFolders/inbox/messages?$filter=${filter}&$select=id`
  );
  for (const m of stale.value) await graphFromPane(pane, "DELETE", `/me/messages/${m.id}`);
  return pane;
}

