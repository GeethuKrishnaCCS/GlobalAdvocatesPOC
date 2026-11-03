import { expect, test, type Frame } from "@playwright/test";
import {
  captureConsole,
  docUrlFor,
  graphFromPane,
  openPaneFor,
  saveRunScreenshot,
  signIn,
  type OfficeWebHost,
} from "./sideload";

/* global Buffer */

/**
 * Test targets on ccsdev01 (ADVTest marker): matter site, library folder, saved file name.
 * Each run saves a new test-marked file and deletes it (and its PDF) at the end; they go to the site recycle bin.
 */
const MATTER_ID = "advtestmatter01";
const MATTER_REF = "ADVTestMatter01";
const FOLDER_ID = "advtest-foldera";
const FOLDER = "ADVTest_FolderA";
const FILE_BASE = "ADVTest_E2E_SaveToMatter";
const SAVE_TIMEOUT = 120_000;
/** Word runs from the dev server; Excel and PowerPoint through the deployed add-in (ADV_E2E_HOSTED_ORIGIN). */
const HOSTS: OfficeWebHost[] = ["word", "excel", "powerpoint"];
const EXTENSIONS: Record<OfficeWebHost, string> = { word: ".docx", excel: ".xlsx", powerpoint: ".pptx" };


test.afterEach(async ({ page }, testInfo) => {
  if (testInfo.status !== "skipped") await saveRunScreenshot(page, testInfo);
});

async function pick(pane: Frame, dropdownTestId: string, option: string): Promise<void> {
  await pane.getByTestId(dropdownTestId).click();
  await pane.getByRole("option", { name: option }).click();
}

interface IFields {
  DocType?: string;
  AdvDirection?: string;
  _ExtendedDescription?: string;
}

/** Deletes a saved test file by its web URL (moves it to the site recycle bin). */
async function deleteByUrl(pane: Frame, webUrl: string): Promise<void> {
  const share = "u!" + Buffer.from(webUrl).toString("base64").replace(/=+$/, "").replace(/\//g, "_").replace(/\+/g, "-");
  const item = await graphFromPane<{ id: string; parentReference: { driveId: string } }>(
    pane,
    "GET",
    `/shares/${share}/driveItem?$select=id,parentReference`
  );
  await graphFromPane(pane, "DELETE", `/drives/${item.parentReference.driveId}/items/${item.id}`);
}

/** Reads the saved file's library fields with the pane's own Graph token (same delegated access as the user). */
async function fieldsOf(pane: Frame, linkTestId: string): Promise<IFields> {
  const webUrl = await pane.getByTestId(linkTestId).getAttribute("href");
  return pane.evaluate(async (url) => {
    const token = JSON.parse(localStorage.getItem("adv.dialogToken") ?? "{}").accessToken;
    const share = "u!" + btoa(url!).replace(/=+$/, "").replace(/\//g, "_").replace(/\+/g, "-");
    const response = await fetch(
      `https://graph.microsoft.com/v1.0/shares/${share}/driveItem/listItem/fields?$select=DocType,AdvDirection,_ExtendedDescription`,
      { headers: { Authorization: `Bearer ${token}` } }
    );
    return (await response.json()) as IFields;
  }, webUrl);
}

for (const host of HOSTS) {
  const ext = EXTENSIONS[host];
test(`${host}: save to matter, then save again as a new version with a PDF copy`, async ({ page }, testInfo) => {
  const docUrl = docUrlFor(host);
  test.skip(!docUrl, `no test document configured for ${host} (e2e/.env)`);
  test.setTimeout(5 * 60_000);
  captureConsole(page, testInfo);

  const pane = await openPaneFor(page, host, docUrl!, "Save to Matter");
  await signIn(page, pane);

  // 1. Matter → 2. Folder
  await pane.getByTestId(`matter-step-item-${MATTER_ID}`).click();
  await pane.getByTestId(`folder-step-item-${FOLDER_ID}`).click();
  await expect(pane.getByTestId("folder-step-path")).toContainText(FOLDER);
  await pane.getByTestId("save-to-matter-next").click();

  // 3. Details: defaults from the open document, name rules, typed extension not doubled.
  await expect(pane.getByTestId("details-step-location-path")).toHaveText(`${MATTER_REF} › ${FOLDER}`);
  await expect(pane.getByTestId("details-step-file-name")).toHaveValue(/^ADVTest_/);
  await pane.getByTestId("details-step-file-name").fill("CON");
  await expect(pane.getByTestId("details-step-file-name-field")).toContainText("reserved name");
  await expect(pane.getByTestId("save-to-matter-submit")).toBeDisabled();
  // A fresh file per run, so the first save is checked as 1.0; deleted again at the end (ADVTest marker).
  const fileBase = `${FILE_BASE}_${host}_${Date.now()}`;
  await pane.getByTestId("details-step-file-name").fill(`${fileBase}${ext}`);
  await expect(pane.getByTestId("save-to-matter-summary")).toHaveText(`${fileBase}${ext} → ${MATTER_REF} › ${FOLDER}`);

  const description = `E2E ${new Date().toISOString()}`;
  await pick(pane, "details-step-doc-type", "Correspondence");
  await pick(pane, "details-step-direction", "Inward");
  await pane.getByTestId("details-step-description").fill(description);
  await page.screenshot({ path: testInfo.outputPath("details-step.png") });
  await pane.getByTestId("save-to-matter-submit").click();

  try {
    await expect(pane.getByTestId("save-result-document-link")).toHaveText(`${fileBase}${ext}`, {
      timeout: SAVE_TIMEOUT,
    });
    // Properties go onto the uploaded version: first save 1.0, not 2.0 (needs the SharePoint scope, see config).
    await expect(pane.getByTestId("save-result-document-version")).toHaveText("Version 1.0");
    expect(await fieldsOf(pane, "save-result-document-link")).toMatchObject({
      DocType: "Correspondence",
      AdvDirection: "Inward",
    });

    // Same name, same folder → exactly one new version; clearing a choice clears the column; plus a PDF copy.
    await pane.getByTestId("save-result-save-another").click();
    await pane.getByTestId("details-step-direction-field").getByRole("button", { name: /clear/i }).click();
    await expect(pane.getByTestId("details-step-direction")).toHaveText(/Select/);
    await pane.getByTestId("details-step-pdf-copy").check();
    await pane.getByTestId("save-to-matter-submit").click();
    await expect(pane.getByTestId("save-result-pdf-link")).toHaveText(`${fileBase}.pdf`, { timeout: SAVE_TIMEOUT });
    await expect(pane.getByTestId("save-result-document-version")).toHaveText("Version 2.0");
    await expect(pane.getByTestId("save-result-pdf-version")).toHaveText("Version 1.0");
    await expect(pane.getByTestId("save-to-matter-error")).toHaveCount(0);

    const fields = await fieldsOf(pane, "save-result-document-link");
    expect(fields.DocType).toBe("Correspondence");
    expect(fields.AdvDirection ?? null).toBeNull();
    expect(fields._ExtendedDescription).toContain("E2E ");
    expect((await fieldsOf(pane, "save-result-pdf-link")).DocType).toBe("Correspondence");
  } finally {
    for (const link of ["save-result-document-link", "save-result-pdf-link"]) {
      const href = await pane.getByTestId(link).getAttribute("href", { timeout: 1_000 }).catch(() => null);
      if (href) await deleteByUrl(pane, href).catch(() => undefined);
    }
  }
});
}
