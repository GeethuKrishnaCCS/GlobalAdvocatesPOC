import { expect, test, type Frame } from "@playwright/test";
import { captureConsole, docUrlFor, openPane, saveRunScreenshot, signIn } from "./sideload";

/**
 * Test targets on ccsdev01 (ADVTest marker): matter site, library folder, saved file name.
 * One fixed name on purpose: every run adds versions to the same test file (cleanup: one item, see task.md ledger).
 */
const MATTER_ID = "advtestmatter01";
const MATTER_REF = "ADVTestMatter01";
const FOLDER_ID = "advtest-foldera";
const FOLDER = "ADVTest_FolderA";
const FILE_BASE = "ADVTest_E2E_SaveToMatter";
const SAVE_TIMEOUT = 120_000;

test.afterEach(async ({ page }, testInfo) => {
  if (testInfo.status !== "skipped") await saveRunScreenshot(page, testInfo);
});

async function pick(pane: Frame, dropdownTestId: string, option: string): Promise<void> {
  await pane.getByTestId(dropdownTestId).click();
  await pane.getByRole("option", { name: option }).click();
}

/** "3.0" → 3, "0.2" → 0.2. */
const versionOf = async (pane: Frame, testId: string): Promise<number> =>
  parseFloat(((await pane.getByTestId(testId).textContent()) ?? "").replace(/[^\d.]/g, ""));

interface IFields {
  DocType?: string;
  AdvDirection?: string;
  _ExtendedDescription?: string;
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

test("word: save to matter, then save again as a new version with a PDF copy", async ({ page }, testInfo) => {
  const docUrl = docUrlFor("word");
  test.skip(!docUrl, "no test document configured for word (e2e/.env)");
  test.setTimeout(5 * 60_000);
  captureConsole(page, testInfo);

  const pane = await openPane(page, docUrl!, "Save to Matter");
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
  await pane.getByTestId("details-step-file-name").fill(`${FILE_BASE}.docx`);
  await expect(pane.getByTestId("save-to-matter-summary")).toHaveText(`${FILE_BASE}.docx → ${MATTER_REF} › ${FOLDER}`);

  const description = `E2E ${new Date().toISOString()}`;
  await pick(pane, "details-step-doc-type", "Correspondence");
  await pick(pane, "details-step-direction", "Inward");
  await pane.getByTestId("details-step-description").fill(description);
  await page.screenshot({ path: testInfo.outputPath("details-step.png") });
  await pane.getByTestId("save-to-matter-submit").click();

  await expect(pane.getByTestId("save-result-document-link")).toHaveText(`${FILE_BASE}.docx`, {
    timeout: SAVE_TIMEOUT,
  });
  const firstVersion = await versionOf(pane, "save-result-document-version");
  expect(firstVersion).toBeGreaterThan(0);
  expect(await fieldsOf(pane, "save-result-document-link")).toMatchObject({
    DocType: "Correspondence",
    AdvDirection: "Inward",
  });

  // Same name, same folder → new version of the same file; clearing a choice clears the column; plus a PDF copy.
  await pane.getByTestId("save-result-save-another").click();
  await pane.getByTestId("details-step-direction-field").getByRole("button", { name: /clear/i }).click();
  await expect(pane.getByTestId("details-step-direction")).toHaveText(/Select/);
  await pane.getByTestId("details-step-pdf-copy").check();
  await pane.getByTestId("save-to-matter-submit").click();
  await expect(pane.getByTestId("save-result-pdf-link")).toHaveText(`${FILE_BASE}.pdf`, { timeout: SAVE_TIMEOUT });
  expect(await versionOf(pane, "save-result-document-version")).toBeGreaterThan(firstVersion);
  await expect(pane.getByTestId("save-to-matter-error")).toHaveCount(0);

  const fields = await fieldsOf(pane, "save-result-document-link");
  expect(fields.DocType).toBe("Correspondence");
  expect(fields.AdvDirection ?? null).toBeNull();
  expect(fields._ExtendedDescription).toContain("E2E ");
  expect((await fieldsOf(pane, "save-result-pdf-link")).DocType).toBe("Correspondence");
});
