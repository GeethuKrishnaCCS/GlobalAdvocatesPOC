import { expect, test } from "@playwright/test";
import { captureConsole, docUrlFor, openPane, type OfficeWebHost } from "./sideload";

/** Office.HostType value the pane shows in `app-host`. */
const HOST_LABEL: Record<OfficeWebHost, string> = { word: "Word", excel: "Excel", powerpoint: "PowerPoint" };

for (const host of Object.keys(HOST_LABEL) as OfficeWebHost[]) {
  test(`${host}: ribbon opens pane and NAA signs in`, async ({ page }, testInfo) => {
    const docUrl = docUrlFor(host);
    test.skip(!docUrl, `no test document configured for ${host} (e2e/.env)`);
    captureConsole(page, testInfo);

    const pane = await openPane(page, docUrl!, "Save to Matter");

    await expect(pane.getByTestId("app-title")).toHaveText("Save to Matter");
    await expect(pane.getByTestId("app-host")).toContainText(HOST_LABEL[host]);
    // Silent NAA should sign in without a click; on failure the pane's error text is in the assertion output.
    await expect(pane.getByTestId("app-user").or(pane.getByTestId("app-error"))).toBeVisible();
    await expect(pane.getByTestId("app-user")).toContainText("@");
  });
}
