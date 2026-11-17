import { expect, test } from "@playwright/test";
import { captureConsole, docUrlFor, openPane, saveRunScreenshot } from "./sideload";

/* global console, getComputedStyle, document, Office */

test.afterEach(async ({ page }, testInfo) => {
  if (testInfo.status !== "skipped") await saveRunScreenshot(page, testInfo);
});

/** Relative luminance of "rgb(r, g, b)". */
const luminance = (rgb: string): number => {
  const [r, g, b] = (rgb.match(/\d+/g) ?? ["255", "255", "255"]).slice(0, 3).map((v) => Number(v) / 255);
  return 0.2126 * r + 0.7152 * g + 0.0722 * b;
};

for (const scheme of ["light", "dark"] as const) {
  test.describe(`${scheme} colour scheme`, () => {
    test.use({ colorScheme: scheme });

    test(`word: task pane follows ${scheme} mode`, async ({ page }, testInfo) => {
      const docUrl = docUrlFor("word");
      test.skip(!docUrl, "no test document configured for word (e2e/.env)");
      captureConsole(page, testInfo);

      const pane = await openPane(page, docUrl!, "Save to Matter");
      await expect(pane.getByTestId("app-title")).toBeVisible();
      const seen = await pane.evaluate(() => ({
        officeBackground: Office.context.officeTheme?.bodyBackgroundColor ?? null,
        paneBackground: getComputedStyle(document.querySelector("[data-testid=pane-theme]")!).backgroundColor,
        text: getComputedStyle(document.querySelector("[data-testid=app-title]")!).color,
      }));
      console.log(`[theme ${scheme}] ${JSON.stringify(seen)}`);

      // Office's own theme wins when it reports one; otherwise the browser scheme decides.
      const officeDark = seen.officeBackground ? luminance(hexToRgb(seen.officeBackground)) < 0.4 : undefined;
      const expectDark = officeDark ?? scheme === "dark";
      expect(luminance(seen.paneBackground) < 0.4, "pane background dark").toBe(expectDark);
      expect(luminance(seen.text) < 0.4, "text dark on light / light on dark").toBe(!expectDark);
    });
  });
}

function hexToRgb(hex: string): string {
  const h = hex.replace("#", "");
  return `rgb(${parseInt(h.slice(0, 2), 16)}, ${parseInt(h.slice(2, 4), 16)}, ${parseInt(h.slice(4, 6), 16)})`;
}
