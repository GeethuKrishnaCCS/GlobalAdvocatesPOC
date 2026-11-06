import { defineConfig } from "@playwright/test";
import { existsSync } from "node:fs";

/* global process */

/** Signed-in Microsoft 365 session, written once by `npm run e2e:login` (user signs in; never automated). */
export const STORAGE_STATE = ".auth/office-web.json";
const DEV_SERVER_URL = "https://localhost:3000/taskpane.html";

if (existsSync("e2e/.env")) {
  process.loadEnvFile("e2e/.env");
}

export default defineConfig({
  testDir: "e2e",
  outputDir: "e2e/out",
  timeout: 120_000,
  expect: { timeout: 30_000 },
  retries: 0,
  workers: 1,
  reporter: [["list"]],
  use: {
    channel: "msedge",
    ignoreHTTPSErrors: true,
    screenshot: "only-on-failure",
    trace: "retain-on-failure",
    launchOptions: {
      // Keep the popup blocker on, like a user's browser: sign-in popups must open within the click's user activation.
      ignoreDefaultArgs: ["--disable-popup-blocking"],
      // Office on the web loads the pane from https://localhost:3000; Local Network Access prompts would block icons/pane.
      args: ["--disable-features=LocalNetworkAccessChecks,BlockInsecurePrivateNetworkRequests,PrivateNetworkAccessSendPreflights"],
    },
  },
  // ADV_E2E_NO_DEV_SERVER=1: production-only runs (hosted add-in), proving nothing depends on localhost.
  webServer: process.env.ADV_E2E_NO_DEV_SERVER ? undefined : {
    command: "npm run dev-server",
    url: DEV_SERVER_URL,
    ignoreHTTPSErrors: true,
    reuseExistingServer: true,
    timeout: 180_000,
  },
  projects: [
    // Pure helpers (no browser): npm run test:unit
    { name: "unit", testMatch: /.*\.unit\.ts/ },
    { name: "login", testMatch: /login\.setup\.ts/, use: { headless: false } },
    // DEV copy upload for the E2E account (Teams personal custom app): npm run package:dev && npm run e2e:dev-app
    {
      name: "dev-app",
      testMatch: /dev-app\.setup\.ts/,
      use: { storageState: existsSync(STORAGE_STATE) ? STORAGE_STATE : undefined },
    },
    {
      name: "office-web",
      testMatch: /.*\.spec\.ts/,
      use: { storageState: existsSync(STORAGE_STATE) ? STORAGE_STATE : undefined },
    },
  ],
});
