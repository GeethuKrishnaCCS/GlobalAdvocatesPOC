import { LogLevel, type Configuration } from "@azure/msal-browser";
import { ENTRA_AUTHORITY, ENTRA_CLIENT_ID } from "../config";

/* global window, console, process */

/** Registered SPA redirect URI on the Entra app (path only; origin is the add-in host). */
export const REDIRECT_PATH = "/taskpane.html";

/** MSAL settings shared by the pane (NAA) and the sign-in dialog (standard redirect). */
export function msalConfig(): Configuration {
  return {
    auth: {
      clientId: ENTRA_CLIENT_ID,
      authority: ENTRA_AUTHORITY,
      // Must match the registered URI exactly; the pane URL carries ?view=…
      redirectUri: `${window.location.origin}${REDIRECT_PATH}`,
    },
    cache: { cacheLocation: "localStorage" },
    // Dev builds: MSAL trace in the console, which the E2E harness captures (no PII).
    system:
      process.env.NODE_ENV === "production"
        ? undefined
        : {
            loggerOptions: {
              logLevel: LogLevel.Verbose,
              piiLoggingEnabled: false,
              loggerCallback: (_level, message) => console.log(`[msal] ${message}`),
            },
          },
  };
}
