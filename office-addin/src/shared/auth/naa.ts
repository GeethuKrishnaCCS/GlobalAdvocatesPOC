import {
  createNestablePublicClientApplication,
  LogLevel,
  type AuthenticationResult,
  type IPublicClientApplication,
  type SilentRequest,
} from "@azure/msal-browser";
import { ENTRA_AUTHORITY, ENTRA_CLIENT_ID, GRAPH_SCOPES } from "../config";

/* global Office, window, console, process */

/** Registered SPA redirect URI on the Entra app (path only; origin is the add-in host). */
const REDIRECT_PATH = "/taskpane.html";
/** Query flag that turns taskpane.html into the MSAL popup-relay page (see index.tsx). */
export const AUTH_RELAY_PARAM = "authRelay";

let pcaPromise: Promise<IPublicClientApplication> | undefined;

function getPca(): Promise<IPublicClientApplication> {
  pcaPromise ??= createNestablePublicClientApplication({
    auth: {
      clientId: ENTRA_CLIENT_ID,
      authority: ENTRA_AUTHORITY,
      // Must match the registered URI exactly; the pane URL carries ?view=…
      redirectUri: `${window.location.origin}${REDIRECT_PATH}`,
      // No NAA bridge (e.g. Office on the web dev sideload): the pane is a partitioned cross-site iframe, so the popup
      // result must come back through a first-party relay page. Ignored when NAA brokers the token.
      popupRelayUri: `${window.location.origin}${REDIRECT_PATH}?${AUTH_RELAY_PARAM}=1`,
    },
    cache: { cacheLocation: "localStorage" },
    // Dev builds: MSAL trace in the pane console, which the E2E harness captures (no PII).
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
  });
  return pcaPromise;
}

/** True when the host supports Nested App Authentication (NAA); otherwise only the popup works. */
export function isNaaSupported(): boolean {
  // requirements is undefined when the page runs outside an Office host.
  return Office.context?.requirements?.isSetSupported("NestedAppAuth", "1.1") ?? false;
}

/** Signed-in Office user's UPN; Word/Excel/PowerPoint on the web need it for ssoSilent. */
async function getLoginHint(): Promise<string | undefined> {
  try {
    return (await Office.auth.getAuthContext())?.userPrincipalName || undefined;
  } catch {
    return undefined;
  }
}

/** Token from the MSAL cache for an account already signed in here (e.g. by an earlier popup), else undefined. */
async function tryCachedToken(
  pca: IPublicClientApplication,
  loginHint: string | undefined
): Promise<string | undefined> {
  const accounts = pca.getAllAccounts();
  const account =
    accounts.find((a) => a.username.toLowerCase() === loginHint?.toLowerCase()) ?? accounts[0];
  if (!account) return undefined;
  try {
    return (await pca.acquireTokenSilent({ scopes: GRAPH_SCOPES, account })).accessToken;
  } catch {
    return undefined;
  }
}

/**
 * Graph access token for the signed-in Office user. Cache first, then silent SSO (ssoSilent with the Office login
 * hint, brokered by NAA where available); with `allowPopup` (from a user click) an MSAL popup replaces the silent step.
 * Pattern from OfficeDev/Office-Add-in-samples Samples/auth/Office-Add-in-SSO-NAA (authConfig.ts).
 */
export async function getGraphToken(allowPopup = false): Promise<string> {
  const pca = await getPca();
  const loginHint = await getLoginHint();
  const cached = await tryCachedToken(pca, loginHint);
  if (cached) return cached;
  if (allowPopup) {
    // Popup straight away: retrying silent SSO first (it already failed on load) outlasts the click's user activation,
    // so the browser blocks the popup (popup_window_error).
    return (await pca.acquireTokenPopup({ scopes: GRAPH_SCOPES, loginHint })).accessToken;
  }
  const result: AuthenticationResult = loginHint
    ? await pca.ssoSilent({ scopes: GRAPH_SCOPES, loginHint })
    : await pca.acquireTokenSilent({ scopes: GRAPH_SCOPES } as SilentRequest);
  return result.accessToken;
}
