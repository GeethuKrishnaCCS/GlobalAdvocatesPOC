import {
  createNestablePublicClientApplication,
  type AuthenticationResult,
  type IPublicClientApplication,
  type SilentRequest,
} from "@azure/msal-browser";
import { ENTRA_AUTHORITY, ENTRA_CLIENT_ID, GRAPH_SCOPES } from "../config";

/* global Office, window */

/** Registered SPA redirect URI on the Entra app (path only; origin is the add-in host). */
const REDIRECT_PATH = "/taskpane.html";

let pcaPromise: Promise<IPublicClientApplication> | undefined;

function getPca(): Promise<IPublicClientApplication> {
  pcaPromise ??= createNestablePublicClientApplication({
    auth: {
      clientId: ENTRA_CLIENT_ID,
      authority: ENTRA_AUTHORITY,
      // Must match the registered URI exactly; the pane URL carries ?view=…
      redirectUri: `${window.location.origin}${REDIRECT_PATH}`,
    },
    cache: { cacheLocation: "localStorage" },
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

/**
 * Graph access token for the signed-in Office user. Silent first (ssoSilent with the Office login hint,
 * else acquireTokenSilent); falls back to an MSAL popup, which must run from a user click.
 * Pattern from OfficeDev/Office-Add-in-samples Samples/auth/Office-Add-in-SSO-NAA (authConfig.ts).
 */
export async function getGraphToken(allowPopup = false): Promise<string> {
  const pca = await getPca();
  const loginHint = await getLoginHint();
  try {
    const result: AuthenticationResult = loginHint
      ? await pca.ssoSilent({ scopes: GRAPH_SCOPES, loginHint })
      : await pca.acquireTokenSilent({ scopes: GRAPH_SCOPES } as SilentRequest);
    return result.accessToken;
  } catch (silentError) {
    if (!allowPopup) {
      throw silentError;
    }
    const result: AuthenticationResult = await pca.acquireTokenPopup({
      scopes: GRAPH_SCOPES,
      loginHint,
    });
    return result.accessToken;
  }
}
