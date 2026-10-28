import {
  createNestablePublicClientApplication,
  type AuthenticationResult,
  type IPublicClientApplication,
} from "@azure/msal-browser";
import { ENTRA_AUTHORITY, ENTRA_CLIENT_ID, GRAPH_SCOPES } from "../config";

/* global Office */

let pcaPromise: Promise<IPublicClientApplication> | undefined;

function getPca(): Promise<IPublicClientApplication> {
  pcaPromise ??= createNestablePublicClientApplication({
    auth: { clientId: ENTRA_CLIENT_ID, authority: ENTRA_AUTHORITY },
  });
  return pcaPromise;
}

/** True when the host supports Nested App Authentication (NAA); otherwise only the popup works. */
export function isNaaSupported(): boolean {
  // requirements is undefined when the page runs outside an Office host.
  return Office.context?.requirements?.isSetSupported("NestedAppAuth", "1.1") ?? false;
}

/**
 * Graph access token for the signed-in Office user. Silent first (NAA brokers the Office account);
 * falls back to an MSAL popup, which must run from a user click.
 * Pattern from OfficeDev/Office-Add-in-samples Samples/auth/Outlook-Add-in-SSO-NAA.
 */
export async function getGraphToken(allowPopup = false): Promise<string> {
  const pca = await getPca();
  const request = { scopes: GRAPH_SCOPES };
  try {
    const result: AuthenticationResult = await pca.acquireTokenSilent(request);
    return result.accessToken;
  } catch (silentError) {
    if (!allowPopup) {
      throw silentError;
    }
    const result: AuthenticationResult = await pca.acquireTokenPopup(request);
    return result.accessToken;
  }
}
