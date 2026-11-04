import {
  createNestablePublicClientApplication,
  type AuthenticationResult,
  type IPublicClientApplication,
  type SilentRequest,
} from "@azure/msal-browser";
import { GRAPH_SCOPES, SHAREPOINT_SCOPES } from "../config";
import { openAuthDialog, type IDialogToken } from "./authDialog";
import { msalConfig } from "./msalConfig";

/* global Office, window, localStorage */

/** Pane-side copy of a token the sign-in dialog returned (the pane's storage is partitioned from the dialog's). */
const DIALOG_TOKEN_KEY = "adv.dialogToken";
/** Treat a token this close to expiry as expired, so a Graph call never starts with one about to lapse. */
const EXPIRY_BUFFER_MS = 5 * 60 * 1000;

let pcaPromise: Promise<IPublicClientApplication> | undefined;

function getPca(): Promise<IPublicClientApplication> {
  pcaPromise ??= createNestablePublicClientApplication(msalConfig());
  return pcaPromise;
}

/**
 * True when Office brokers tokens through Nested App Auth. The requirement set alone is not enough: Office on the web
 * with a URL-param dev sideload reports NestedAppAuth 1.1 but injects no bridge.
 */
export function isNaaSupported(): boolean {
  const bridge = (window as unknown as { nestedAppAuthBridge?: unknown }).nestedAppAuthBridge;
  // requirements is undefined when the page runs outside an Office host.
  return (
    !!bridge && (Office.context?.requirements?.isSetSupported("NestedAppAuth", "1.1") ?? false)
  );
}

let loginHintPromise: Promise<string | undefined> | undefined;

/**
 * Signed-in Office user's UPN; used as login hint so sign-in skips the account picker. Asked once per pane load (the
 * silent attempt on load primes it), so the Sign in click does not spend its user activation waiting on the host.
 */
function getLoginHint(): Promise<string | undefined> {
  loginHintPromise ??= Office.auth.getAuthContext().then(
    (context) => context?.userPrincipalName || undefined,
    () => undefined
  );
  return loginHintPromise;
}

/** Token from the MSAL cache for an account already signed in here, else undefined. */
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

/** In-memory copy, so a pane whose storage is blocked still works until it is reloaded. */
let memoryToken: IDialogToken | undefined;

/** Stored dialog token if still valid and issued to the current Office user (the pane's storage is per browser). */
function readDialogToken(loginHint: string | undefined): IDialogToken | undefined {
  let token = memoryToken;
  if (!token) {
    try {
      token = JSON.parse(localStorage.getItem(DIALOG_TOKEN_KEY) ?? "null") ?? undefined;
    } catch {
      token = undefined;
    }
  }
  if (!token || token.expiresOn - EXPIRY_BUFFER_MS <= Date.now()) return undefined;
  if (loginHint && token.username.toLowerCase() !== loginHint.toLowerCase()) return undefined;
  return token;
}

function saveDialogToken(token: IDialogToken): void {
  memoryToken = token;
  try {
    localStorage.setItem(DIALOG_TOKEN_KEY, JSON.stringify(token));
  } catch {
    // Storage blocked: the in-memory copy serves until the pane reloads.
  }
}

/** Forgets the dialog token, e.g. after Graph rejected it with 401. */
export function clearDialogToken(): void {
  memoryToken = undefined;
  try {
    localStorage.removeItem(DIALOG_TOKEN_KEY);
  } catch {
    // Nothing stored.
  }
}

/**
 * Graph access token for the signed-in Office user.
 * - NAA bridge present: MSAL cache → ssoSilent (Office login hint) → with `allowPopup`, a brokered popup.
 * - No bridge (Office on the web dev sideload): the token from the last sign-in dialog until it expires; with
 *   `allowPopup`, the Office sign-in dialog. Silent SSO is skipped there: it cannot succeed in the partitioned pane and
 *   only delayed the Sign in button.
 * `allowPopup` must come from a user click, and nothing slow may run before the popup/dialog opens (user activation).
 */
export async function getGraphToken(allowPopup = false): Promise<string> {
  if (!isNaaSupported()) {
    const loginHint = await getLoginHint();
    const stored = readDialogToken(loginHint);
    if (stored) return stored.accessToken;
    if (!allowPopup) throw new Error("interaction_required: sign in to continue.");
    const token = await openAuthDialog(loginHint);
    saveDialogToken(token);
    return token.accessToken;
  }
  const pca = await getPca();
  const loginHint = await getLoginHint();
  const cached = await tryCachedToken(pca, loginHint);
  if (cached) return cached;
  const result: AuthenticationResult = allowPopup
    ? await pca.acquireTokenPopup({ scopes: GRAPH_SCOPES, loginHint })
    : loginHint
      ? await pca.ssoSilent({ scopes: GRAPH_SCOPES, loginHint })
      : await pca.acquireTokenSilent({ scopes: GRAPH_SCOPES } as SilentRequest);
  return result.accessToken;
}

const SHAREPOINT_CONSENT_REQUIRED =
  "sharepoint_consent_required: Global Advocate has no SharePoint permission for this account.";

/**
 * SharePoint REST token for the signed-in user (setting document properties without adding a version).
 * Cached token first; otherwise, from a click (`allowPopup`), the sign-in dialog or a brokered popup gets one, which
 * also repairs a sign-in made before the SharePoint scope was consented. Throws when the scope is not consented.
 * Like Graph: nothing slow may run before the dialog opens (user activation).
 */
export async function getSharePointToken(allowPopup = false): Promise<string> {
  const loginHint = await getLoginHint();
  if (!isNaaSupported()) {
    const valid = (t: IDialogToken | undefined): string | undefined =>
      t?.sharePoint && t.sharePoint.expiresOn - EXPIRY_BUFFER_MS > Date.now()
        ? t.sharePoint.accessToken
        : undefined;
    const cached = valid(readDialogToken(loginHint));
    if (cached) return cached;
    if (!allowPopup) throw new Error("interaction_required: sign in to continue.");
    const token = await openAuthDialog(loginHint);
    saveDialogToken(token);
    const fresh = valid(token);
    if (!fresh) throw new Error(SHAREPOINT_CONSENT_REQUIRED);
    return fresh;
  }
  const pca = await getPca();
  const account =
    pca.getAllAccounts().find((a) => a.username.toLowerCase() === loginHint?.toLowerCase()) ??
    pca.getAllAccounts()[0];
  try {
    return (await pca.acquireTokenSilent({ scopes: SHAREPOINT_SCOPES, account })).accessToken;
  } catch (silentError) {
    if (!allowPopup) throw silentError;
    try {
      return (await pca.acquireTokenPopup({ scopes: SHAREPOINT_SCOPES, loginHint })).accessToken;
    } catch (e) {
      throw /consent|AADSTS65001/i.test(String(e)) ? new Error(SHAREPOINT_CONSENT_REQUIRED) : e;
    }
  }
}
