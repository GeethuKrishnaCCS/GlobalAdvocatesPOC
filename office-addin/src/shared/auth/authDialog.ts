import {
  createStandardPublicClientApplication,
  type AuthenticationResult,
} from "@azure/msal-browser";
import { GRAPH_SCOPES } from "../config";
import { msalConfig, REDIRECT_PATH } from "./msalConfig";

/* global Office, window, sessionStorage, URLSearchParams */

/** Query flag that turns taskpane.html into the sign-in dialog page. */
export const AUTH_DIALOG_PARAM = "authDialog";
const LOGIN_HINT_PARAM = "loginHint";
/** Survives the MSAL redirect round trip, which returns to the bare redirect URI (query lost). */
const DIALOG_MARK_KEY = "adv.authDialog";
/** Set just before the MSAL redirect; a return without a result then fails instead of redirecting again (loop). */
const REDIRECTED_KEY = "adv.authRedirected";
/** Office dialog event: the user closed the dialog window. */
const DIALOG_CLOSED_BY_USER = 12006;
/** Office dialog error: the browser blocked the window. */
const DIALOG_BLOCKED = 12011;

/** Token handed from the sign-in dialog to the pane. */
export interface IDialogToken {
  accessToken: string;
  expiresOn: number;
  username: string;
}

type DialogMessage = { type: "token"; token: IDialogToken } | { type: "error"; message: string };

/** True when this page load is the sign-in dialog (first load, or back from the identity provider). */
export function isAuthDialog(): boolean {
  if (new URLSearchParams(window.location.search).has(AUTH_DIALOG_PARAM)) return true;
  return (
    sessionStorage.getItem(DIALOG_MARK_KEY) === "1" &&
    /[#&](code|error)=/.test(window.location.hash)
  );
}

/**
 * Pane side: opens the Office sign-in dialog and resolves with the token it posts back. Call from a click: the window
 * opens straight away; if the browser still blocks it (12011), Office's own "allow window" prompt is shown instead.
 */
export function openAuthDialog(loginHint: string | undefined): Promise<IDialogToken> {
  const query = new URLSearchParams({ [AUTH_DIALOG_PARAM]: "1" });
  if (loginHint) query.set(LOGIN_HINT_PARAM, loginHint);
  const url = `${window.location.origin}${REDIRECT_PATH}?${query}`;

  return new Promise((resolve, reject) => {
    const open = (promptBeforeOpen: boolean): void =>
      Office.context.ui.displayDialogAsync(
        url,
        { height: 60, width: 30, promptBeforeOpen },
        (result) => {
          if (result.status === Office.AsyncResultStatus.Failed) {
            if (result.error.code === DIALOG_BLOCKED && !promptBeforeOpen) open(true);
            else reject(new Error(`dialog_error ${result.error.code}: ${result.error.message}`));
            return;
          }
          const dialog = result.value;
          dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
            const { message, origin } = arg as { message: string; origin: string | undefined };
            // Only our own page may hand over a token.
            if (origin && origin !== window.location.origin) return;
            dialog.close();
            try {
              const parsed = JSON.parse(message) as DialogMessage;
              if (parsed.type === "token") resolve(parsed.token);
              else reject(new Error(parsed.message));
            } catch {
              reject(new Error("dialog_error: unreadable sign-in response."));
            }
          });
          dialog.addEventHandler(Office.EventType.DialogEventReceived, (arg) => {
            const code = (arg as { error: number }).error;
            reject(
              new Error(
                code === DIALOG_CLOSED_BY_USER
                  ? "user_cancelled: sign-in window closed."
                  : `dialog_error ${code}`
              )
            );
          });
        }
      );
    open(false);
  });
}

function send(message: DialogMessage): void {
  Office.context.ui.messageParent(JSON.stringify(message), {
    targetOrigin: window.location.origin,
  });
}

function toToken(result: AuthenticationResult): IDialogToken {
  return {
    accessToken: result.accessToken,
    expiresOn: result.expiresOn?.getTime() ?? 0,
    username: result.account?.username ?? "",
  };
}

/**
 * Dialog side (runs after Office.onReady in the dialog window): MSAL redirect sign-in in a first-party window. Its
 * MSAL cache keeps the refresh token, so later sign-ins return without any prompt.
 */
export async function runAuthDialog(): Promise<void> {
  try {
    const params = new URLSearchParams(window.location.search);
    if (params.has(AUTH_DIALOG_PARAM)) {
      sessionStorage.setItem(DIALOG_MARK_KEY, "1");
      sessionStorage.setItem(LOGIN_HINT_PARAM, params.get(LOGIN_HINT_PARAM) ?? "");
    }
    const loginHint = sessionStorage.getItem(LOGIN_HINT_PARAM) || undefined;
    const pca = await createStandardPublicClientApplication(msalConfig());

    const redirected = await pca.handleRedirectPromise();
    if (redirected) {
      send({ type: "token", token: toToken(redirected) });
      return;
    }
    if (sessionStorage.getItem(REDIRECTED_KEY) === "1") {
      send({
        type: "error",
        message: "redirect_failed: sign-in did not complete. Select Sign in to try again.",
      });
      return;
    }
    const account =
      pca.getAllAccounts().find((a) => a.username.toLowerCase() === loginHint?.toLowerCase()) ??
      (loginHint ? undefined : pca.getAllAccounts()[0]);
    if (account) {
      try {
        send({
          type: "token",
          token: toToken(await pca.acquireTokenSilent({ scopes: GRAPH_SCOPES, account })),
        });
        return;
      } catch {
        // Refresh token expired or consent changed: fall through to the interactive redirect.
      }
    }
    sessionStorage.setItem(REDIRECTED_KEY, "1");
    await pca.acquireTokenRedirect({ scopes: GRAPH_SCOPES, loginHint });
  } catch (e) {
    send({ type: "error", message: (e as Error).message });
  }
}
