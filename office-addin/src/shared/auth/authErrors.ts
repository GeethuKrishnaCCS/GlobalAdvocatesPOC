/** Entra (AADSTS) error codes users actually hit, in words. Unknown codes fall back to the raw message. */
const AADSTS_MESSAGES: Record<string, string> = {
  "50020":
    "This account is not a member of the organisation that owns Global Advocate. Use your work account.",
  "50058": "You are not signed in to Microsoft 365 in this browser. Sign in again.",
  "50076": "Multi-factor authentication is required. Complete it in the sign-in window.",
  "50079": "Your organisation requires you to set up multi-factor authentication first.",
  "50105": "Your account has not been given access to Global Advocate. Ask your administrator.",
  "53003": "A Conditional Access policy blocked the sign-in. Contact your administrator.",
  "65001":
    "Global Advocate needs permission to your data. Ask your administrator to grant consent.",
  "700016":
    "The Global Advocate app registration was not found in this tenant. Contact your administrator.",
  "90072":
    "This account is from another organisation and is not a guest here. Use your work account.",
};

/** MSAL / Office dialog error codes, in words. */
const CLIENT_MESSAGES: Record<string, string> = {
  user_cancelled: "Sign-in was cancelled. Select Sign in to try again.",
  popup_window_error:
    "The browser blocked the sign-in window. Allow pop-ups for this site, then try again.",
  empty_window_error:
    "The browser blocked the sign-in window. Allow pop-ups for this site, then try again.",
  timed_out: "Sign-in took too long. Select Sign in to try again.",
  interaction_required: "Sign in to continue.",
  sharepoint_consent_required:
    "Global Advocate needs SharePoint permission to set document properties. Ask your administrator to grant consent.",
  consent_required:
    "Global Advocate needs permission to your data. Ask your administrator to grant consent.",
  no_network_connectivity: "No network connection. Check your connection and try again.",
  // Office dialog API: 12009 user ignored the prompt, 12011 browser blocked the dialog.
  "dialog_error 12009": "The sign-in window was not allowed to open. Select Sign in and allow it.",
  "dialog_error 12011":
    "The browser blocked the sign-in window. Allow pop-ups for this site, then try again.",
};

/**
 * Readable message for a sign-in failure, with the Entra correlation ID when present (support needs it).
 * Pure: unit-testable without a browser.
 */
export function friendlyAuthError(error: unknown): string {
  const raw = error instanceof Error ? error.message : String(error);
  const aadsts = /AADSTS(\d+)/.exec(raw)?.[1];
  const clientKey = Object.keys(CLIENT_MESSAGES).find(
    (key) => raw.startsWith(key) || raw.includes(`${key}:`)
  );
  const text =
    (aadsts && AADSTS_MESSAGES[aadsts]) ?? (clientKey && CLIENT_MESSAGES[clientKey]) ?? raw;
  const correlationId = /Correlation ID: ([0-9a-f-]{36})/i.exec(raw)?.[1];
  return correlationId ? `${text} (Correlation ID: ${correlationId})` : text;
}
