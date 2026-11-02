import * as React from "react";
import { createRoot } from "react-dom/client";
import { broadcastResponseToMainFrame } from "@azure/msal-browser/redirect-bridge";
import { runPopupRelay } from "@azure/msal-browser/popup-relay";
import App from "./components/App";
import { Body1, Button, FluentProvider, webLightTheme } from "@fluentui/react-components";
import { PaneView } from "../shared/config";
import { AUTH_RELAY_PARAM } from "../shared/auth/naa";

/* global document, Office, window, HTMLElement */

const rootElement: HTMLElement | null = document.getElementById("container");
const root = rootElement ? createRoot(rootElement) : undefined;
const AUTHORITY_ORIGINS = ["https://login.microsoftonline.com"];

function viewFromUrl(): PaneView | undefined {
  const value = new URLSearchParams(window.location.search).get("view");
  return Object.values(PaneView).find((v) => v === value);
}

/** This page is also the MSAL popup redirect target (no NAA bridge, e.g. web dev sideload): hand the result back. */
function isAuthPopupRedirect(): boolean {
  return !!window.opener && /[#&](code|error)=/.test(window.location.hash);
}

/** First-party relay page MSAL opens from the partitioned pane iframe (auth.popupRelayUri). */
function isAuthRelay(): boolean {
  return new URLSearchParams(window.location.search).has(AUTH_RELAY_PARAM);
}

if (isAuthRelay()) {
  // runPopupRelay opens the sign-in popup, so it needs this click to get past popup blockers.
  root?.render(
    <FluentProvider theme={webLightTheme}>
      <Body1 data-testid="auth-relay-text">Continue to sign in to Global Advocate.</Body1>{" "}
      <Button
        appearance="primary"
        data-testid="auth-relay-continue"
        onClick={() => runPopupRelay({ allowedAuthorityOrigins: AUTHORITY_ORIGINS })}
      >
        Continue
      </Button>
    </FluentProvider>
  );
} else if (isAuthPopupRedirect()) {
  void broadcastResponseToMainFrame();
} else {
  /* Render application after Office initializes */
  Office.onReady((info) => {
    root?.render(
      <FluentProvider theme={webLightTheme}>
        <App host={info.host} view={viewFromUrl()} />
      </FluentProvider>
    );
  });
}
