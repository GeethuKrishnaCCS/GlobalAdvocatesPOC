import * as React from "react";
import { createRoot } from "react-dom/client";
import { broadcastResponseToMainFrame } from "@azure/msal-browser/redirect-bridge";
import App from "./components/App";
import { Body1, FluentProvider } from "@fluentui/react-components";
import { PaneView } from "../shared/config";
import { isAuthDialog, runAuthDialog } from "../shared/auth/authDialog";
import { getGraphToken } from "../shared/auth/naa";
import { ErrorBoundary } from "../shared/components/ErrorBoundary";
import { useOfficeTheme } from "../shared/hooks/useOfficeTheme";

/* global document, Office, window, HTMLElement */

const rootElement: HTMLElement | null = document.getElementById("container");
const root = rootElement ? createRoot(rootElement) : undefined;

/** Fluent provider following Office / OS light or dark mode. */
const Themed: React.FC<{ children: React.ReactNode }> = ({ children }) => {
  const theme = useOfficeTheme();
  React.useEffect(() => {
    // The body background fills the whole pane (CSS propagates it to the canvas). Not min-height on the provider:
    // that broke Fluent Dropdown opening on click.
    document.body.style.backgroundColor = theme.colorNeutralBackground1;
  }, [theme]);
  return (
    <FluentProvider theme={theme} data-testid="pane-theme">
      {children}
    </FluentProvider>
  );
};

function viewFromUrl(): PaneView | undefined {
  const value = new URLSearchParams(window.location.search).get("view");
  return Object.values(PaneView).find((v) => v === value);
}

/**
 * MSAL's hidden iframe (ssoSilent) or popup landed back on this page (the registered redirect URI) with the auth
 * response. MSAL v5 waits for the page to hand it over (BroadcastChannel); without this every silent sign-in timed out.
 */
function isMsalRedirect(): boolean {
  return (window.parent !== window || !!window.opener) && /[#&](code|error)=/.test(window.location.hash);
}

if (isAuthDialog()) {
  // This page is also the Office sign-in dialog (no NAA bridge, e.g. web dev sideload): sign in, post the token back.
  root?.render(
    <Themed>
      <Body1 data-testid="auth-dialog-text">Signing in to Global Advocate…</Body1>
    </Themed>
  );
  Office.onReady(() => void runAuthDialog());
} else if (isMsalRedirect()) {
  void broadcastResponseToMainFrame();
} else {
  // E2E harness reads the user's Graph token through the app's own silent path (same-origin scripts can read the
  // MSAL cache anyway, so this exposes nothing new).
  (window as unknown as { advAuth: unknown }).advAuth = { graphToken: () => getGraphToken(false) };
  /* Render application after Office initializes */
  Office.onReady((info) => {
    root?.render(
      <Themed>
        <ErrorBoundary>
          <App host={info.host} view={viewFromUrl()} />
        </ErrorBoundary>
      </Themed>
    );
  });
}
