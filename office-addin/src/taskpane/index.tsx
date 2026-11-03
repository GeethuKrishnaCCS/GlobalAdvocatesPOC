import * as React from "react";
import { createRoot } from "react-dom/client";
import App from "./components/App";
import { Body1, FluentProvider, webLightTheme } from "@fluentui/react-components";
import { PaneView } from "../shared/config";
import { isAuthDialog, runAuthDialog } from "../shared/auth/authDialog";
import { ErrorBoundary } from "../shared/components/ErrorBoundary";

/* global document, Office, window, HTMLElement */

const rootElement: HTMLElement | null = document.getElementById("container");
const root = rootElement ? createRoot(rootElement) : undefined;

function viewFromUrl(): PaneView | undefined {
  const value = new URLSearchParams(window.location.search).get("view");
  return Object.values(PaneView).find((v) => v === value);
}

if (isAuthDialog()) {
  // This page is also the Office sign-in dialog (no NAA bridge, e.g. web dev sideload): sign in, post the token back.
  root?.render(
    <FluentProvider theme={webLightTheme}>
      <Body1 data-testid="auth-dialog-text">Signing in to Global Advocate…</Body1>
    </FluentProvider>
  );
  Office.onReady(() => void runAuthDialog());
} else {
  /* Render application after Office initializes */
  Office.onReady((info) => {
    root?.render(
      <FluentProvider theme={webLightTheme}>
        <ErrorBoundary>
          <App host={info.host} view={viewFromUrl()} />
        </ErrorBoundary>
      </FluentProvider>
    );
  });
}
