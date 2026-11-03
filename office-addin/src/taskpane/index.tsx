import * as React from "react";
import { createRoot } from "react-dom/client";
import App from "./components/App";
import { Body1, FluentProvider, makeStyles } from "@fluentui/react-components";
import { PaneView } from "../shared/config";
import { isAuthDialog, runAuthDialog } from "../shared/auth/authDialog";
import { ErrorBoundary } from "../shared/components/ErrorBoundary";
import { useOfficeTheme } from "../shared/hooks/useOfficeTheme";

/* global document, Office, window, HTMLElement */

const rootElement: HTMLElement | null = document.getElementById("container");
const root = rootElement ? createRoot(rootElement) : undefined;

const useStyles = makeStyles({
  // The provider paints the theme background; it must cover the whole pane, not just the content.
  root: { minHeight: "100vh" },
});

/** Fluent provider following Office / OS light or dark mode. */
const Themed: React.FC<{ children: React.ReactNode }> = ({ children }) => {
  const theme = useOfficeTheme();
  const styles = useStyles();
  React.useEffect(() => {
    // Overscroll / rubber-band areas outside the provider match too.
    document.body.style.backgroundColor = theme.colorNeutralBackground1;
  }, [theme]);
  return (
    <FluentProvider theme={theme} className={styles.root} data-testid="pane-theme">
      {children}
    </FluentProvider>
  );
};

function viewFromUrl(): PaneView | undefined {
  const value = new URLSearchParams(window.location.search).get("view");
  return Object.values(PaneView).find((v) => v === value);
}

if (isAuthDialog()) {
  // This page is also the Office sign-in dialog (no NAA bridge, e.g. web dev sideload): sign in, post the token back.
  root?.render(
    <Themed>
      <Body1 data-testid="auth-dialog-text">Signing in to Global Advocate…</Body1>
    </Themed>
  );
  Office.onReady(() => void runAuthDialog());
} else {
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
