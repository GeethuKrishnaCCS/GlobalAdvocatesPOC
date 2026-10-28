import * as React from "react";
import { createRoot } from "react-dom/client";
import App from "./components/App";
import { FluentProvider, webLightTheme } from "@fluentui/react-components";
import { PaneView } from "../shared/config";

/* global document, Office, window, HTMLElement */

const rootElement: HTMLElement | null = document.getElementById("container");
const root = rootElement ? createRoot(rootElement) : undefined;

function viewFromUrl(): PaneView | undefined {
  const value = new URLSearchParams(window.location.search).get("view");
  return Object.values(PaneView).find((v) => v === value);
}

/* Render application after Office initializes */
Office.onReady((info) => {
  root?.render(
    <FluentProvider theme={webLightTheme}>
      <App host={info.host} view={viewFromUrl()} />
    </FluentProvider>
  );
});
