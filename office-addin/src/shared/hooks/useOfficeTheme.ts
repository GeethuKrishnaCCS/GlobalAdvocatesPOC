import * as React from "react";
import { webDarkTheme, webLightTheme, type Theme } from "@fluentui/react-components";

/* global Office, window */

const DARK_QUERY = "(prefers-color-scheme: dark)";
/** Relative luminance below which a background counts as dark. */
const DARK_LUMINANCE = 0.4;

/** "#1f1f1f" → true. Pure; undefined for anything that is not a 6-digit hex colour. */
export function isDarkColor(hex: string | undefined): boolean | undefined {
  const match = /^#?([0-9a-f]{6})$/i.exec(hex ?? "");
  if (!match) return undefined;
  const [r, g, b] = [0, 2, 4].map((i) => parseInt(match[1].slice(i, i + 2), 16) / 255);
  return 0.2126 * r + 0.7152 * g + 0.0722 * b < DARK_LUMINANCE;
}

/** Dark when the Office theme says so (Office's own dark mode); otherwise follow the browser / OS setting. */
function detectDark(): boolean {
  const officeDark = isDarkColor(Office.context?.officeTheme?.bodyBackgroundColor);
  return officeDark ?? window.matchMedia?.(DARK_QUERY).matches ?? false;
}

/** Fluent theme matching Office / the OS, updated live when either switches. */
export function useOfficeTheme(): Theme {
  const [dark, setDark] = React.useState(detectDark);

  React.useEffect(() => {
    const update = (): void => setDark(detectDark());
    const media = window.matchMedia?.(DARK_QUERY);
    const onMedia = (): void => update();
    media?.addEventListener("change", onMedia);
    // OfficeThemeChanged is not raised by every host; where it is not, the media query and the next load cover it.
    try {
      Office.context.document?.addHandlerAsync?.(Office.EventType.OfficeThemeChanged, update);
    } catch {
      // Host without the event.
    }
    return () => {
      media?.removeEventListener("change", onMedia);
      try {
        Office.context.document?.removeHandlerAsync?.(Office.EventType.OfficeThemeChanged, {
          handler: update,
        });
      } catch {
        // Host without the event.
      }
    };
  }, []);

  return dark ? webDarkTheme : webLightTheme;
}
