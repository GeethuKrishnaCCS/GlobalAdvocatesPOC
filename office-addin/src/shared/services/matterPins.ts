import type { IMatter } from "./matterService";

/* global Office */

/** Roaming setting (per mailbox, follows the user to every Outlook client) holding the pinned matter references. */
const PINS_SETTING = "pinnedMatters";

/** Pinned matters first, each group keeping the given order. Pure. */
export function pinnedFirst(matters: IMatter[], pinned: ReadonlySet<string>): IMatter[] {
  return [
    ...matters.filter((m) => pinned.has(m.ref)),
    ...matters.filter((m) => !pinned.has(m.ref)),
  ];
}

/** Matter references from a stored setting; anything else (missing, corrupted) reads as no pins. Pure. */
export function parsePins(stored: unknown): string[] {
  return Array.isArray(stored)
    ? stored.filter((v): v is string => typeof v === "string" && !!v)
    : [];
}

/** The user's pinned matter references. */
export function loadPins(): string[] {
  return parsePins(Office.context.roamingSettings?.get(PINS_SETTING));
}

/** Stores the user's pinned matter references in their mailbox. */
export function savePins(refs: string[]): Promise<void> {
  const settings = Office.context.roamingSettings;
  if (!settings) return Promise.reject(new Error("This Outlook cannot store pinned matters."));
  settings.set(PINS_SETTING, refs);
  return new Promise((resolve, reject) =>
    settings.saveAsync((result) =>
      result.status === Office.AsyncResultStatus.Succeeded
        ? resolve()
        : reject(new Error(result.error?.message ?? "Outlook could not save your pinned matters."))
    )
  );
}
