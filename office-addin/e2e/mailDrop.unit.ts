import { expect, test } from "@playwright/test";
import { droppedFileName, folderDropKey, isEmailFile, matterDropKey } from "../src/shared/services/mailDrop";
import { parsePins, pinnedFirst } from "../src/shared/services/matterPins";
import type { IMatter } from "../src/shared/services/matterService";

const matter = (ref: string): IMatter => ({ siteId: `site-${ref}`, ref, title: ref, webUrl: "" });

test("droppedFileName_Should_CleanForSharePointAndKeepExtension_When_Dropped", () => {
  expect(droppedFileName("Re: Order 12/2026?.eml")).toBe("Re- Order 12-2026-.eml");
  expect(droppedFileName("  Hearing notes.msg ")).toBe("Hearing notes.msg");
  expect(droppedFileName(".eml")).toBe("(no subject).eml");
  expect(droppedFileName("???.eml")).toBe("---.eml");
  expect(droppedFileName("Notes")).toBe("Notes");
});

test("droppedFileName_Should_StayWithinSharePointNameLength_When_SubjectIsLong", () => {
  const name = droppedFileName(`${"x".repeat(400)}.eml`);
  expect(name.length).toBe(250);
  expect(name.endsWith(".eml")).toBe(true);
});

test("isEmailFile_Should_AcceptEmlAndMsgOnly_When_Listing", () => {
  expect(isEmailFile("2026-11-05 Order.eml")).toBe(true);
  expect(isEmailFile("Order.MSG")).toBe(true);
  expect(isEmailFile("Order.docx")).toBe(false);
  expect(isEmailFile("eml")).toBe(false);
});

test("dropKeys_Should_NotCollide_When_MatterAndFolderShareAnId", () => {
  expect(matterDropKey("x")).not.toBe(folderDropKey("x"));
});

test("pinnedFirst_Should_ListPinnedFirstKeepingOrder_When_SomePinned", () => {
  const matters = ["A", "B", "C", "D"].map(matter);
  expect(pinnedFirst(matters, new Set(["C", "B"])).map((m) => m.ref)).toEqual(["B", "C", "A", "D"]);
  expect(pinnedFirst(matters, new Set()).map((m) => m.ref)).toEqual(["A", "B", "C", "D"]);
  // A pinned matter the user can no longer access is simply not listed.
  expect(pinnedFirst(matters, new Set(["Gone"])).map((m) => m.ref)).toEqual(["A", "B", "C", "D"]);
});

test("parsePins_Should_ReadNoPins_When_SettingMissingOrCorrupted", () => {
  expect(parsePins(["A", "B"])).toEqual(["A", "B"]);
  expect(parsePins(undefined)).toEqual([]);
  expect(parsePins("A")).toEqual([]);
  expect(parsePins(["A", 1, null, ""])).toEqual(["A"]);
});
