import { expect, test } from "@playwright/test";
import type { IMatterFolderStatus } from "../src/shared/services/matterFolderService";
import { clockTime, folderStatusLine, statusForPins } from "../src/shared/services/matterFolderStatus";

const status = (matterRef: string, filedCount = 0, lastError: string | null = null): IMatterFolderStatus => ({
  matterRef,
  filedCount,
  lastSweepUtc: null,
  lastError,
});
const at = new Date(2026, 10, 13, 10, 5);

test("clockTime_Should_PadHoursAndMinutes_When_Formatting", () => {
  expect(clockTime(at)).toBe("10:05");
  expect(clockTime(new Date(2026, 0, 1, 7, 9))).toBe("07:09");
});

test("statusForPins_Should_DropUnpinnedMatters_When_BackendStillListsThem", () => {
  const rows = [status("A"), status("B")];
  expect(statusForPins(rows, new Set(["B"]))).toEqual([status("B")]);
});

test("folderStatusLine_Should_BeUndefined_When_NothingIsPinned", () => {
  expect(folderStatusLine([status("A")], new Set(), at, "boom")).toBeUndefined();
});

test("folderStatusLine_Should_SummarisePinnedMatters_When_NoErrors", () => {
  const rows = [status("A", 5), status("B", 7), status("C", 99)];
  expect(folderStatusLine(rows, new Set(["A", "B"]), at, undefined)).toEqual({
    text: "Mailbox folders: 2 matters, 12 emails filed, last checked 10:05",
    isError: false,
  });
  expect(folderStatusLine([status("A", 1)], new Set(["A"]), undefined, undefined)?.text).toBe(
    "Mailbox folders: 1 matter, 1 email filed"
  );
});

test("folderStatusLine_Should_ShowChecking_When_StatusNotReadYet", () => {
  expect(folderStatusLine(undefined, new Set(["A"]), undefined, undefined)).toEqual({
    text: "Mailbox folders: checking…",
    isError: false,
  });
});

test("folderStatusLine_Should_ShowRegistrationErrorFirst_When_BothErrorsExist", () => {
  const rows = [status("A", 0, "sweep failed")];
  expect(folderStatusLine(rows, new Set(["A"]), at, "A: denied")).toEqual({
    text: "Mailbox folders: A: denied",
    isError: true,
  });
});

test("folderStatusLine_Should_ShowFirstLastError_When_PinnedMatterFailed", () => {
  const rows = [status("X", 0, "ignored: not pinned"), status("A", 0, "first"), status("B", 0, "second")];
  expect(folderStatusLine(rows, new Set(["A", "B"]), at, undefined)).toEqual({
    text: "Mailbox folders: A: first",
    isError: true,
  });
});
