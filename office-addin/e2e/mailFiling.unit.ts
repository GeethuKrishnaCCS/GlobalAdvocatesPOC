import { expect, test } from "@playwright/test";
import {
  attachProblem,
  base64ToBlob,
  bytesToBase64,
  classifyAttachment,
  emlFileName,
  folderPathSegments,
  isAlreadyAttached,
  parseFilingTarget,
} from "../src/shared/services/mailFiling";

test("emlFileName_Should_PrefixDateAndCleanSubject_When_Filing", () => {
  expect(emlFileName("Re: Order 12/2026?", new Date(2026, 10, 5, 23, 30))).toBe("2026-11-05 Re- Order 12-2026-.eml");
  expect(emlFileName("   ", new Date(2026, 0, 2))).toBe("2026-01-02 (no subject).eml");
  expect(emlFileName("x".repeat(300), new Date(2026, 0, 2)).length).toBeLessThanOrEqual(250);
});

test("classifyAttachment_Should_TickFilesSkipInlineBlockCloud_When_Listing", () => {
  expect(classifyAttachment({ id: "1", name: "Order.pdf", size: 10, attachmentType: "file", isInline: false })).toEqual({
    id: "1", fileName: "Order.pdf", size: 10, kind: "file", fileable: true, ticked: true,
  });
  expect(classifyAttachment({ id: "2", name: "image001.png", size: 5, attachmentType: "file", isInline: true })).toMatchObject({
    fileable: true, ticked: false,
  });
  expect(classifyAttachment({ id: "3", name: "Big.docx", size: 0, attachmentType: "cloud", isInline: false })).toMatchObject({
    fileable: false, ticked: false, reason: "Shared as a link, not fileable",
  });
  expect(classifyAttachment({ id: "4", name: "Fwd: Hearing", size: 9, attachmentType: "item", isInline: false })).toMatchObject({
    fileName: "Fwd- Hearing.eml", kind: "item", fileable: true, ticked: true,
  });
});

test("parseFilingTarget_Should_PreferFolderPropertyThenCategory_When_Present", () => {
  expect(parseFilingTarget("b!drive|01ITEM", ["Matter: ADVTestMatter01"])).toEqual({ driveId: "b!drive", folderId: "01ITEM" });
  expect(parseFilingTarget(undefined, ["Blue", "Matter: ADVTestMatter01"])).toEqual({ matterRef: "ADVTestMatter01" });
  expect(parseFilingTarget("garbage", [])).toBeUndefined();
  expect(parseFilingTarget(undefined, undefined)).toBeUndefined();
});

test("isAlreadyAttached_Should_MatchNameIgnoringCase_When_Compared", () => {
  expect(isAlreadyAttached("Order.DOCX", ["order.docx", "x.pdf"])).toBe(true);
  expect(isAlreadyAttached("Order.pdf", ["order.docx"])).toBe(false);
});

test("attachProblem_Should_RejectOverOutlookLimit_When_TooLarge", () => {
  expect(attachProblem("Big.pdf", 26 * 1024 * 1024)).toMatch(/Big\.pdf.*too large/);
  expect(attachProblem("Small.pdf", 1024)).toBeUndefined();
});

test("base64ToBlob_Should_DecodeBytes_When_Given", async () => {
  const blob = base64ToBlob("SGVsbG8=", "message/rfc822");
  expect(blob.type).toBe("message/rfc822");
  expect(await blob.text()).toBe("Hello");
});

test("bytesToBase64_Should_RoundTrip_When_Encoding", async () => {
  const bytes = new Uint8Array(70_000).map((_, i) => i % 256);
  const base64 = bytesToBase64(bytes.buffer);
  expect(base64.startsWith("AAECAwQF")).toBe(true);
  expect(new Uint8Array(await base64ToBlob(base64, "").arrayBuffer())).toEqual(bytes);
});

test("folderPathSegments_Should_SplitDecodedPath_When_UnderRoot", () => {
  expect(folderPathSegments("/drives/b!x/root:/Correspondence/My%20Matter")).toEqual(["Correspondence", "My Matter"]);
  expect(folderPathSegments("/drives/b!x/root:")).toEqual([]);
  expect(folderPathSegments(undefined)).toEqual([]);
});
