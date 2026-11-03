import { expect, test } from "@playwright/test";
import { cleanBaseName, fileNameProblem, splitFileName, toBaseName } from "../src/shared/services/fileName";
import { testIdPart } from "../src/shared/testId";
import { friendlyAuthError } from "../src/shared/auth/authErrors";

test("cleanBaseName_Should_ReplaceIllegalCharacters_When_NameHasThem", () => {
  expect(cleanBaseName(' Letter: "draft" <v2>?. ')).toBe("Letter- -draft- -v2--");
  expect(cleanBaseName("a/b\\c|d*e")).toBe("a-b-c-d-e");
});

test("fileNameProblem_Should_RejectEmptyReservedOrLong_When_Invalid", () => {
  expect(fileNameProblem("  ..  ", ".docx")).toBe("Enter a file name.");
  expect(fileNameProblem("CON", ".docx")).toMatch(/reserved/);
  expect(fileNameProblem("x".repeat(250), ".docx")).toMatch(/too long/);
  expect(fileNameProblem("Letter to client – NIF review", ".docx")).toBeUndefined();
});

test("splitFileName_Should_KeepInnerDots_When_Splitting", () => {
  expect(splitFileName("Letter.v2.docx")).toEqual({ baseName: "Letter.v2", extension: ".docx" });
  expect(splitFileName("README")).toEqual({ baseName: "README", extension: "" });
});

test("friendlyAuthError_Should_MapCodesAndKeepCorrelationId_When_Known", () => {
  expect(friendlyAuthError(new Error("user_cancelled: sign-in window closed."))).toMatch(/cancelled/);
  expect(friendlyAuthError(new Error("dialog_error 12011: blocked"))).toMatch(/blocked the sign-in window/);
  const aad = "invalid_grant: AADSTS65001: The user or administrator has not consented. Trace ID: x Correlation ID: 0c3a4f2e-1111-2222-3333-444455556666";
  expect(friendlyAuthError(new Error(aad))).toBe(
    "Global Advocate needs permission to your data. Ask your administrator to grant consent. (Correlation ID: 0c3a4f2e-1111-2222-3333-444455556666)"
  );
  expect(friendlyAuthError(new Error("something odd"))).toBe("something odd");
});
test("fileNameProblem_Should_RejectSharePointReservedParts_When_Present", () => {
  expect(fileNameProblem("~$Letter", ".docx")).toMatch(/reserved/);
  expect(fileNameProblem("a_vti_b", ".docx")).toMatch(/reserved/);
});

test("toBaseName_Should_DropTypedExtension_When_UserTypesIt", () => {
  expect(toBaseName("Letter.docx", ".docx")).toBe("Letter");
  expect(toBaseName("Letter.DOCX ", ".docx")).toBe("Letter");
  expect(toBaseName("Letter.v2", ".docx")).toBe("Letter.v2");
  expect(fileNameProblem(".docx", ".docx")).toBe("Enter a file name.");
});

test("testIdPart_Should_ProduceKebabCase_When_GivenNames", () => {
  expect(testIdPart("ADVTest_FolderA")).toBe("advtest-foldera");
  expect(testIdPart(" Correspondence › Inward ")).toBe("correspondence-inward");
});
