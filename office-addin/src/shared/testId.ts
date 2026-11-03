/** Kebab-case fragment for a data-testid built from user data ("ADVTest_FolderA" → "advtest-foldera"). Pure. */
export function testIdPart(value: string): string {
  return value
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "");
}
