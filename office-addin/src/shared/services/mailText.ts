/** Pure helpers for the Send as Attachment draft (unit-tested without a browser). */

/** Very small address check: something@domain.tld, no spaces. Exchange does the real validation. */
const EMAIL = /^[^\s@;,]+@[^\s@;,]+\.[^\s@;,]+$/;

/** Splits "a@x.com; b@y.com, c@z.com" into addresses (empty parts dropped). */
export function parseRecipients(value: string): string[] {
  return value
    .split(/[;,]/)
    .map((part) => part.trim())
    .filter(Boolean);
}

/** The first address that does not look like an email address, or undefined when all do. */
export function invalidRecipient(value: string): string | undefined {
  return parseRecipients(value).find((address) => !EMAIL.test(address));
}

/** Subject with the matter tag in front ("[ADVTestMatter01] Letter"), never added twice. */
export function taggedSubject(subject: string, matterRef: string | undefined): string {
  const trimmed = subject.trim();
  if (!matterRef) return trimmed;
  const tag = `[${matterRef}]`;
  return trimmed.startsWith(tag) ? trimmed : `${tag} ${trimmed}`.trim();
}
