/** Pure helpers for the Send as Attachment draft (unit-tested without a browser). */

/** Subject with the matter tag in front ("[ADVTestMatter01] Letter"), never added twice. */
export function taggedSubject(subject: string, matterRef: string | undefined): string {
  const trimmed = subject.trim();
  if (!matterRef) return trimmed;
  const tag = `[${matterRef}]`;
  return trimmed.startsWith(tag) ? trimmed : `${tag} ${trimmed}`.trim();
}

/**
 * Outlook on the web link that opens a draft in compose (its Graph webLink opens the read view, where the user would
 * first have to select Edit). Falls back to the webLink when it has no ItemID.
 */
export function composeLink(webLink: string): string {
  try {
    const url = new URL(webLink);
    const itemId = url.searchParams.get("ItemID");
    if (!itemId) return webLink;
    const id = encodeURIComponent(itemId);
    return `${url.origin}/mail/deeplink/compose/${id}?ItemID=${id}&exvsurl=1`;
  } catch {
    return webLink;
  }
}
