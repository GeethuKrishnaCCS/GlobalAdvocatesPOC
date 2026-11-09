/* global Office */
// Compose banner (event-based activation, OnMessageCompose): offers Attach from Matter whenever an email is composed:
// new, reply, reply all, forward, and reopening a draft (OnNewMessageCompose skipped drafts).
// Plain script, not bundled: classic Outlook on Windows loads this file directly in its JavaScript-only runtime.

var BANNER_KEY = "advAttachFromMatter";

function onMessageComposeHandler(event) {
  Office.context.mailbox.item.notificationMessages.addAsync(
    BANNER_KEY,
    {
      type: Office.MailboxEnums.ItemNotificationMessageType.InsightMessage,
      message: "Global Advocate: attach documents from a matter.",
      // Required; the unified manifest ignores it and uses the first ribbon control icon.
      icon: "Icon.16x16",
      actions: [
        {
          actionText: "Attach from Matter",
          actionType: Office.MailboxEnums.ActionType.ShowTaskPane,
          // The compose button in both manifests.
          commandId: "AttachFromMatterButton",
          contextData: "{}",
        },
      ],
    },
    function () {
      // A banner that could not be shown must never hold up the new email.
      event.completed();
    }
  );
}

Office.actions.associate("onMessageComposeHandler", onMessageComposeHandler);
