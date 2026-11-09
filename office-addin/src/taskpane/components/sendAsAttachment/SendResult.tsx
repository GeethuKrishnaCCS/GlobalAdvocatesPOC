import * as React from "react";
import { Body1, Button, Caption1, makeStyles, tokens } from "@fluentui/react-components";
import { CheckmarkCircle24Filled, Mail20Regular, Open20Regular } from "@fluentui/react-icons";
import type { ISendResult } from "../../../shared/hooks/useSendAsAttachment";
import { OUTLOOK_APP_URL, openInBrowser } from "../../../shared/openInBrowser";

const useStyles = makeStyles({
  root: { display: "flex", flexDirection: "column", gap: tokens.spacingVerticalM },
  header: { display: "flex", alignItems: "center", gap: tokens.spacingHorizontalS },
  icon: { color: tokens.colorPaletteGreenForeground1 },
  list: { margin: 0, paddingInlineStart: tokens.spacingHorizontalL },
});

export interface ISendResultProps {
  result: ISendResult;
  subject: string;
  onStartOver: () => void;
}

/**
 * Web: draft opened in Outlook on the web (or, if the browser blocked the window, a link to open it).
 * Desktop: the draft is in Outlook Drafts; "Open Outlook app" brings Outlook up (it cannot open the draft itself),
 * the web link is secondary because the system browser may be signed in with another account.
 */
export const SendResult: React.FC<ISendResultProps> = ({ result, subject, onStartOver }) => {
  const styles = useStyles();
  const { desktop } = result;
  // The form that had focus is gone; move keyboard focus to the next useful action.
  const firstRef = React.useRef<HTMLAnchorElement>(null);
  React.useEffect(() => firstRef.current?.focus(), []);
  const message = desktop ? "Email ready in your Outlook Drafts" : result.opened ? "Opened in Outlook" : "Email ready";
  return (
    <div className={styles.root} data-testid="send-result">
      <div className={styles.header} data-testid="send-result-header">
        <CheckmarkCircle24Filled className={styles.icon} data-testid="send-result-icon" />
        <Body1 data-testid="send-result-message">
          {message}: {subject}
        </Body1>
      </div>
      <ul className={styles.list} data-testid="send-result-attachments">
        {result.attachmentNames.map((name) => (
          <li key={name} data-testid="send-result-attachment">
            <Caption1 data-testid="send-result-attachment-name">{name}</Caption1>
          </li>
        ))}
      </ul>
      {desktop && (
        // A plain link: Office hands ms-outlook: to Windows (asking once whether the add-in may open such links).
        <Button
          as="a"
          href={OUTLOOK_APP_URL}
          appearance="primary"
          icon={<Mail20Regular data-testid="send-result-outlook-app-icon" />}
          ref={firstRef}
          data-testid="send-result-outlook-app"
        >
          Open Outlook app
        </Button>
      )}
      <Button
        as="a"
        href={result.draft.webLink}
        target="_blank"
        rel="noreferrer"
        appearance={desktop || result.opened ? "secondary" : "primary"}
        icon={<Open20Regular data-testid="send-result-open-icon" />}
        onClick={(e) => openInBrowser(e, result.draft.webLink)}
        ref={desktop ? undefined : firstRef}
        data-testid="send-result-open"
      >
        {desktop ? "Open in Outlook on the web" : "Open draft in Outlook"}
      </Button>
      <Caption1 data-testid="send-result-note">
        {desktop
          ? `In Outlook, open Drafts and select "${subject}" to add recipients and your message. Outlook on the web opens in your browser with the account it is signed in with, which may not be this one.`
          : "Add recipients and your message in Outlook. If Outlook asked you to sign in first, select Open draft in Outlook again. The email is also in your Outlook Drafts, in the Outlook app too."}
      </Caption1>
      <Button onClick={onStartOver} data-testid="send-result-start-over">
        Send again
      </Button>
    </div>
  );
};
