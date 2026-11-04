import * as React from "react";
import { Body1, Button, Caption1, makeStyles, tokens } from "@fluentui/react-components";
import { CheckmarkCircle24Filled, Open20Regular } from "@fluentui/react-icons";
import type { ISendResult } from "../../../shared/hooks/useSendAsAttachment";
import { openInBrowser } from "../../../shared/openInBrowser";

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

/** Draft opened in Outlook (or, if the browser blocked the window, a link to open it) and start another. */
export const SendResult: React.FC<ISendResultProps> = ({ result, subject, onStartOver }) => {
  const styles = useStyles();
  // The form that had focus is gone; move keyboard focus to the next useful action.
  const openRef = React.useRef<HTMLAnchorElement>(null);
  React.useEffect(() => openRef.current?.focus(), []);
  return (
    <div className={styles.root} data-testid="send-result">
      <div className={styles.header} data-testid="send-result-header">
        <CheckmarkCircle24Filled className={styles.icon} data-testid="send-result-icon" />
        <Body1 data-testid="send-result-message">
          {result.opened ? "Opened in Outlook" : "Email ready"}: {subject}
        </Body1>
      </div>
      <ul className={styles.list} data-testid="send-result-attachments">
        {result.attachmentNames.map((name) => (
          <li key={name} data-testid="send-result-attachment">
            <Caption1 data-testid="send-result-attachment-name">{name}</Caption1>
          </li>
        ))}
      </ul>
      <Button
        as="a"
        href={result.draft.webLink}
        target="_blank"
        rel="noreferrer"
        appearance={result.opened ? "secondary" : "primary"}
        icon={<Open20Regular data-testid="send-result-open-icon" />}
        onClick={(e) => openInBrowser(e, result.draft.webLink)}
        ref={openRef}
        data-testid="send-result-open"
      >
        Open draft in Outlook
      </Button>
      <Caption1 data-testid="send-result-note">
        Add recipients and your message in Outlook. If Outlook asked you to sign in first, select Open draft in Outlook
        again. The email is also in your Outlook Drafts, in the Outlook app too.
      </Caption1>
      <Button onClick={onStartOver} data-testid="send-result-start-over">
        Send again
      </Button>
    </div>
  );
};
