import * as React from "react";
import {
  Body1,
  Button,
  Caption1,
  MessageBar,
  MessageBarBody,
  Spinner,
  makeStyles,
  tokens,
} from "@fluentui/react-components";
import { Document20Regular } from "@fluentui/react-icons";
import { useSendAsAttachment } from "../../../shared/hooks/useSendAsAttachment";
import { AttachSection } from "./AttachSection";
import { MatterSection } from "./MatterSection";
import { MessageSection } from "./MessageSection";
import { SendResult } from "./SendResult";

const useStyles = makeStyles({
  root: { display: "flex", flexDirection: "column", gap: tokens.spacingVerticalL },
  file: { display: "flex", alignItems: "center", gap: tokens.spacingHorizontalS },
  fileText: { display: "flex", flexDirection: "column", minWidth: 0, overflowWrap: "anywhere" },
  footer: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalS,
    paddingTop: tokens.spacingVerticalS,
    borderTop: `${tokens.strokeWidthThin} solid ${tokens.colorNeutralStroke2}`,
  },
});

/** Send as Attachment: an Outlook draft with the open document attached, tagged with its matter. */
export const SendAsAttachment: React.FC<{ host: Office.HostType }> = ({ host }) => {
  const styles = useStyles();
  const { state, actions, computed } = useSendAsAttachment(host);
  const busy = state.isCreating;

  return (
    <div className={styles.root} data-testid="send-as-attachment">
      <div className={styles.file} data-testid="send-file">
        <Document20Regular data-testid="send-file-icon" />
        <span className={styles.fileText}>
          <Body1 data-testid="send-file-name">{computed.fileLabel}</Body1>
          <Caption1 data-testid="send-file-location">{computed.fileLocation}</Caption1>
        </span>
      </div>
      {state.error && (
        <MessageBar intent="error" layout="multiline" data-testid="send-error">
          <MessageBarBody>{state.error}</MessageBarBody>
        </MessageBar>
      )}
      {state.result ? (
        <SendResult
          result={state.result}
          subject={computed.subjectToSend}
          onStartOver={actions.startOver}
        />
      ) : (
        <>
          <AttachSection
            form={state.form}
            attachments={computed.attachments}
            disabled={busy || state.retryPending}
            onChange={actions.update}
          />
          <MatterSection
            form={state.form}
            matters={state.matters}
            matter={state.matter}
            filingTarget={state.filingTarget}
            fileEmail={computed.fileEmail}
            loadFailed={!!state.error && !state.matters}
            disabled={busy}
            locked={state.retryPending}
            onChange={actions.update}
            onSelectMatter={(m) => void actions.selectMatter(m)}
            onRetryLoad={() => void actions.loadMatters()}
          />
          {state.isLoading && <Spinner size="tiny" label="Loading…" data-testid="send-loading" />}
          <MessageSection
            form={state.form}
            subjectToSend={computed.subjectToSend}
            toError={computed.toError}
            ccError={computed.ccError}
            disabled={busy}
            onChange={actions.update}
          />
          <div className={styles.footer} data-testid="send-footer">
            <Caption1 data-testid="send-summary">{computed.summary}</Caption1>
            <Button
              appearance="primary"
              disabled={!computed.canCreate || state.isLoading}
              icon={busy ? <Spinner size="tiny" data-testid="send-create-spinner" /> : undefined}
              onClick={() => void actions.create()}
              data-testid="send-create"
            >
              {busy ? "Creating draft…" : "Create draft"}
            </Button>
          </div>
        </>
      )}
    </div>
  );
};
