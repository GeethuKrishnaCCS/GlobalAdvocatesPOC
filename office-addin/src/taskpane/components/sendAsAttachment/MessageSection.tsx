import * as React from "react";
import { Field, Input, Subtitle2, Textarea, makeStyles, tokens } from "@fluentui/react-components";
import type { ISendForm } from "../../../shared/hooks/useSendAsAttachment";

const useStyles = makeStyles({
  root: { display: "flex", flexDirection: "column", gap: tokens.spacingVerticalS },
});

export interface IMessageSectionProps {
  form: ISendForm;
  subjectToSend: string;
  toError: string | undefined;
  ccError: string | undefined;
  disabled: boolean;
  onChange: (patch: Partial<ISendForm>) => void;
}

/** To / Cc / Subject / Message, all editable; the subject hint shows the tagged subject that will be used. */
export const MessageSection: React.FC<IMessageSectionProps> = ({ form, subjectToSend, toError, ccError, disabled, onChange }) => {
  const styles = useStyles();
  return (
    <div className={styles.root} data-testid="send-message">
      <Subtitle2 data-testid="send-message-heading">Email</Subtitle2>
      <Field label="To" validationMessage={toError} hint="Separate addresses with ; or ," data-testid="send-message-to-field">
        <Input inputMode="email" value={form.to} disabled={disabled} onChange={(_, d) => onChange({ to: d.value })} data-testid="send-message-to" />
      </Field>
      <Field label="Cc" validationMessage={ccError} data-testid="send-message-cc-field">
        <Input inputMode="email" value={form.cc} disabled={disabled} onChange={(_, d) => onChange({ cc: d.value })} data-testid="send-message-cc" />
      </Field>
      <Field label="Subject" hint={`Sent as: ${subjectToSend}`} data-testid="send-message-subject-field">
        <Input value={form.subject} disabled={disabled} onChange={(_, d) => onChange({ subject: d.value })} data-testid="send-message-subject" />
      </Field>
      <Field label="Message" data-testid="send-message-body-field">
        <Textarea
          value={form.message}
          resize="vertical"
          disabled={disabled}
          onChange={(_, d) => onChange({ message: d.value })}
          data-testid="send-message-body"
        />
      </Field>
    </div>
  );
};
