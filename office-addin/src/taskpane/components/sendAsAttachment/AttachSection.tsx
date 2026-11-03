import * as React from "react";
import { Caption1, Field, Input, Subtitle2, ToggleButton, makeStyles, tokens } from "@fluentui/react-components";
import type { AttachFormat, IAttachment, ISendForm } from "../../../shared/hooks/useSendAsAttachment";

const FORMATS: { value: AttachFormat; label: string }[] = [
  { value: "native", label: "Word" },
  { value: "pdf", label: "PDF" },
  { value: "both", label: "Both" },
];

const useStyles = makeStyles({
  root: { display: "flex", flexDirection: "column", gap: tokens.spacingVerticalS },
  formats: { display: "grid", gridTemplateColumns: "repeat(3, 1fr)" },
  format: { minWidth: 0 },
});

export interface IAttachSectionProps {
  form: ISendForm;
  attachments: IAttachment[];
  disabled: boolean;
  onChange: (patch: Partial<ISendForm>) => void;
}

/** "Attach as" Word / PDF / Both, and the (renamable, extension-locked) name of each attachment. */
export const AttachSection: React.FC<IAttachSectionProps> = ({ form, attachments, disabled, onChange }) => {
  const styles = useStyles();
  return (
    <div className={styles.root} data-testid="send-attach">
      <Subtitle2 data-testid="send-attach-heading">Send as attachment</Subtitle2>
      <Caption1 data-testid="send-attach-intro">The current version of this document is attached to a new email.</Caption1>
      <Field label="Attach as" data-testid="send-attach-format-field">
        <div className={styles.formats} role="group" aria-label="Attach as" data-testid="send-attach-format">
          {FORMATS.map((f) => (
            <ToggleButton
              key={f.value}
              checked={form.format === f.value}
              className={styles.format}
              disabled={disabled}
              onClick={() => onChange({ format: f.value })}
              data-testid={`send-attach-format-${f.value}`}
            >
              {f.label}
            </ToggleButton>
          ))}
        </div>
      </Field>
      {attachments.map((a) => (
        <Field
          key={a.format}
          label={a.format === "pdf" ? "PDF attachment name" : "Attachment name"}
          validationMessage={a.problem}
          data-testid={`send-attach-name-${a.format}-field`}
        >
          <Input
            value={a.baseName}
            contentAfter={<Caption1 data-testid={`send-attach-extension-${a.format}`}>{a.extension}</Caption1>}
            disabled={disabled}
            onChange={(_, data) => onChange(a.format === "pdf" ? { pdfName: data.value } : { nativeName: data.value })}
            data-testid={`send-attach-name-${a.format}`}
          />
        </Field>
      ))}
    </div>
  );
};
