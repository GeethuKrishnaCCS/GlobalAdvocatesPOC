import * as React from "react";
import { Button, Caption1, Checkbox, Dropdown, Field, Option, Subtitle2, makeStyles, tokens } from "@fluentui/react-components";
import { FILING_FOLDER_PATH, MATTER_CATEGORY_PREFIX } from "../../../shared/config";
import type { ISendForm } from "../../../shared/hooks/useSendAsAttachment";
import type { IMatter } from "../../../shared/services/matterService";
import { testIdPart } from "../../../shared/testId";

const useStyles = makeStyles({
  root: { display: "flex", flexDirection: "column", gap: tokens.spacingVerticalS },
  dropdown: { minWidth: "auto" },
});

export interface IMatterSectionProps {
  form: ISendForm;
  matters: IMatter[] | undefined;
  matter: IMatter | undefined;
  /** "driveId|itemId", null when the matter has no filing folder, undefined while unknown. */
  filingTarget: string | null | undefined;
  /** Whether the draft will carry the filing marker (ticked, tagged, folder exists). */
  fileEmail: boolean;
  /** Loading the matter list failed: offer Retry. */
  loadFailed: boolean;
  disabled: boolean;
  /** A draft waits for its attachments: tag and filing are locked until it is finished. */
  locked: boolean;
  onChange: (patch: Partial<ISendForm>) => void;
  onSelectMatter: (matter: IMatter) => void;
  onRetryLoad: () => void;
}

/** Matter tag ([ref] subject prefix + Outlook category) and the "file the sent email" marker. */
export const MatterSection: React.FC<IMatterSectionProps> = (props) => {
  const { form, matters, matter, filingTarget, fileEmail, loadFailed, disabled, locked } = props;
  const styles = useStyles();
  const ref = matter?.ref ?? "<matter>";
  const folderLabel = FILING_FOLDER_PATH.split("/").join(" › ");
  const label = (m: IMatter): string => `${m.ref} – ${m.title}`;

  return (
    <div className={styles.root} data-testid="send-matter">
      <Subtitle2 data-testid="send-matter-heading">Matter</Subtitle2>
      <Checkbox
        checked={form.tagMatter}
        disabled={disabled || locked}
        onChange={(_, data) => props.onChange({ tagMatter: !!data.checked })}
        label={`Tag the email with the matter workspace. Adds [${ref}] to the subject and the Outlook category "${MATTER_CATEGORY_PREFIX}${ref}", so replies can be filed to this matter.`}
        data-testid="send-matter-tag"
      />
      <Field
        label="Matter workspace"
        required={form.tagMatter}
        validationMessage={form.tagMatter && matters && !matter ? "Select the matter to tag." : undefined}
        data-testid="send-matter-picker-field"
      >
        <Dropdown
          placeholder={matters ? "Select a matter" : loadFailed ? "Matters could not be loaded" : "Loading matters…"}
          value={matter ? label(matter) : ""}
          selectedOptions={matter ? [matter.siteId] : []}
          className={styles.dropdown}
          disabled={disabled || locked || !form.tagMatter || !matters}
          onOptionSelect={(_, data) => {
            const picked = matters?.find((m) => m.siteId === data.optionValue);
            if (picked) props.onSelectMatter(picked);
          }}
          data-testid="send-matter-picker"
        >
          {(matters ?? []).map((m) => (
            <Option key={m.siteId} value={m.siteId} text={label(m)} data-testid={`send-matter-option-${testIdPart(m.ref)}`}>
              {label(m)}
            </Option>
          ))}
        </Dropdown>
      </Field>
      {loadFailed && !matters && (
        <Button size="small" disabled={disabled} onClick={props.onRetryLoad} data-testid="send-matter-retry">
          Retry loading matters
        </Button>
      )}
      <Checkbox
        checked={fileEmail}
        disabled={disabled || locked || !form.tagMatter || !filingTarget}
        onChange={(_, data) => props.onChange({ fileSentEmail: !!data.checked })}
        label={`File the sent email to ${folderLabel}`}
        data-testid="send-matter-file-email"
      />
      {filingTarget === null && (
        <Caption1 data-testid="send-matter-no-filing-folder">This matter has no {folderLabel} folder.</Caption1>
      )}
    </div>
  );
};
