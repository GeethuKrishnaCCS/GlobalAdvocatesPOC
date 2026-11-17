import * as React from "react";
import {
  Body1,
  Caption1,
  Card,
  Checkbox,
  Dropdown,
  Field,
  Input,
  Link,
  Option,
  Subtitle2,
  Textarea,
  makeStyles,
  tokens,
} from "@fluentui/react-components";
import { FolderOpen20Regular } from "@fluentui/react-icons";
import { DIRECTIONS, DOCUMENT_TYPES } from "../../../shared/config";
import type { ISaveDetails } from "../../../shared/hooks/useSaveToMatter";
import { testIdPart } from "../../../shared/testId";

const useStyles = makeStyles({
  root: { display: "flex", flexDirection: "column", gap: tokens.spacingVerticalM },
  heading: { display: "flex", flexDirection: "column", gap: tokens.spacingVerticalXXS },
  location: { display: "flex", alignItems: "center", gap: tokens.spacingHorizontalS },
  locationText: { display: "flex", flexDirection: "column", flexGrow: 1, minWidth: 0, overflowWrap: "anywhere" },
  // Narrow panes stack the pair; Fluent Dropdown has a 250px min-width that would overflow the grid cell.
  pair: { display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(130px, 1fr))", gap: tokens.spacingHorizontalS },
  dropdown: { minWidth: "auto" },
  link: { flexShrink: 0 },
});

interface IChoiceFieldProps {
  label: string;
  testId: string;
  options: string[];
  value: string;
  disabled: boolean;
  onChange: (value: string) => void;
}

/** Optional single choice; clearable, so a value picked by mistake can be removed before saving. */
const ChoiceField: React.FC<IChoiceFieldProps> = ({ label, testId, options, value, disabled, onChange }) => {
  const styles = useStyles();
  return (
    <Field label={label} data-testid={`${testId}-field`}>
      <Dropdown
        placeholder="Select"
        clearable
        value={value}
        selectedOptions={value ? [value] : []}
        className={styles.dropdown}
        disabled={disabled}
        onOptionSelect={(_, data) => onChange(data.optionValue ?? "")}
        data-testid={testId}
      >
        {options.map((o) => (
          <Option key={o} value={o} data-testid={`${testId}-option-${testIdPart(o)}`}>
            {o}
          </Option>
        ))}
      </Dropdown>
    </Field>
  );
};

export interface IDetailsStepProps {
  locationLabel: string;
  matterTitle: string;
  extension: string;
  details: ISaveDetails;
  fileNameError: string | undefined;
  disabled: boolean;
  onChange: (patch: Partial<ISaveDetails>) => void;
  onChangeLocation: () => void;
}

/** Step 3: file name (extension locked) and the properties the document gets in the matter. */
export const DetailsStep: React.FC<IDetailsStepProps> = (props) => {
  const { locationLabel, matterTitle, extension, details, fileNameError, disabled, onChange, onChangeLocation } = props;
  const styles = useStyles();

  return (
    <div className={styles.root} data-testid="details-step">
      <div className={styles.heading}>
        <Subtitle2 data-testid="details-step-heading">Document details</Subtitle2>
        <Caption1 data-testid="details-step-intro">These become the document&apos;s properties in the matter workspace.</Caption1>
      </div>
      <Card size="small" className={styles.location} orientation="horizontal" data-testid="details-step-location">
        <FolderOpen20Regular data-testid="details-step-location-icon" />
        <span className={styles.locationText}>
          <Body1 data-testid="details-step-location-path">{locationLabel}</Body1>
          <Caption1 data-testid="details-step-location-matter">{matterTitle}</Caption1>
        </span>
        <Link as="button" className={styles.link} onClick={onChangeLocation} disabled={disabled} data-testid="details-step-change-location">
          Change
        </Link>
      </Card>
      <Field label="File name" required validationMessage={fileNameError} data-testid="details-step-file-name-field">
        <Input
          value={details.baseName}
          contentAfter={<Caption1 data-testid="details-step-extension">{extension}</Caption1>}
          disabled={disabled}
          onChange={(_, data) => onChange({ baseName: data.value })}
          data-testid="details-step-file-name"
        />
      </Field>
      <Subtitle2 data-testid="details-step-properties-heading">Properties</Subtitle2>
      <div className={styles.pair}>
        <ChoiceField label="Document type" testId="details-step-doc-type" options={DOCUMENT_TYPES} value={details.docType} disabled={disabled} onChange={(docType) => onChange({ docType })} />
        <ChoiceField label="Direction" testId="details-step-direction" options={DIRECTIONS} value={details.direction} disabled={disabled} onChange={(direction) => onChange({ direction })} />
      </div>
      <Field label="Description" data-testid="details-step-description-field">
        <Textarea
          value={details.description}
          resize="vertical"
          disabled={disabled}
          onChange={(_, data) => onChange({ description: data.value })}
          data-testid="details-step-description"
        />
      </Field>
      <Checkbox
        label="Also save a PDF copy"
        checked={details.pdfCopy}
        disabled={disabled}
        onChange={(_, data) => onChange({ pdfCopy: !!data.checked })}
        data-testid="details-step-pdf-copy"
      />
    </div>
  );
};
