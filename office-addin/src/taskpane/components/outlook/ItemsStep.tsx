import * as React from "react";
import { Caption1, Checkbox, makeStyles, tokens } from "@fluentui/react-components";
import { EMAIL_ROW_ID, type IFilingRow } from "../../../shared/hooks/useFileToMatter";
import { testIdPart } from "../../../shared/testId";

const useStyles = makeStyles({
  root: { display: "flex", flexDirection: "column", gap: tokens.spacingVerticalXS },
  row: { display: "flex", flexDirection: "column" },
  note: { paddingLeft: tokens.spacingHorizontalXXL, color: tokens.colorNeutralForeground3 },
});

export interface IItemsStepProps {
  rows: IFilingRow[];
  ticked: ReadonlySet<string>;
  disabled: boolean;
  onToggle: (rowId: string, on: boolean) => void;
}

const KB = 1024;
const KIND_LABELS: Record<IFilingRow["kind"], string> = { email: "Email", item: "Attached email", file: "Attachment" };

function note(row: IFilingRow): string {
  if (row.reason) return row.reason;
  return [KIND_LABELS[row.kind], row.size ? `${Math.max(1, Math.round(row.size / KB))} KB` : undefined]
    .filter(Boolean)
    .join(" · ");
}

/** Step 3: tick the email and the attachments to file. */
export const ItemsStep: React.FC<IItemsStepProps> = ({ rows, ticked, disabled, onToggle }) => {
  const styles = useStyles();
  return (
    <div className={styles.root} role="list" aria-label="Items to file" data-testid="items-step">
      {rows.map((row) => {
        const id = row.id === EMAIL_ROW_ID ? EMAIL_ROW_ID : testIdPart(row.fileName);
        return (
          <div key={row.id} role="listitem" className={styles.row} data-testid={`items-step-row-${id}`}>
            <Checkbox
              label={row.fileName}
              checked={ticked.has(row.id)}
              disabled={disabled || !row.fileable}
              onChange={(_, data) => onToggle(row.id, !!data.checked)}
              data-testid={`items-step-check-${id}`}
            />
            <Caption1 className={styles.note} data-testid={`items-step-note-${id}`}>
              {note(row)}
            </Caption1>
          </div>
        );
      })}
    </div>
  );
};
