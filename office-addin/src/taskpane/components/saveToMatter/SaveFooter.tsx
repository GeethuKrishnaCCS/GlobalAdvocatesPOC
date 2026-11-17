import * as React from "react";
import { Button, Caption1, Spinner, makeStyles, tokens } from "@fluentui/react-components";

const useStyles = makeStyles({
  root: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalS,
    paddingTop: tokens.spacingVerticalS,
    borderTop: `${tokens.strokeWidthThin} solid ${tokens.colorNeutralStroke2}`,
    // Always visible: sticks to the bottom of the pane while the steps above scroll underneath.
    position: "sticky",
    bottom: 0,
    zIndex: 1,
    paddingBottom: tokens.spacingVerticalM,
    backgroundColor: tokens.colorNeutralBackground1,
  },
  buttons: { display: "flex", gap: tokens.spacingHorizontalS },
  primary: { flexGrow: 1 },
});

export interface ISaveFooterProps {
  /** "folder" confirms the open folder; "details" submits. */
  mode: "folder" | "details";
  summary: string;
  busy: boolean;
  isSaving: boolean;
  canSubmit: boolean;
  onBack: () => void;
  onConfirmFolder: () => void;
  onSubmit: () => void;
}

/** Back + primary action, with the "name → location" summary on the Details step. */
export const SaveFooter: React.FC<ISaveFooterProps> = (props) => {
  const { mode, summary, busy, isSaving, canSubmit, onBack, onConfirmFolder, onSubmit } = props;
  const styles = useStyles();
  return (
    <div className={styles.root} data-testid="save-to-matter-footer">
      {mode === "details" && <Caption1 data-testid="save-to-matter-summary">{summary}</Caption1>}
      <div className={styles.buttons}>
        <Button disabled={busy} onClick={onBack} data-testid="save-to-matter-back">
          Back
        </Button>
        {mode === "folder" ? (
          <Button appearance="primary" className={styles.primary} disabled={busy} onClick={onConfirmFolder} data-testid="save-to-matter-next">
            Save here
          </Button>
        ) : (
          <Button
            appearance="primary"
            className={styles.primary}
            disabled={!canSubmit}
            icon={isSaving ? <Spinner size="tiny" data-testid="save-to-matter-submit-spinner" /> : undefined}
            onClick={onSubmit}
            data-testid="save-to-matter-submit"
          >
            {isSaving ? "Saving…" : "Submit"}
          </Button>
        )}
      </div>
    </div>
  );
};
