import * as React from "react";
import { Button, Caption1, Spinner, makeStyles, tokens } from "@fluentui/react-components";

const useStyles = makeStyles({
  root: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalS,
    paddingTop: tokens.spacingVerticalS,
    borderTop: `${tokens.strokeWidthThin} solid ${tokens.colorNeutralStroke2}`,
  },
  buttons: { display: "flex", gap: tokens.spacingHorizontalS },
  primary: { flexGrow: 1 },
});

export interface IPaneFooterProps {
  /** data-testid prefix of the pane ("file-to-matter" → file-to-matter-back / -primary). */
  testId: string;
  summary?: string;
  primaryLabel: string;
  busyLabel: string;
  busy: boolean;
  canSubmit: boolean;
  onBack: () => void;
  onPrimary: () => void;
}

/** Back + one primary action for the Outlook panes. */
export const PaneFooter: React.FC<IPaneFooterProps> = (props) => {
  const { testId, summary, primaryLabel, busyLabel, busy, canSubmit, onBack, onPrimary } = props;
  const styles = useStyles();
  return (
    <div className={styles.root} data-testid={`${testId}-footer`}>
      {summary && <Caption1 data-testid={`${testId}-summary`}>{summary}</Caption1>}
      <div className={styles.buttons}>
        <Button disabled={busy} onClick={onBack} data-testid={`${testId}-back`}>
          Back
        </Button>
        <Button
          appearance="primary"
          className={styles.primary}
          disabled={!canSubmit || busy}
          icon={busy ? <Spinner size="tiny" data-testid={`${testId}-primary-spinner`} /> : undefined}
          onClick={onPrimary}
          data-testid={`${testId}-primary`}
        >
          {busy ? busyLabel : primaryLabel}
        </Button>
      </div>
    </div>
  );
};
